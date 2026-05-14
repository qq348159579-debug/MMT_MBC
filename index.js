const path = require("path");
const fs = require("fs");
const express = require("express");
const cors = require("cors");
const dotenv = require("dotenv");
const sql = require("mssql");

dotenv.config({ path: path.join(__dirname, ".env") });

const PORT = Number(process.env.PORT || 3000);
const SERVER_BUILD = "2026-04-30-deleteIds-v1";

function parseSqlTarget() {
  const server = process.env.SQL_SERVER || "";
  if (server.includes("\\")) {
    const [host, instanceName] = server.split("\\");
    return { host, instanceName };
  }
  const host = process.env.SQL_HOST || server || "localhost";
  const instanceName = process.env.SQL_INSTANCE || undefined;
  return { host, instanceName };
}

function getSqlConfig() {
  const { host, instanceName } = parseSqlTarget();
  const encrypt = String(process.env.SQL_ENCRYPT || "true").toLowerCase() === "true";
  const trustServerCertificate = String(process.env.SQL_TRUST_SERVER_CERT || "true").toLowerCase() === "true";

  return {
    user: process.env.SQL_USER,
    password: process.env.SQL_PASSWORD,
    database: process.env.SQL_DATABASE || "TMS",
    server: host,
    options: {
      encrypt,
      trustServerCertificate,
      ...(instanceName ? { instanceName } : {}),
    },
    pool: {
      max: 10,
      min: 0,
      idleTimeoutMillis: 30000,
    },
  };
}

let pool;

async function getPool() {
  if (pool) return pool;
  const cfg = getSqlConfig();
  pool = await sql.connect(cfg);
  return pool;
}

async function ensureSchema() {
  const p = await getPool();
  // Single-tenant "state blob" storage for fastest migration (v0).
  // Later we can normalize tables without breaking the app by adding new endpoints.
  await p.request().query(`
IF OBJECT_ID(N'dbo.tms_state', N'U') IS NULL
BEGIN
  CREATE TABLE dbo.tms_state (
    id INT NOT NULL CONSTRAINT PK_tms_state PRIMARY KEY,
    state_json NVARCHAR(MAX) NOT NULL,
    version BIGINT NOT NULL CONSTRAINT DF_tms_state_version DEFAULT (1),
    updated_at DATETIME2(3) NOT NULL CONSTRAINT DF_tms_state_updated_at DEFAULT SYSUTCDATETIME(),
    updated_by NVARCHAR(128) NULL
  );
  INSERT INTO dbo.tms_state (id, state_json, version) VALUES (1, N'{}', 1);
END
`);

  // Migration: add version column if table existed before optimistic locking.
  await p.request().query(`
IF COL_LENGTH('dbo.tms_state', 'version') IS NULL
BEGIN
  ALTER TABLE dbo.tms_state ADD version BIGINT NOT NULL CONSTRAINT DF_tms_state_version2 DEFAULT (1);
END
`);

  await p.request().query(`
IF OBJECT_ID(N'dbo.tms_audit', N'U') IS NULL
BEGIN
  CREATE TABLE dbo.tms_audit (
    id BIGINT IDENTITY(1,1) NOT NULL CONSTRAINT PK_tms_audit PRIMARY KEY,
    action NVARCHAR(64) NOT NULL,
    details NVARCHAR(MAX) NULL,
    operator NVARCHAR(128) NULL,
    created_at DATETIME2(3) NOT NULL CONSTRAINT DF_tms_audit_created_at DEFAULT SYSUTCDATETIME()
  );
END
`);

  // Row-level store for procurement tracking table (material requirements)
  await p.request().query(`
IF OBJECT_ID(N'dbo.material_requirements', N'U') IS NULL
BEGIN
  CREATE TABLE dbo.material_requirements (
    id NVARCHAR(96) NOT NULL CONSTRAINT PK_material_requirements PRIMARY KEY,
    row_json NVARCHAR(MAX) NOT NULL,
    updated_at DATETIME2(3) NOT NULL CONSTRAINT DF_material_requirements_updated_at DEFAULT SYSUTCDATETIME(),
    updated_by NVARCHAR(128) NULL
  );
END
`);

  // Notifications (in-app notification center)
  await p.request().query(`
IF OBJECT_ID(N'dbo.tms_notifications', N'U') IS NULL
BEGIN
  CREATE TABLE dbo.tms_notifications (
    id BIGINT IDENTITY(1,1) NOT NULL CONSTRAINT PK_tms_notifications PRIMARY KEY,
    type NVARCHAR(64) NOT NULL,
    merge_key NVARCHAR(256) NOT NULL,
    title NVARCHAR(256) NOT NULL,
    summary NVARCHAR(MAX) NULL,
    payload_json NVARCHAR(MAX) NULL,
    created_at DATETIME2(3) NOT NULL CONSTRAINT DF_tms_notifications_created_at DEFAULT SYSUTCDATETIME(),
    updated_at DATETIME2(3) NOT NULL CONSTRAINT DF_tms_notifications_updated_at DEFAULT SYSUTCDATETIME(),
    last_event_at DATETIME2(3) NOT NULL CONSTRAINT DF_tms_notifications_last_event_at DEFAULT SYSUTCDATETIME(),
    last_operator NVARCHAR(128) NULL
  );
  CREATE INDEX IX_tms_notifications_merge_key ON dbo.tms_notifications(merge_key, type, last_event_at DESC);
END
`);

  await p.request().query(`
IF OBJECT_ID(N'dbo.tms_notification_events', N'U') IS NULL
BEGIN
  CREATE TABLE dbo.tms_notification_events (
    id BIGINT IDENTITY(1,1) NOT NULL CONSTRAINT PK_tms_notification_events PRIMARY KEY,
    notification_id BIGINT NOT NULL,
    event_at DATETIME2(3) NOT NULL,
    field NVARCHAR(64) NULL,
    operator NVARCHAR(128) NULL,
    reason NVARCHAR(MAX) NULL,
    before_json NVARCHAR(MAX) NULL,
    after_json NVARCHAR(MAX) NULL,
    CONSTRAINT FK_tms_notification_events_notification
      FOREIGN KEY (notification_id) REFERENCES dbo.tms_notifications(id) ON DELETE CASCADE
  );
  CREATE INDEX IX_tms_notification_events_notification_id ON dbo.tms_notification_events(notification_id, event_at DESC);
END
`);

  await p.request().query(`
IF OBJECT_ID(N'dbo.tms_notification_recipients', N'U') IS NULL
BEGIN
  CREATE TABLE dbo.tms_notification_recipients (
    notification_id BIGINT NOT NULL,
    user_id NVARCHAR(96) NOT NULL,
    read_at DATETIME2(3) NULL,
    CONSTRAINT PK_tms_notification_recipients PRIMARY KEY (notification_id, user_id),
    CONSTRAINT FK_tms_notification_recipients_notification
      FOREIGN KEY (notification_id) REFERENCES dbo.tms_notifications(id) ON DELETE CASCADE
  );
  CREATE INDEX IX_tms_notification_recipients_user_unread ON dbo.tms_notification_recipients(user_id, read_at);
END
`);
}

async function loadStateObj(p) {
  const result = await p.request().query(`SELECT state_json FROM dbo.tms_state WHERE id=1`);
  const row = result.recordset?.[0];
  try {
    return JSON.parse(row?.state_json || "{}");
  } catch {
    return {};
  }
}

function uniq(list) {
  return Array.from(new Set((list || []).map(x => String(x || "").trim()).filter(Boolean)));
}

function safeStr(x) {
  return String(x || "").trim();
}

function addDaysDelta(oldVal, newVal) {
  try {
    const a = oldVal ? new Date(oldVal) : null;
    const b = newVal ? new Date(newVal) : null;
    if (!a || !b || isNaN(a.getTime()) || isNaN(b.getTime())) return null;
    const d = Math.round((b.getTime() - a.getTime()) / 86400000);
    return d;
  } catch {
    return null;
  }
}

function buildRecipientsFromState(state, { projectId, fieldKey, operatorId }) {
  const employees = Array.isArray(state?.employees) ? state.employees : [];
  const projects = Array.isArray(state?.projects) ? state.projects : [];
  const template = Array.isArray(state?.materialTrackingTemplate) ? state.materialTrackingTemplate : [];

  const purchasers = employees.filter(e => e?.role === "purchaser" || e?.role === "admin").map(e => e.id);
  const pmId = projects.find(p => String(p?.id || "") === String(projectId || ""))?.managerId;

  // Column assignees (per-column responsibility). Fallback when empty handled by caller.
  const tplKey = fieldKey === "expectedArrival" ? "deliveryTime" : fieldKey; // expectedArrival shares ETA responsibility
  const assignees = template.find(c => String(c?.key || "") === tplKey)?.assigneeIds || [];

  return uniq([operatorId, pmId, ...purchasers, ...(assignees || [])]);
}

async function upsertMergedNotification(p, {
  type,
  mergeKey,
  title,
  summary,
  payload,
  recipients,
  event,
  operator,
}) {
  const now = new Date();
  const eventAt = event?.eventAt ? new Date(event.eventAt) : now;
  const fiveMinAgo = new Date(eventAt.getTime() - 5 * 60 * 1000);

  const existing = await p.request()
    .input("type", sql.NVarChar(64), type)
    .input("mergeKey", sql.NVarChar(256), mergeKey)
    .input("fiveMinAgo", sql.DateTime2(3), fiveMinAgo)
    .query(`
SELECT TOP (1) id, last_event_at
FROM dbo.tms_notifications
WHERE type=@type AND merge_key=@mergeKey AND last_event_at >= @fiveMinAgo
ORDER BY last_event_at DESC
`);
  const found = existing.recordset?.[0];

  const payloadJson = payload ? JSON.stringify(payload) : null;
  const recips = uniq(recipients);

  let notificationId;
  if (found?.id) {
    notificationId = Number(found.id);
    await p.request()
      .input("id", sql.BigInt, notificationId)
      .input("title", sql.NVarChar(256), title)
      .input("summary", sql.NVarChar(sql.MAX), summary || null)
      .input("payload", sql.NVarChar(sql.MAX), payloadJson)
      .input("eventAt", sql.DateTime2(3), eventAt)
      .input("operator", sql.NVarChar(128), operator || null)
      .query(`
UPDATE dbo.tms_notifications
SET title=@title,
    summary=@summary,
    payload_json=@payload,
    updated_at=SYSUTCDATETIME(),
    last_event_at=@eventAt,
    last_operator=@operator
WHERE id=@id
`);
  } else {
    const ins = await p.request()
      .input("type", sql.NVarChar(64), type)
      .input("mergeKey", sql.NVarChar(256), mergeKey)
      .input("title", sql.NVarChar(256), title)
      .input("summary", sql.NVarChar(sql.MAX), summary || null)
      .input("payload", sql.NVarChar(sql.MAX), payloadJson)
      .input("eventAt", sql.DateTime2(3), eventAt)
      .input("operator", sql.NVarChar(128), operator || null)
      .query(`
INSERT INTO dbo.tms_notifications (type, merge_key, title, summary, payload_json, last_event_at, last_operator)
OUTPUT INSERTED.id AS id
VALUES (@type, @mergeKey, @title, @summary, @payload, @eventAt, @operator)
`);
    notificationId = Number(ins.recordset?.[0]?.id);
  }

  // Append event (always keep trace with timestamp)
  await p.request()
    .input("nid", sql.BigInt, notificationId)
    .input("eventAt", sql.DateTime2(3), eventAt)
    .input("field", sql.NVarChar(64), event?.field || null)
    .input("operator", sql.NVarChar(128), operator || null)
    .input("reason", sql.NVarChar(sql.MAX), event?.reason || null)
    .input("before", sql.NVarChar(sql.MAX), event?.before ? JSON.stringify(event.before) : null)
    .input("after", sql.NVarChar(sql.MAX), event?.after ? JSON.stringify(event.after) : null)
    .query(`
INSERT INTO dbo.tms_notification_events (notification_id, event_at, field, operator, reason, before_json, after_json)
VALUES (@nid, @eventAt, @field, @operator, @reason, @before, @after)
`);

  // Upsert recipients and reset unread for merged updates
  if (recips.length) {
    // Insert missing recipients
    const reqQ = p.request().input("nid", sql.BigInt, notificationId);
    const values = recips.map((uid, idx) => {
      const k = `u${idx}`;
      reqQ.input(k, sql.NVarChar(96), uid);
      return `(@nid, @${k})`;
    });
    await reqQ.query(`
MERGE dbo.tms_notification_recipients AS tgt
USING (VALUES ${values.join(",")}) AS src(notification_id, user_id)
ON tgt.notification_id = src.notification_id AND tgt.user_id = src.user_id
WHEN NOT MATCHED THEN
  INSERT (notification_id, user_id, read_at) VALUES (src.notification_id, src.user_id, NULL);
`);

    // Reset unread for all recipients on every new event (strong accountability)
    const reqQ2 = p.request().input("nid", sql.BigInt, notificationId);
    const names = recips.map((uid, idx) => {
      const k = `uu${idx}`;
      reqQ2.input(k, sql.NVarChar(96), uid);
      return `@${k}`;
    });
    await reqQ2.query(`
UPDATE dbo.tms_notification_recipients
SET read_at=NULL
WHERE notification_id=@nid AND user_id IN (${names.join(",")})
`);
  }

  return notificationId;
}

const app = express();
app.use(cors());
app.use(express.json({ limit: "25mb" }));

app.get("/api/health", async (req, res) => {
  try {
    await ensureSchema();
    res.json({ ok: true, db: "ok", build: SERVER_BUILD });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e?.message || e) });
  }
});

app.get("/api/state", async (req, res) => {
  try {
    await ensureSchema();
    const p = await getPool();
    const result = await p.request().query(`SELECT state_json, version, updated_at, updated_by FROM dbo.tms_state WHERE id=1`);
    const row = result.recordset?.[0];
    const json = row?.state_json || "{}";
    res.json({
      state: JSON.parse(json),
      meta: { version: Number(row?.version || 1), updatedAt: row?.updated_at, updatedBy: row?.updated_by },
    });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.put("/api/state", async (req, res) => {
  try {
    await ensureSchema();
    const state = req.body?.state;
    const operator = req.body?.operator || null;
    const expectedVersion = req.body?.expectedVersion;
    if (state === undefined) return res.status(400).json({ error: "Missing body.state" });

    const p = await getPool();
    const json = JSON.stringify(state);
    if (expectedVersion === undefined || expectedVersion === null || expectedVersion === "") {
      return res.status(400).json({ error: "Missing body.expectedVersion" });
    }

    const ev = Number(expectedVersion);
    if (!Number.isFinite(ev)) return res.status(400).json({ error: "expectedVersion must be a number" });

    const updateResult = await p
      .request()
      .input("json", sql.NVarChar(sql.MAX), json)
      .input("by", sql.NVarChar(128), operator)
      .input("ev", sql.BigInt, ev)
      .query(`
UPDATE dbo.tms_state
SET state_json=@json, updated_at=SYSUTCDATETIME(), updated_by=@by, version = version + 1
WHERE id=1 AND version=@ev;
SELECT @@ROWCOUNT AS affected;
`);

    const affected = updateResult.recordset?.[0]?.affected || 0;
    if (!affected) {
      const cur = await p.request().query(`SELECT version, updated_at, updated_by FROM dbo.tms_state WHERE id=1`);
      const row = cur.recordset?.[0];
      return res.status(409).json({
        error: "conflict",
        message: "State has been updated by someone else. Please reload.",
        meta: { version: Number(row?.version || 1), updatedAt: row?.updated_at, updatedBy: row?.updated_by },
      });
    }

    const cur = await p.request().query(`SELECT version, updated_at, updated_by FROM dbo.tms_state WHERE id=1`);
    const row = cur.recordset?.[0];

    if (req.body?.audit) {
      await p
        .request()
        .input("action", sql.NVarChar(64), req.body.audit.action || "save_state")
        .input("details", sql.NVarChar(sql.MAX), req.body.audit.details || null)
        .input("operator", sql.NVarChar(128), operator)
        .query(`INSERT INTO dbo.tms_audit (action, details, operator) VALUES (@action, @details, @operator)`);
    }

    res.json({ ok: true, meta: { version: Number(row?.version || 1), updatedAt: row?.updated_at, updatedBy: row?.updated_by } });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.get("/api/audit", async (req, res) => {
  try {
    await ensureSchema();
    const p = await getPool();
    const top = Math.min(500, Math.max(1, Number(req.query.top || 200)));
    const result = await p.request().query(`
SELECT TOP (${top}) id, action, operator, created_at, details
FROM dbo.tms_audit
ORDER BY id DESC
`);
    res.json({ items: result.recordset || [] });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.get("/api/material-requirements", async (req, res) => {
  try {
    await ensureSchema();
    const p = await getPool();
    const result = await p.request().query(`
SELECT id, row_json, updated_at, updated_by
FROM dbo.material_requirements
ORDER BY updated_at DESC
`);
    const projectId = String(req.query.projectId || "").trim();
    const bomId = String(req.query.bomId || "").trim();
    const bomType = String(req.query.bomType || req.query.type || "").trim();

    const items = (result.recordset || [])
      .map(r => {
      let obj = {};
      try { obj = JSON.parse(r.row_json || "{}"); } catch {}
      return { ...obj, id: r.id, _meta: { updatedAt: r.updated_at, updatedBy: r.updated_by } };
      })
      .filter(it => {
        if (projectId && String(it.projectId || "") !== projectId) return false;
        if (bomType && String(it.type || "") !== bomType) return false;
        if (bomId) {
          const bid = String(it.bomId || "bom-default");
          if (bid !== bomId) return false;
        }
        return true;
      });
    res.json({ items, meta: { filtered: !!(projectId || bomId || bomType), total: (result.recordset || []).length, returned: items.length } });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ---- Notifications API (server-side) ----
app.get("/api/notifications/unread-count", async (req, res) => {
  try {
    await ensureSchema();
    const p = await getPool();
    const userId = safeStr(req.query.userId || req.headers["x-user-id"]);
    if (!userId) return res.status(400).json({ error: "Missing userId" });
    const r = await p.request().input("uid", sql.NVarChar(96), userId).query(`
SELECT COUNT(1) AS unread
FROM dbo.tms_notification_recipients
WHERE user_id=@uid AND read_at IS NULL
`);
    res.json({ unread: Number(r.recordset?.[0]?.unread || 0) || 0 });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.get("/api/notifications", async (req, res) => {
  try {
    await ensureSchema();
    const p = await getPool();
    const userId = safeStr(req.query.userId || req.headers["x-user-id"]);
    const top = Math.min(200, Math.max(1, Number(req.query.top || 200)));
    if (!userId) return res.status(400).json({ error: "Missing userId" });
    const r = await p.request().input("uid", sql.NVarChar(96), userId).query(`
SELECT TOP (${top})
  n.id, n.type, n.merge_key, n.title, n.summary, n.payload_json,
  n.created_at, n.updated_at, n.last_event_at, n.last_operator,
  r.read_at,
  (SELECT COUNT(1) FROM dbo.tms_notification_events e WHERE e.notification_id=n.id) AS event_count
FROM dbo.tms_notification_recipients r
JOIN dbo.tms_notifications n ON n.id = r.notification_id
WHERE r.user_id=@uid
ORDER BY (CASE WHEN r.read_at IS NULL THEN 0 ELSE 1 END), n.last_event_at DESC
`);
    const items = (r.recordset || []).map(x => {
      let payload = null;
      try { payload = x.payload_json ? JSON.parse(x.payload_json) : null; } catch {}
      return {
        id: Number(x.id),
        type: x.type,
        mergeKey: x.merge_key,
        title: x.title,
        summary: x.summary,
        payload,
        createdAt: x.created_at,
        updatedAt: x.updated_at,
        lastEventAt: x.last_event_at,
        lastOperator: x.last_operator,
        readAt: x.read_at,
        eventCount: Number(x.event_count || 0) || 0,
      };
    });
    res.json({ items });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.get("/api/notifications/:id", async (req, res) => {
  try {
    await ensureSchema();
    const p = await getPool();
    const userId = safeStr(req.query.userId || req.headers["x-user-id"]);
    if (!userId) return res.status(400).json({ error: "Missing userId" });
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return res.status(400).json({ error: "Invalid id" });
    // Authorization: must be a recipient
    const chk = await p.request().input("nid", sql.BigInt, id).input("uid", sql.NVarChar(96), userId).query(`
SELECT 1 AS ok FROM dbo.tms_notification_recipients WHERE notification_id=@nid AND user_id=@uid
`);
    if (!chk.recordset?.length) return res.status(403).json({ error: "forbidden" });

    const n = await p.request().input("nid", sql.BigInt, id).query(`
SELECT id, type, merge_key, title, summary, payload_json, created_at, updated_at, last_event_at, last_operator
FROM dbo.tms_notifications
WHERE id=@nid
`);
    const row = n.recordset?.[0];
    if (!row) return res.status(404).json({ error: "not_found" });
    let payload = null;
    try { payload = row.payload_json ? JSON.parse(row.payload_json) : null; } catch {}

    const e = await p.request().input("nid", sql.BigInt, id).query(`
SELECT TOP (200) id, event_at, field, operator, reason, before_json, after_json
FROM dbo.tms_notification_events
WHERE notification_id=@nid
ORDER BY event_at DESC, id DESC
`);
    const events = (e.recordset || []).map(x => {
      let before = null, after = null;
      try { before = x.before_json ? JSON.parse(x.before_json) : null; } catch {}
      try { after = x.after_json ? JSON.parse(x.after_json) : null; } catch {}
      return {
        id: Number(x.id),
        eventAt: x.event_at,
        field: x.field,
        operator: x.operator,
        reason: x.reason,
        before,
        after,
      };
    });
    res.json({
      item: {
        id: Number(row.id),
        type: row.type,
        mergeKey: row.merge_key,
        title: row.title,
        summary: row.summary,
        payload,
        createdAt: row.created_at,
        updatedAt: row.updated_at,
        lastEventAt: row.last_event_at,
        lastOperator: row.last_operator,
      },
      events,
    });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.post("/api/notifications/mark-read", async (req, res) => {
  try {
    await ensureSchema();
    const p = await getPool();
    const userId = safeStr(req.body?.userId || req.headers["x-user-id"]);
    const ids = Array.isArray(req.body?.ids) ? req.body.ids : (req.body?.id ? [req.body.id] : []);
    if (!userId) return res.status(400).json({ error: "Missing userId" });
    const clean = uniq(ids);
    if (!clean.length) return res.json({ ok: true, marked: 0 });
    const reqQ = p.request().input("uid", sql.NVarChar(96), userId);
    const names = clean.map((id, idx) => {
      const k = `id${idx}`;
      reqQ.input(k, sql.BigInt, Number(id));
      return `@${k}`;
    });
    const r = await reqQ.query(`
UPDATE dbo.tms_notification_recipients
SET read_at = SYSUTCDATETIME()
WHERE user_id=@uid AND notification_id IN (${names.join(",")});
SELECT @@ROWCOUNT AS affected;
`);
    res.json({ ok: true, marked: Number(r.recordset?.[0]?.affected || 0) || 0 });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.post("/api/notifications/mark-all-read", async (req, res) => {
  try {
    await ensureSchema();
    const p = await getPool();
    const userId = safeStr(req.body?.userId || req.headers["x-user-id"]);
    if (!userId) return res.status(400).json({ error: "Missing userId" });
    const r = await p.request().input("uid", sql.NVarChar(96), userId).query(`
UPDATE dbo.tms_notification_recipients
SET read_at = SYSUTCDATETIME()
WHERE user_id=@uid AND read_at IS NULL;
SELECT @@ROWCOUNT AS affected;
`);
    res.json({ ok: true, marked: Number(r.recordset?.[0]?.affected || 0) || 0 });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// Explicit notification for "引用生成需求完成/覆盖写入"
app.post("/api/notifications/generate-copy-done", async (req, res) => {
  try {
    await ensureSchema();
    const p = await getPool();
    const operatorId = safeStr(req.body?.operatorId || req.body?.operator || "");
    const operatorName = safeStr(req.body?.operatorName || "");
    const targetProjectId = safeStr(req.body?.targetProjectId || req.body?.projectId || "");
    const deviceModelId = safeStr(req.body?.deviceModelId || "");
    const deviceCount = Math.max(1, Math.floor(Number(req.body?.deviceCount) || 1));
    const bomListCount = Math.max(0, Math.floor(Number(req.body?.bomListCount) || 0));
    const rowCount = Math.max(0, Math.floor(Number(req.body?.rowCount) || 0));
    if (!operatorId || !targetProjectId || !deviceModelId) {
      return res.status(400).json({ error: "Missing operatorId/targetProjectId/deviceModelId" });
    }

    const state = await loadStateObj(p);
    const recips = buildRecipientsFromState(state, { projectId: targetProjectId, fieldKey: "status", operatorId });
    const projectName = state?.projects?.find?.(x => String(x?.id || "") === targetProjectId)?.name || targetProjectId;
    const deviceName = state?.deviceModels?.find?.(x => String(x?.id || "") === deviceModelId)?.name || deviceModelId;
    const who = operatorName || operatorId;

    const title = `【引用生成需求】${projectName} ← ${deviceName}（覆盖写入完成）`;
    const summary = `操作人：${who}；BOM清单：${bomListCount} 个；采购行：${rowCount} 行；台数：×${deviceCount}`;
    const mergeKey = `PROC_GENERATE_COPY_DONE::${targetProjectId}::${deviceModelId}`;

    const payload = { targetProjectId, projectName, deviceModelId, deviceName, deviceCount, bomListCount, rowCount };
    const eventAt = new Date();
    const nid = await upsertMergedNotification(p, {
      type: "PROC_GENERATE_COPY_DONE",
      mergeKey,
      title,
      summary,
      payload,
      recipients: recips,
      operator: operatorId,
      event: {
        eventAt,
        field: "generateCopy",
        reason: safeStr(req.body?.reason || "") || null,
        before: null,
        after: payload,
      },
    });

    res.json({ ok: true, id: nid });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// Clear all procurement rows in a BOM (server-side authoritative delete)
app.post("/api/material-requirements/clear", async (req, res) => {
  try {
    await ensureSchema();
    const p = await getPool();
    const projectId = String(req.body?.projectId || "").trim();
    const bomId = String(req.body?.bomId || "").trim() || "bom-default";
    const bomType = String(req.body?.bomType || req.body?.type || "").trim();
    const operator = req.body?.operator || null;
    if (!projectId || !bomType) return res.status(400).json({ error: "Missing body.projectId or body.bomType" });

    const result = await p.request().query(`SELECT id, row_json FROM dbo.material_requirements`);
    const candidates = [];
    for (const r of (result.recordset || [])) {
      let obj = {};
      try { obj = JSON.parse(r.row_json || "{}"); } catch {}
      const pid = String(obj.projectId || "");
      const typ = String(obj.type || "");
      const bid = String(obj.bomId || "bom-default");
      if (pid === projectId && typ === bomType && bid === bomId) candidates.push(String(r.id));
    }

    let deleted = 0;
    // delete in batches (to reduce round-trips)
    const batchSize = 200;
    for (let i = 0; i < candidates.length; i += batchSize) {
      const batch = candidates.slice(i, i + batchSize);
      // build parameterized IN clause
      const reqQ = p.request();
      const names = batch.map((id, idx) => {
        const k = `id${idx}`;
        reqQ.input(k, sql.NVarChar(96), id);
        return `@${k}`;
      });
      const q = `DELETE FROM dbo.material_requirements WHERE id IN (${names.join(",")}); SELECT @@ROWCOUNT AS affected;`;
      const r = await reqQ.query(q);
      deleted += Number(r.recordset?.[0]?.affected || 0) || 0;
    }

    console.log(`[material-requirements] clear: projectId=${projectId}, bomType=${bomType}, bomId=${bomId}, deleted=${deleted}, operator=${operator || ""}`);
    res.json({ ok: true, deleted, projectId, bomType, bomId });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

function deepEqualPrimitive(a, b) {
  // for our use: string/number/boolean/null/undefined
  if (a === b) return true;
  if (a === null && b === undefined) return true;
  if (a === undefined && b === null) return true;
  return false;
}

function getByPath(obj, path) {
  const parts = path.split(".");
  let cur = obj;
  for (const p of parts) {
    if (!cur || typeof cur !== "object") return undefined;
    cur = cur[p];
  }
  return cur;
}

function setByPath(obj, path, value) {
  const parts = path.split(".");
  let cur = obj;
  for (let i = 0; i < parts.length - 1; i++) {
    const k = parts[i];
    if (!cur[k] || typeof cur[k] !== "object") cur[k] = {};
    cur = cur[k];
  }
  cur[parts[parts.length - 1]] = value;
}

/**
 * Field-level 3-way merge:
 * - client provides base(old) and next(new)
 * - server compares current value vs base:
 *    - equal => safe apply next
 *    - not equal => conflict (unless forced)
 */
app.post("/api/material-requirements/patch", async (req, res) => {
  try {
    await ensureSchema();
    const p = await getPool();
    const stateObj = await loadStateObj(p);
    const patches = Array.isArray(req.body?.patches) ? req.body.patches : [];
    const deleteIds = Array.isArray(req.body?.deleteIds) ? req.body.deleteIds : [];
    const operator = req.body?.operator || null;
    if (!patches.length && !deleteIds.length) return res.json({ ok: true, applied: 0, deleted: 0, conflicts: [] });

    console.log(`[material-requirements] patch: patches=${patches.length}, deleteIds=${deleteIds.length}, operator=${operator || ""}`);

    const conflicts = [];
    let applied = 0;
    let deleted = 0;

    const deleteSet = new Set(deleteIds.map(x => String(x || "").trim()).filter(Boolean));

    // 0) Apply deletions first
    for (const idRaw of deleteSet) {
      const id = String(idRaw || "").trim();
      if (!id) continue;
      const r = await p.request().input("id", sql.NVarChar(96), id).query(`
DELETE FROM dbo.material_requirements WHERE id=@id;
SELECT @@ROWCOUNT AS affected;
`);
      const affected = r.recordset?.[0]?.affected || 0;
      if (affected) deleted += Number(affected) || 0;
    }

    // Process sequentially to keep logic simple
    for (const patch of patches) {
      const id = String(patch?.id || "").trim();
      if (!id) continue;
      if (deleteSet.has(id)) continue; // deletion wins
      const changes = patch?.changes && typeof patch.changes === "object" ? patch.changes : {};
      const forceFields = Array.isArray(patch?.forceFields) ? patch.forceFields : [];
      const upsertRow = patch?.upsertRow && typeof patch.upsertRow === "object" ? patch.upsertRow : null;

      const curResult = await p.request().input("id", sql.NVarChar(96), id).query(`
SELECT row_json FROM dbo.material_requirements WHERE id=@id
`);
      const exists = (curResult.recordset || []).length > 0;

      if (!exists) {
        if (!upsertRow) {
          // nothing to apply
          continue;
        }
        const json = JSON.stringify(upsertRow);
        await p.request()
          .input("id", sql.NVarChar(96), id)
          .input("json", sql.NVarChar(sql.MAX), json)
          .input("by", sql.NVarChar(128), operator)
          .query(`
INSERT INTO dbo.material_requirements (id, row_json, updated_at, updated_by)
VALUES (@id, @json, SYSUTCDATETIME(), @by)
`);
        applied++;
        continue;
      }

      let curObj = {};
      try { curObj = JSON.parse(curResult.recordset[0].row_json || "{}"); } catch {}
      const beforeSnapshot = JSON.parse(JSON.stringify(curObj || {}));

      let changed = false;
      const patchConflicts = [];
      const notifyEvents = [];

      for (const field of Object.keys(changes)) {
        const c = changes[field] || {};
        const baseVal = c.old;
        const nextVal = c.new;
        const currentVal = getByPath(curObj, field);
        const isForced = forceFields.includes(field);
        if (isForced || deepEqualPrimitive(currentVal, baseVal)) {
          setByPath(curObj, field, nextVal);
          changed = true;

          // Candidate notification events (keep per-event timestamp)
          if (field === "status" || field === "deliveryTime" || field === "expectedArrival") {
            // record only if actually changed
            if (!deepEqualPrimitive(currentVal, nextVal)) {
              notifyEvents.push({
                field,
                oldValue: currentVal,
                newValue: nextVal,
              });
            }
          }
        } else {
          patchConflicts.push({
            id,
            field,
            serverValue: currentVal,
            baseValue: baseVal,
            yourValue: nextVal,
          });
        }
      }

      if (changed) {
        // Enforce: P1 ETA/交期修改必须填写变更原因（强制）
        const hasEtaChange = notifyEvents.some(ev => ev.field === "deliveryTime" || ev.field === "expectedArrival");
        if (hasEtaChange) {
          const reason = safeStr(curObj.changeReason);
          if (!reason) {
            return res.status(400).json({ ok: false, error: "missing_change_reason", message: "ETA/交期修改必须填写变更原因（changeReason）" });
          }
        }

        const json = JSON.stringify(curObj);
        await p.request()
          .input("id", sql.NVarChar(96), id)
          .input("json", sql.NVarChar(sql.MAX), json)
          .input("by", sql.NVarChar(128), operator)
          .query(`
UPDATE dbo.material_requirements
SET row_json=@json, updated_at=SYSUTCDATETIME(), updated_by=@by
WHERE id=@id
`);
        applied++;

        // Create notifications after persistence (server authoritative)
        for (const ev of notifyEvents) {
          const projectId = safeStr(curObj.projectId);
          const bomType = safeStr(curObj.type);
          const bomId = safeStr(curObj.bomId || "bom-default") || "bom-default";
          const model = safeStr(curObj.model || curObj.code);
          const name = safeStr(curObj.name);

          const projectName = stateObj?.projects?.find?.(x => String(x?.id || "") === projectId)?.name || projectId;
          const bomName = (stateObj?.materialBoms || []).find?.(b => String(b?.projectId || "") === projectId && String(b?.type || "") === bomType && String(b?.id || "") === bomId)?.name || bomId;

          const recipients = buildRecipientsFromState(stateObj, { projectId, fieldKey: ev.field, operatorId: safeStr(operator) });
          const eventAt = new Date();

          if (ev.field === "status") {
            const title = `【采购状态】${projectName} / ${model || id}：${safeStr(ev.oldValue) || "-"} → ${safeStr(ev.newValue) || "-"}`;
            const summary = `BOM：${bomName}（${bomType}）；物料：${model}${name ? " · " + name : ""}`;
            const mergeKey = `PROC_STATUS_MILESTONE::${projectId}::${bomType}::${bomId}::${id}::status`;
            const payload = { projectId, projectName, bomType, bomId, bomName, requirementId: id, model, name };
            await upsertMergedNotification(p, {
              type: "PROC_STATUS_MILESTONE",
              mergeKey,
              title,
              summary,
              payload,
              recipients,
              operator: safeStr(operator),
              event: {
                eventAt,
                field: "status",
                reason: safeStr(curObj.changeReason || "") || null,
                before: { status: ev.oldValue, snapshot: beforeSnapshot },
                after: { status: ev.newValue, snapshot: curObj },
              },
            });
          } else if (ev.field === "deliveryTime" || ev.field === "expectedArrival") {
            const delta = addDaysDelta(ev.oldValue, ev.newValue);
            const deltaStr = delta === null ? "" : (delta >= 0 ? `（+${delta}天）` : `（${delta}天）`);
            const fieldLabel = ev.field === "deliveryTime" ? "交付日期" : "ETA";
            const title = `【交期变更】${projectName} / ${model || id} ${fieldLabel}：${safeStr(ev.oldValue) || "-"} → ${safeStr(ev.newValue) || "-"}${deltaStr}`;
            const summary = `原因：${safeStr(curObj.changeReason)}；BOM：${bomName}（${bomType}）`;
            const mergeKey = `PROC_ETA_CHANGED::${projectId}::${bomType}::${bomId}::${id}::${ev.field}`;
            const payload = { projectId, projectName, bomType, bomId, bomName, requirementId: id, model, name, field: ev.field };
            await upsertMergedNotification(p, {
              type: "PROC_ETA_CHANGED",
              mergeKey,
              title,
              summary,
              payload,
              recipients,
              operator: safeStr(operator),
              event: {
                eventAt,
                field: ev.field,
                reason: safeStr(curObj.changeReason),
                before: { [ev.field]: ev.oldValue },
                after: { [ev.field]: ev.newValue, deltaDays: delta },
              },
            });
          }
        }
      }

      if (patchConflicts.length) conflicts.push(...patchConflicts);
    }

    if (conflicts.length) {
      return res.status(409).json({ ok: false, error: "conflict", applied, deleted, conflicts });
    }
    res.json({ ok: true, applied, deleted, conflicts: [] });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// Serve built single-file app (optional)
const distDir = path.join(__dirname, "..", "dist");
app.use("/", express.static(distDir));
app.get("/", (req, res) => {
  const file = path.join(distDir, "MBC工时与物料管理工具.html");
  if (fs.existsSync(file)) return res.sendFile(file);
  res.status(404).send("Build not found. Please run `npm run build` in the project root.");
});

app.listen(PORT, async () => {
  console.log(`[tms-server] listening on http://0.0.0.0:${PORT}`);
  try {
    await ensureSchema();
    console.log("[tms-server] db schema OK");
  } catch (e) {
    console.error("[tms-server] db schema failed:", e?.message || e);
  }
});
