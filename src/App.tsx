import { FormEvent, useEffect, useMemo, useState } from 'react';
import seedDatabase from '../database.json';
import { translations, type Language } from '../translations';
import type {
  CollarType,
  MaterialRequirement,
  MaterialStatus,
  MaterialType,
  MaterialUrgency,
  TimesheetEntry,
  UserRole,
} from '../types';

type SeedEmployee = {
  id: string;
  nameZh: string;
  nameEn: string;
  type: CollarType;
  password: string;
  role: UserRole;
};

type SeedProject = {
  id: string;
  name: string;
  budgets: Record<string, number>;
  managerId: string;
};

type Category2 = {
  id: string;
  nameZh: string;
  nameEn: string;
};

type Category3 = Category2 & {
  parentId: string;
};

type SeedDatabase = {
  employees: SeedEmployee[];
  projects: SeedProject[];
  categories2: Category2[];
  categories3: Category3[];
  entries: TimesheetEntry[];
};

type ViewKey = 'timesheet' | 'reports' | 'materials' | 'admin' | 'data';

type MaterialRecord = MaterialRequirement & {
  unit?: string;
  supplier?: string;
};

type TimesheetForm = {
  date: string;
  projectId: string;
  category2: string;
  category3: string;
  hours: string;
  remark: string;
};

type MaterialForm = {
  code: string;
  name: string;
  model: string;
  brand: string;
  specs: string;
  projectId: string;
  type: MaterialType;
  quantity: string;
  unit: string;
  urgency: MaterialUrgency;
  expectedArrival: string;
  status: MaterialStatus;
  supplier: string;
  inventoryQuantity: string;
  comments: string;
};

const database = seedDatabase as SeedDatabase;

const STORAGE_KEYS = {
  entries: 'mbc-timesheet-entries',
  materials: 'mbc-material-records',
  language: 'mbc-language',
  sessionUserId: 'mbc-session-user-id',
};

const materialStatuses: MaterialStatus[] = [
  'Draft',
  'Pending Review',
  'Approved',
  'Rejected',
  'Bidding',
  'Bidding Complete',
  'PR Pending',
  'PR Approved',
  'PR Rejected',
  'PO Issued',
  'In Production',
  'In Transit',
  'Delivered',
];

const urgencies: MaterialUrgency[] = ['Low', 'Medium', 'High', 'Urgent'];

const today = () => new Date().toISOString().slice(0, 10);

function useStoredState<T>(key: string, initialValue: T) {
  const [value, setValue] = useState<T>(() => {
    try {
      const stored = window.localStorage.getItem(key);
      return stored ? (JSON.parse(stored) as T) : initialValue;
    } catch {
      return initialValue;
    }
  });

  useEffect(() => {
    try {
      window.localStorage.setItem(key, JSON.stringify(value));
    } catch {
      // Local storage may be unavailable in some embedded browsers; the app can still run in memory.
    }
  }, [key, value]);

  return [value, setValue] as const;
}

function displayName(item: { nameZh: string; nameEn: string }, language: Language) {
  return language === 'zh' ? item.nameZh : item.nameEn;
}

function roleLabel(role: UserRole, language: Language) {
  const labels: Record<UserRole, { zh: string; en: string }> = {
    admin: { zh: '管理员', en: 'Admin' },
    member: { zh: '普通成员', en: 'Member' },
    purchaser: { zh: '采购员', en: 'Purchaser' },
    pm: { zh: '项目经理', en: 'Project Manager' },
    approver: { zh: '审批人', en: 'Approver' },
    warehouse: { zh: '仓库', en: 'Warehouse' },
  };
  return labels[role][language];
}

function collarLabel(type: CollarType, language: Language) {
  return type === 'WC'
    ? language === 'zh'
      ? '白领'
      : 'White Collar'
    : language === 'zh'
      ? '蓝领'
      : 'Blue Collar';
}

function makeId(prefix: string) {
  if ('crypto' in window && window.crypto.randomUUID) {
    return `${prefix}-${window.crypto.randomUUID()}`;
  }
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function downloadText(filename: string, content: string, type = 'text/plain;charset=utf-8') {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function toCsv(rows: Array<Record<string, string | number | undefined>>) {
  if (rows.length === 0) {
    return '';
  }
  const headers = Object.keys(rows[0]);
  const escapeCell = (value: string | number | undefined) => {
    const raw = value == null ? '' : String(value);
    return `"${raw.replace(/"/g, '""')}"`;
  };
  return [
    headers.join(','),
    ...rows.map((row) => headers.map((header) => escapeCell(row[header])).join(',')),
  ].join('\n');
}

function sumHours(entries: TimesheetEntry[]) {
  return entries.reduce((total, entry) => total + Number(entry.hours || 0), 0);
}

function isWithinRange(date: string, startDate: string, endDate: string) {
  return (!startDate || date >= startDate) && (!endDate || date <= endDate);
}

function emptyTimesheetForm(): TimesheetForm {
  const firstProject = database.projects[0]?.id ?? '';
  const firstCategory2 = database.categories2[0]?.id ?? '';
  const firstCategory3 =
    database.categories3.find((category) => category.parentId === firstCategory2)?.id ?? '';

  return {
    date: today(),
    projectId: firstProject,
    category2: firstCategory2,
    category3: firstCategory3,
    hours: '8',
    remark: '',
  };
}

function emptyMaterialForm(): MaterialForm {
  return {
    code: '',
    name: '',
    model: '',
    brand: '',
    specs: '',
    projectId: database.projects[0]?.id ?? '',
    type: 'Mechanical',
    quantity: '1',
    unit: 'pcs',
    urgency: 'Medium',
    expectedArrival: today(),
    status: 'Draft',
    supplier: '',
    inventoryQuantity: '0',
    comments: '',
  };
}

export default function App() {
  const [language, setLanguage] = useStoredState<Language>(STORAGE_KEYS.language, 'zh');
  const [sessionUserId, setSessionUserId] = useStoredState<string | null>(
    STORAGE_KEYS.sessionUserId,
    null,
  );
  const [entries, setEntries] = useStoredState<TimesheetEntry[]>(
    STORAGE_KEYS.entries,
    database.entries ?? [],
  );
  const [materials, setMaterials] = useStoredState<MaterialRecord[]>(STORAGE_KEYS.materials, []);
  const [view, setView] = useState<ViewKey>('timesheet');

  const currentUser = useMemo(
    () => database.employees.find((employee) => employee.id === sessionUserId) ?? null,
    [sessionUserId],
  );

  if (!currentUser) {
    return (
      <LoginPage
        language={language}
        onLanguageChange={setLanguage}
        onLogin={(userId) => setSessionUserId(userId)}
      />
    );
  }

  const visibleEntries =
    currentUser.role === 'admin'
      ? entries
      : entries.filter((entry) => entry.userId === currentUser.id);

  const visibleMaterials = materials;

  return (
    <div className="app-shell">
      <Sidebar
        currentUser={currentUser}
        language={language}
        view={view}
        onViewChange={setView}
        onLanguageChange={setLanguage}
        onLogout={() => setSessionUserId(null)}
      />
      <main className="content">
        <HeaderCards entries={visibleEntries} materials={visibleMaterials} language={language} />

        {view === 'timesheet' && (
          <TimesheetPage
            currentUser={currentUser}
            entries={visibleEntries}
            language={language}
            onCreate={(entry) => setEntries((items) => [entry, ...items])}
            onDelete={(id) => setEntries((items) => items.filter((entry) => entry.id !== id))}
          />
        )}

        {view === 'reports' && <ReportsPage entries={entries} language={language} />}

        {view === 'materials' && (
          <MaterialsPage
            currentUser={currentUser}
            materials={materials}
            language={language}
            onCreate={(material) => setMaterials((items) => [material, ...items])}
            onUpdate={(material) =>
              setMaterials((items) =>
                items.map((item) => (item.id === material.id ? material : item)),
              )
            }
            onDelete={(id) => setMaterials((items) => items.filter((item) => item.id !== id))}
          />
        )}

        {view === 'admin' && (
          <AdminPage
            entries={entries}
            materials={materials}
            language={language}
            currentUser={currentUser}
          />
        )}

        {view === 'data' && (
          <DataPage
            entries={entries}
            materials={materials}
            language={language}
            onImport={(payload) => {
              setEntries(payload.entries);
              setMaterials(payload.materials);
            }}
          />
        )}
      </main>
    </div>
  );
}

function LoginPage({
  language,
  onLanguageChange,
  onLogin,
}: {
  language: Language;
  onLanguageChange: (language: Language) => void;
  onLogin: (userId: string) => void;
}) {
  const t = translations[language] as Record<string, string>;
  const [employeeId, setEmployeeId] = useState(database.employees[0]?.id ?? '');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');

  function handleSubmit(event: FormEvent) {
    event.preventDefault();
    const employee = database.employees.find((item) => item.id === employeeId);
    if (!employee || employee.password !== password) {
      setError(t.loginError);
      return;
    }
    onLogin(employee.id);
  }

  return (
    <div className="login-screen">
      <div className="login-card">
        <div className="brand-mark">MBC</div>
        <p className="eyebrow">{t.enterpriseEdition}</p>
        <h1>{t.appName}</h1>
        <p className="muted">
          {language === 'zh'
            ? '部门内网使用，支持工时登记、统计报表和物料跟踪。'
            : 'Internal team tool for timesheets, reporting, and material tracking.'}
        </p>

        <form className="stack" onSubmit={handleSubmit}>
          <label>
            {t.employee}
            <select value={employeeId} onChange={(event) => setEmployeeId(event.target.value)}>
              {database.employees.map((employee) => (
                <option key={employee.id} value={employee.id}>
                  {displayName(employee, language)} - {roleLabel(employee.role, language)}
                </option>
              ))}
            </select>
          </label>

          <label>
            {t.password}
            <input
              value={password}
              onChange={(event) => setPassword(event.target.value)}
              placeholder={language === 'zh' ? '默认密码为 1' : 'Default password is 1'}
              type="password"
            />
          </label>

          {error && <div className="form-error">{error}</div>}

          <button className="primary-button" type="submit">
            {t.loginNow}
          </button>
        </form>

        <button
          className="ghost-button full-width"
          onClick={() => onLanguageChange(language === 'zh' ? 'en' : 'zh')}
          type="button"
        >
          {language === 'zh' ? 'English' : '中文'}
        </button>
      </div>
    </div>
  );
}

function Sidebar({
  currentUser,
  language,
  view,
  onViewChange,
  onLanguageChange,
  onLogout,
}: {
  currentUser: SeedEmployee;
  language: Language;
  view: ViewKey;
  onViewChange: (view: ViewKey) => void;
  onLanguageChange: (language: Language) => void;
  onLogout: () => void;
}) {
  const t = translations[language] as Record<string, string>;
  const navItems: Array<{ key: ViewKey; label: string }> = [
    { key: 'timesheet', label: t.myTimesheet },
    { key: 'reports', label: t.reports },
    { key: 'materials', label: t.procurement },
    { key: 'admin', label: t.adminPanel },
    { key: 'data', label: t.importExport },
  ];

  return (
    <aside className="sidebar">
      <div>
        <div className="sidebar-title">
          <div className="brand-mark small">MBC</div>
          <div>
            <strong>{language === 'zh' ? '工时与物料' : 'Hours & Materials'}</strong>
            <span>{t.enterpriseEdition}</span>
          </div>
        </div>

        <nav className="nav-list">
          {navItems.map((item) => (
            <button
              className={view === item.key ? 'active' : ''}
              key={item.key}
              onClick={() => onViewChange(item.key)}
              type="button"
            >
              {item.label}
            </button>
          ))}
        </nav>
      </div>

      <div className="user-card">
        <strong>{displayName(currentUser, language)}</strong>
        <span>
          {roleLabel(currentUser.role, language)} / {collarLabel(currentUser.type, language)}
        </span>
        <div className="button-row">
          <button
            className="ghost-button"
            onClick={() => onLanguageChange(language === 'zh' ? 'en' : 'zh')}
            type="button"
          >
            {language === 'zh' ? 'EN' : '中文'}
          </button>
          <button className="ghost-button" onClick={onLogout} type="button">
            {t.signOut}
          </button>
        </div>
      </div>
    </aside>
  );
}

function HeaderCards({
  entries,
  materials,
  language,
}: {
  entries: TimesheetEntry[];
  materials: MaterialRecord[];
  language: Language;
}) {
  const t = translations[language] as Record<string, string>;
  const overdueMaterials = materials.filter(
    (material) => material.expectedArrival < today() && material.status !== 'Delivered',
  ).length;

  return (
    <section className="kpi-grid">
      <KpiCard label={t.totalHours} value={sumHours(entries).toFixed(1)} />
      <KpiCard label={language === 'zh' ? '工时记录' : 'Timesheet Entries'} value={entries.length} />
      <KpiCard label={t.materialRequirements} value={materials.length} />
      <KpiCard label={t.overdueAlert} value={overdueMaterials} tone={overdueMaterials ? 'warn' : ''} />
    </section>
  );
}

function KpiCard({ label, value, tone = '' }: { label: string; value: string | number; tone?: string }) {
  return (
    <div className={`kpi-card ${tone}`}>
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function TimesheetPage({
  currentUser,
  entries,
  language,
  onCreate,
  onDelete,
}: {
  currentUser: SeedEmployee;
  entries: TimesheetEntry[];
  language: Language;
  onCreate: (entry: TimesheetEntry) => void;
  onDelete: (id: string) => void;
}) {
  const t = translations[language] as Record<string, string>;
  const [form, setForm] = useState<TimesheetForm>(emptyTimesheetForm);
  const category3Options = useMemo(
    () => database.categories3.filter((category) => category.parentId === form.category2),
    [form.category2],
  );

  useEffect(() => {
    if (category3Options.length > 0 && !category3Options.some((item) => item.id === form.category3)) {
      setForm((current) => ({ ...current, category3: category3Options[0].id }));
    }
  }, [category3Options, form.category3]);

  function handleSubmit(event: FormEvent) {
    event.preventDefault();
    const hours = Number(form.hours);
    const project = database.projects.find((item) => item.id === form.projectId);
    const category2 = database.categories2.find((item) => item.id === form.category2);
    const category3 = database.categories3.find((item) => item.id === form.category3);

    if (!project || !category2 || !category3 || !Number.isFinite(hours) || hours <= 0 || hours > 24) {
      alert(language === 'zh' ? '请检查日期、项目、分类和工时。' : 'Please check date, project, category, and hours.');
      return;
    }

    const entry: TimesheetEntry = {
      id: makeId('entry'),
      userId: currentUser.id,
      userName: displayName(currentUser, language),
      userType: currentUser.type,
      date: form.date,
      projectId: project.id,
      projectName: project.name,
      category2: displayName(category2, language),
      category3: displayName(category3, language),
      hours,
      remark: form.remark.trim(),
      createdAt: new Date().toISOString(),
      submittedAt: new Date().toISOString().slice(0, 16).replace('T', ' '),
    };

    onCreate(entry);
    setForm((current) => ({ ...current, hours: '8', remark: '' }));
  }

  return (
    <section className="page-grid">
      <div className="panel">
        <div className="panel-heading">
          <div>
            <p className="eyebrow">{t.newEntry}</p>
            <h2>{t.myTimesheet}</h2>
          </div>
          <span className="pill">{displayName(currentUser, language)}</span>
        </div>

        <form className="form-grid" onSubmit={handleSubmit}>
          <label>
            {t.date}
            <input
              value={form.date}
              onChange={(event) => setForm({ ...form, date: event.target.value })}
              type="date"
            />
          </label>
          <label>
            {t.project}
            <select
              value={form.projectId}
              onChange={(event) => setForm({ ...form, projectId: event.target.value })}
            >
              {database.projects.map((project) => (
                <option key={project.id} value={project.id}>
                  {project.name}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t.category} 2
            <select
              value={form.category2}
              onChange={(event) => setForm({ ...form, category2: event.target.value })}
            >
              {database.categories2.map((category) => (
                <option key={category.id} value={category.id}>
                  {displayName(category, language)}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t.category} 3
            <select
              value={form.category3}
              onChange={(event) => setForm({ ...form, category3: event.target.value })}
            >
              {category3Options.map((category) => (
                <option key={category.id} value={category.id}>
                  {displayName(category, language)}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t.hours}
            <input
              max="24"
              min="0.1"
              step="0.1"
              value={form.hours}
              onChange={(event) => setForm({ ...form, hours: event.target.value })}
              type="number"
            />
          </label>
          <label className="span-2">
            {t.remark}
            <input
              value={form.remark}
              onChange={(event) => setForm({ ...form, remark: event.target.value })}
              placeholder={language === 'zh' ? '可选备注' : 'Optional notes'}
            />
          </label>
          <button className="primary-button align-end" type="submit">
            {t.submit}
          </button>
        </form>
      </div>

      <div className="panel">
        <div className="panel-heading">
          <div>
            <p className="eyebrow">{t.recentEntries}</p>
            <h2>{language === 'zh' ? '工时明细' : 'Timesheet Details'}</h2>
          </div>
          <button
            className="ghost-button"
            onClick={() =>
              downloadText(
                `timesheet-${today()}.csv`,
                `\uFEFF${toCsv(
                  entries.map((entry) => ({
                    date: entry.date,
                    userName: entry.userName,
                    projectName: entry.projectName,
                    category2: entry.category2,
                    category3: entry.category3,
                    hours: entry.hours,
                    remark: entry.remark,
                  })),
                )}`,
                'text/csv;charset=utf-8',
              )
            }
            type="button"
          >
            {t.exportCSV}
          </button>
        </div>
        <DataTable
          emptyText={t.noEntries}
          headers={[t.date, t.employee, t.project, `${t.category} 2`, `${t.category} 3`, t.hours, t.actions]}
        >
          {entries.map((entry) => (
            <tr key={entry.id}>
              <td>{entry.date}</td>
              <td>{entry.userName}</td>
              <td>{entry.projectName}</td>
              <td>{entry.category2}</td>
              <td>{entry.category3}</td>
              <td>{entry.hours}</td>
              <td>
                <button
                  className="link-button danger"
                  onClick={() => {
                    if (window.confirm(t.confirmDelete)) {
                      onDelete(entry.id);
                    }
                  }}
                  type="button"
                >
                  {t.delete}
                </button>
              </td>
            </tr>
          ))}
        </DataTable>
      </div>
    </section>
  );
}

function ReportsPage({ entries, language }: { entries: TimesheetEntry[]; language: Language }) {
  const t = translations[language] as Record<string, string>;
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');

  const filtered = entries.filter((entry) => isWithinRange(entry.date, startDate, endDate));
  const byEmployee = summarize(filtered, (entry) => entry.userName);
  const byProject = summarize(filtered, (entry) => entry.projectName);
  const byCategory = summarize(filtered, (entry) => entry.category2);

  return (
    <section className="panel">
      <div className="panel-heading">
        <div>
          <p className="eyebrow">{t.viewReports}</p>
          <h2>{t.summary}</h2>
        </div>
        <div className="inline-fields">
          <label>
            {t.startDate}
            <input value={startDate} onChange={(event) => setStartDate(event.target.value)} type="date" />
          </label>
          <label>
            {t.endDate}
            <input value={endDate} onChange={(event) => setEndDate(event.target.value)} type="date" />
          </label>
          <button
            className="ghost-button"
            onClick={() =>
              downloadText(
                `report-${today()}.csv`,
                `\uFEFF${toCsv(
                  filtered.map((entry) => ({
                    date: entry.date,
                    userName: entry.userName,
                    projectName: entry.projectName,
                    category2: entry.category2,
                    category3: entry.category3,
                    hours: entry.hours,
                  })),
                )}`,
                'text/csv;charset=utf-8',
              )
            }
            type="button"
          >
            {t.exportReport}
          </button>
        </div>
      </div>

      <div className="summary-grid">
        <SummaryBlock title={t.employeeComparison} rows={byEmployee} />
        <SummaryBlock title={t.project} rows={byProject} />
        <SummaryBlock title={t.category} rows={byCategory} />
      </div>
    </section>
  );
}

function summarize(entries: TimesheetEntry[], getKey: (entry: TimesheetEntry) => string) {
  const map = new Map<string, number>();
  entries.forEach((entry) => {
    const key = getKey(entry) || '-';
    map.set(key, (map.get(key) ?? 0) + Number(entry.hours || 0));
  });
  return Array.from(map.entries())
    .map(([name, hours]) => ({ name, hours }))
    .sort((a, b) => b.hours - a.hours);
}

function SummaryBlock({ title, rows }: { title: string; rows: Array<{ name: string; hours: number }> }) {
  return (
    <div className="summary-block">
      <h3>{title}</h3>
      {rows.length === 0 ? (
        <p className="muted">No data</p>
      ) : (
        rows.map((row) => (
          <div className="summary-row" key={row.name}>
            <span>{row.name}</span>
            <strong>{row.hours.toFixed(1)}h</strong>
          </div>
        ))
      )}
    </div>
  );
}

function MaterialsPage({
  currentUser,
  materials,
  language,
  onCreate,
  onUpdate,
  onDelete,
}: {
  currentUser: SeedEmployee;
  materials: MaterialRecord[];
  language: Language;
  onCreate: (material: MaterialRecord) => void;
  onUpdate: (material: MaterialRecord) => void;
  onDelete: (id: string) => void;
}) {
  const t = translations[language] as Record<string, string>;
  const [form, setForm] = useState<MaterialForm>(emptyMaterialForm);

  function handleSubmit(event: FormEvent) {
    event.preventDefault();
    const quantity = Number(form.quantity);
    const inventoryQuantity = Number(form.inventoryQuantity || 0);

    if (!form.code.trim() || !form.name.trim() || !Number.isFinite(quantity) || quantity <= 0) {
      alert(language === 'zh' ? '请填写物料编码、名称和有效数量。' : 'Please enter code, name, and a valid quantity.');
      return;
    }

    const material: MaterialRecord = {
      id: makeId('material'),
      projectId: form.projectId,
      creatorId: currentUser.id,
      type: form.type,
      code: form.code.trim(),
      name: form.name.trim(),
      model: form.model.trim(),
      brand: form.brand.trim(),
      specs: form.specs.trim(),
      quantity,
      unit: form.unit.trim(),
      urgency: form.urgency,
      expectedArrival: form.expectedArrival,
      status: form.status,
      supplier: form.supplier.trim(),
      inventoryQuantity: Number.isFinite(inventoryQuantity) ? inventoryQuantity : 0,
      comments: form.comments.trim(),
      version: 1,
      updatedAt: new Date().toISOString(),
      createdAt: new Date().toISOString(),
    };

    onCreate(material);
    setForm(emptyMaterialForm());
  }

  return (
    <section className="page-grid">
      <div className="panel">
        <div className="panel-heading">
          <div>
            <p className="eyebrow">{t.submitRequirement}</p>
            <h2>{t.materialRequirements}</h2>
          </div>
        </div>

        <form className="form-grid" onSubmit={handleSubmit}>
          <label>
            {t.materialCode}
            <input value={form.code} onChange={(event) => setForm({ ...form, code: event.target.value })} />
          </label>
          <label>
            {t.materialName}
            <input value={form.name} onChange={(event) => setForm({ ...form, name: event.target.value })} />
          </label>
          <label>
            {t.model}
            <input value={form.model} onChange={(event) => setForm({ ...form, model: event.target.value })} />
          </label>
          <label>
            {t.brand}
            <input value={form.brand} onChange={(event) => setForm({ ...form, brand: event.target.value })} />
          </label>
          <label>
            {t.materialType}
            <select
              value={form.type}
              onChange={(event) => setForm({ ...form, type: event.target.value as MaterialType })}
            >
              <option value="Mechanical">{t.mechanical}</option>
              <option value="Electrical">{t.electrical}</option>
            </select>
          </label>
          <label>
            {t.quantity}
            <input
              min="0"
              step="1"
              type="number"
              value={form.quantity}
              onChange={(event) => setForm({ ...form, quantity: event.target.value })}
            />
          </label>
          <label>
            {t.unit}
            <input value={form.unit} onChange={(event) => setForm({ ...form, unit: event.target.value })} />
          </label>
          <label>
            {t.expectedArrival}
            <input
              type="date"
              value={form.expectedArrival}
              onChange={(event) => setForm({ ...form, expectedArrival: event.target.value })}
            />
          </label>
          <label>
            {t.urgency}
            <select
              value={form.urgency}
              onChange={(event) => setForm({ ...form, urgency: event.target.value as MaterialUrgency })}
            >
              {urgencies.map((urgency) => (
                <option key={urgency} value={urgency}>
                  {urgency}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t.status}
            <select
              value={form.status}
              onChange={(event) => setForm({ ...form, status: event.target.value as MaterialStatus })}
            >
              {materialStatuses.map((status) => (
                <option key={status} value={status}>
                  {status}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t.supplier}
            <input
              value={form.supplier}
              onChange={(event) => setForm({ ...form, supplier: event.target.value })}
            />
          </label>
          <label>
            {t.inventoryQuantity}
            <input
              min="0"
              step="1"
              type="number"
              value={form.inventoryQuantity}
              onChange={(event) => setForm({ ...form, inventoryQuantity: event.target.value })}
            />
          </label>
          <label className="span-2">
            {t.specification}
            <input value={form.specs} onChange={(event) => setForm({ ...form, specs: event.target.value })} />
          </label>
          <label className="span-2">
            {t.remark}
            <input
              value={form.comments}
              onChange={(event) => setForm({ ...form, comments: event.target.value })}
            />
          </label>
          <button className="primary-button align-end" type="submit">
            {t.submit}
          </button>
        </form>
      </div>

      <div className="panel">
        <div className="panel-heading">
          <div>
            <p className="eyebrow">{t.trackingTable}</p>
            <h2>{t.requirementList}</h2>
          </div>
          <button
            className="ghost-button"
            onClick={() =>
              downloadText(
                `materials-${today()}.csv`,
                `\uFEFF${toCsv(
                  materials.map((material) => ({
                    code: material.code,
                    name: material.name,
                    model: material.model,
                    brand: material.brand,
                    quantity: material.quantity,
                    unit: material.unit,
                    urgency: material.urgency,
                    expectedArrival: material.expectedArrival,
                    status: material.status,
                    supplier: material.supplier,
                    inventoryQuantity: material.inventoryQuantity,
                    comments: material.comments,
                  })),
                )}`,
                'text/csv;charset=utf-8',
              )
            }
            type="button"
          >
            {t.exportCSV}
          </button>
        </div>
        <DataTable
          emptyText={language === 'zh' ? '暂无物料需求。' : 'No material requirements yet.'}
          headers={[t.materialCode, t.materialName, t.quantity, t.expectedArrival, t.status, t.inventory, t.actions]}
        >
          {materials.map((material) => {
            const overdue = material.expectedArrival < today() && material.status !== 'Delivered';
            return (
              <tr key={material.id}>
                <td>
                  <strong>{material.code}</strong>
                  <span className="cell-note">{material.model || material.brand}</span>
                </td>
                <td>{material.name}</td>
                <td>
                  {material.quantity} {material.unit}
                </td>
                <td>
                  <span className={overdue ? 'status overdue' : 'status'}>{material.expectedArrival}</span>
                </td>
                <td>
                  <select
                    value={material.status}
                    onChange={(event) =>
                      onUpdate({
                        ...material,
                        status: event.target.value as MaterialStatus,
                        version: material.version + 1,
                        updatedAt: new Date().toISOString(),
                      })
                    }
                  >
                    {materialStatuses.map((status) => (
                      <option key={status} value={status}>
                        {status}
                      </option>
                    ))}
                  </select>
                </td>
                <td>{material.inventoryQuantity ?? 0}</td>
                <td>
                  <button
                    className="link-button"
                    onClick={() =>
                      onUpdate({
                        ...material,
                        status: 'Delivered',
                        actualArrivalTime: today(),
                        inventoryQuantity: material.quantity,
                        version: material.version + 1,
                        updatedAt: new Date().toISOString(),
                      })
                    }
                    type="button"
                  >
                    {language === 'zh' ? '到货' : 'Receive'}
                  </button>
                  <button className="link-button danger" onClick={() => onDelete(material.id)} type="button">
                    {t.delete}
                  </button>
                </td>
              </tr>
            );
          })}
        </DataTable>
      </div>
    </section>
  );
}

function AdminPage({
  entries,
  materials,
  language,
  currentUser,
}: {
  entries: TimesheetEntry[];
  materials: MaterialRecord[];
  language: Language;
  currentUser: SeedEmployee;
}) {
  const t = translations[language] as Record<string, string>;

  return (
    <section className="page-grid">
      <div className="panel">
        <div className="panel-heading">
          <div>
            <p className="eyebrow">{t.adminDashboard}</p>
            <h2>{t.manageUsers}</h2>
          </div>
          <span className="pill">{roleLabel(currentUser.role, language)}</span>
        </div>
        <DataTable emptyText={t.noEntries} headers={[t.employee, t.collarType, t.systemRole]}>
          {database.employees.map((employee) => (
            <tr key={employee.id}>
              <td>{displayName(employee, language)}</td>
              <td>{collarLabel(employee.type, language)}</td>
              <td>{roleLabel(employee.role, language)}</td>
            </tr>
          ))}
        </DataTable>
      </div>

      <div className="panel">
        <div className="panel-heading">
          <div>
            <p className="eyebrow">{t.manageCategories}</p>
            <h2>{language === 'zh' ? '基础数据' : 'Master Data'}</h2>
          </div>
        </div>
        <div className="summary-grid">
          <SummaryBlock
            title={t.projects}
            rows={database.projects.map((project) => ({ name: project.name, hours: sumHours(entries.filter((entry) => entry.projectId === project.id)) }))}
          />
          <SummaryBlock
            title={t.categories}
            rows={database.categories2.map((category) => ({
              name: displayName(category, language),
              hours: sumHours(entries.filter((entry) => entry.category2 === displayName(category, language))),
            }))}
          />
          <SummaryBlock
            title={t.procurement}
            rows={[
              { name: t.materialRequirements, hours: materials.length },
              {
                name: t.overdue,
                hours: materials.filter((material) => material.expectedArrival < today() && material.status !== 'Delivered').length,
              },
            ]}
          />
        </div>
      </div>
    </section>
  );
}

function DataPage({
  entries,
  materials,
  language,
  onImport,
}: {
  entries: TimesheetEntry[];
  materials: MaterialRecord[];
  language: Language;
  onImport: (payload: { entries: TimesheetEntry[]; materials: MaterialRecord[] }) => void;
}) {
  const t = translations[language] as Record<string, string>;
  const [text, setText] = useState('');
  const backup = JSON.stringify({ entries, materials, exportedAt: new Date().toISOString() }, null, 2);

  function handleImport() {
    try {
      const payload = JSON.parse(text) as { entries?: TimesheetEntry[]; materials?: MaterialRecord[] };
      if (!Array.isArray(payload.entries) || !Array.isArray(payload.materials)) {
        throw new Error('Invalid payload');
      }
      onImport({ entries: payload.entries, materials: payload.materials });
      setText('');
      alert(language === 'zh' ? '导入完成。' : 'Import completed.');
    } catch {
      alert(language === 'zh' ? '导入失败，请检查 JSON 格式。' : 'Import failed. Please check the JSON format.');
    }
  }

  return (
    <section className="panel">
      <div className="panel-heading">
        <div>
          <p className="eyebrow">{t.importExport}</p>
          <h2>{language === 'zh' ? '本地数据备份' : 'Local Data Backup'}</h2>
        </div>
        <button
          className="primary-button"
          onClick={() => downloadText(`mbc-backup-${today()}.json`, backup, 'application/json;charset=utf-8')}
          type="button"
        >
          {language === 'zh' ? '导出 JSON' : 'Export JSON'}
        </button>
      </div>
      <p className="muted">
        {language === 'zh'
          ? '当前版本使用浏览器本地存储。更换电脑或浏览器前，请先导出备份。'
          : 'This version uses browser local storage. Export a backup before changing browser or computer.'}
      </p>
      <textarea
        className="backup-textarea"
        onChange={(event) => setText(event.target.value)}
        placeholder={language === 'zh' ? '粘贴备份 JSON 后点击导入' : 'Paste backup JSON and import'}
        value={text}
      />
      <button className="ghost-button" onClick={handleImport} type="button">
        {language === 'zh' ? '导入备份' : 'Import Backup'}
      </button>
    </section>
  );
}

function DataTable({
  headers,
  children,
  emptyText,
}: {
  headers: string[];
  children: React.ReactNode;
  emptyText: string;
}) {
  const rows = Array.isArray(children) ? children.filter(Boolean) : children ? [children] : [];

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {headers.map((header) => (
              <th key={header}>{header}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.length > 0 ? (
            rows
          ) : (
            <tr>
              <td className="empty-cell" colSpan={headers.length}>
                {emptyText}
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}
