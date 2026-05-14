import React, { createContext, useContext, useEffect, useState, useRef, useMemo } from 'react';
import { createPortal } from 'react-dom';
import { 
  ResponsiveContainer, 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid,
  Tooltip,
  Tooltip as RechartsTooltip, 
  PieChart as RePieChart, 
  Pie, 
  Cell,
  Legend
} from 'recharts';
import { 
  FileSpreadsheet, 
  Upload, 
  Save, 
  Languages, 
  Plus, 
  Trash2, 
  AlertCircle,
  CheckCircle2,
  Users,
  Briefcase,
  BarChart3,
  Settings as SettingsIcon,
  ChevronRight,
  Download,
  Clock,
  UserCheck,
  PieChart as LucidePieChart,
  Activity,
  ArrowRight,
  Calendar,
  Filter,
  TrendingUp,
  FileText,
  Database,
  Target,
  User,
  Pencil,
  X,
  ShieldCheck,
  RotateCcw,
  LayoutGrid,
  Copy,
  Layers,
  FolderOpen,
  MessageSquare,
  ChevronDown,
  ChevronLeft,
  Lock,
  AlertTriangle,
  Info,
  ShoppingCart,
  Package,
  ClipboardCheck,
  Bell,
  Search,
  History,
  Gavel,
  Truck,
  Warehouse,
  FileCheck,
  FileWarning,
  List
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  format, 
  startOfMonth, 
  endOfMonth, 
  eachDayOfInterval, 
  isWeekend, 
  getDaysInMonth,
  startOfWeek,
  endOfWeek,
  isSameDay,
  isSameWeek,
  isSameMonth,
  parseISO,
  addMonths,
  addWeeks,
  subWeeks,
  addDays
} from 'date-fns';
import { zhCN } from 'date-fns/locale';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import * as XLSX from 'xlsx-js-style';
import { translations } from '../translations';

import { get, set, del } from 'idb-keyval';

// --- Utilities ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

function obfuscate(data: string): Uint8Array {
  const encoder = new TextEncoder();
  const bytes = encoder.encode(data);
  for (let i = 0; i < bytes.length; i++) {
    bytes[i] = bytes[i] ^ 0x42; // XOR with key 0x42
  }
  return bytes;
}

function deobfuscate(data: Uint8Array | string): string {
  if (typeof data === 'string') {
    try {
      const shifted = data.split('').map(c => String.fromCharCode(c.charCodeAt(0) - 1)).join('');
      return decodeURIComponent(atob(shifted));
    } catch (e) {
      return data;
    }
  }

  const decrypted = new Uint8Array(data);
  for (let i = 0; i < decrypted.length; i++) {
    decrypted[i] = decrypted[i] ^ 0x42;
  }
  try {
    return new TextDecoder().decode(decrypted);
  } catch (e) {
    return new TextDecoder().decode(data);
  }
}

const getVisibleEmployees = (user: Employee | null, employees: Employee[]) => {
  if (!user) return [];
  // 超级管理类似上帝角色，不参与填报数据 (admin-1)
  const godId = 'admin-1';

  // 上帝账号永远可见全部（除 admin-1）
  if (user.id === godId) return employees.filter(e => e.id !== godId);

  // 可配置：每个员工可以设置“可见人员范围”（用于：可查看他人工时）
  // 规则：只要用户配置了 visibleMode，就优先生效（不受 role 影响）
  const mode = (user as any).visibleMode as ('self' | 'all' | 'wc' | 'bc' | 'custom' | undefined);
  const customIds = Array.isArray((user as any).visibleEmployeeIds) ? (user as any).visibleEmployeeIds as string[] : [];
  if (mode === 'all') return employees.filter(e => e.id !== godId);
  if (mode === 'wc') return employees.filter(e => e.type === 'WC' && e.id !== godId);
  if (mode === 'bc') return employees.filter(e => e.type === 'BC' && e.id !== godId);
  if (mode === 'custom') {
    const set = new Set([...customIds, user.id]);
    return employees.filter(e => set.has(e.id) && e.id !== godId);
  }

  // 管理员默认可看全部（除 admin-1）
  if (user.role === 'admin') return employees.filter(e => e.id !== godId);

  // 兼容历史写死逻辑（未配置时仍按旧规则）
  // 易红 (e-2) 看所有人
  if (user.id === 'e-2') return employees.filter(e => e.id !== godId);
  // 丁箭超 (e-3) 看所有人
  if (user.id === 'e-3') return employees.filter(e => e.id !== godId);
  // 麻义俊 (e-10) 看所有蓝领 (BC)
  if (user.id === 'e-10') return employees.filter(e => e.type === 'BC' && e.id !== godId);

  // 默认：只看自己
  return employees.filter(e => e.id === user.id);
};

const getEditableEmployees = (user: Employee | null, employees: Employee[]) => {
  if (!user) return [];
  const godId = 'admin-1';
  if (user.id === godId) return employees.filter(e => e.id !== godId);

  // 可配置：每个员工可以设置“可代填/可修改人员范围”
  // 规则：只要用户配置了 editMode，就优先生效（不受 role 影响）
  const mode = (user as any).editMode as ('self' | 'all' | 'wc' | 'bc' | 'custom' | undefined);
  const customIds = Array.isArray((user as any).editEmployeeIds) ? (user as any).editEmployeeIds as string[] : [];

  const self = new Set([user.id]);
  if (mode === 'all') return employees.filter(e => e.id !== godId);
  if (mode === 'wc') return employees.filter(e => (e.type === 'WC' || e.id === user.id) && e.id !== godId);
  if (mode === 'bc') return employees.filter(e => (e.type === 'BC' || e.id === user.id) && e.id !== godId);
  if (mode === 'custom') {
    const set = new Set([...customIds, user.id]);
    return employees.filter(e => set.has(e.id) && e.id !== godId);
  }

  // 默认：管理员可改所有；其他人仅本人
  if (user.role === 'admin') return employees.filter(e => e.id !== godId);
  return employees.filter(e => e.id === user.id);
};

function getEmpName(emp: Employee | undefined | null, lang: 'zh' | 'en') {
  if (!emp) return '-';
  return lang === 'zh' ? (emp.nameZh || (emp as any).name || '-') : (emp.nameEn || (emp as any).name || '-');
}

function getCatName(cat: any, lang: 'zh' | 'en') {
  if (!cat) return '-';
  return lang === 'zh' ? (cat.nameZh || cat.name || '-') : (cat.nameEn || cat.name || '-');
}

  const getGTXTColor = (ratio: number) => {
    if (ratio >= 90) return 'text-emerald-600 bg-emerald-50 border-emerald-100';
    if (ratio >= 70) return 'text-amber-600 bg-amber-50 border-amber-100';
    if (ratio >= 50) return 'text-orange-600 bg-orange-50 border-orange-100';
    return 'text-rose-600 bg-rose-50 border-rose-100';
  };

function getXBColor(ratio: number) {
  if (ratio >= 100) return 'text-rose-600';
  if (ratio >= 80) return 'text-orange-600';
  if (ratio >= 60) return 'text-amber-500';
  return 'text-blue-600';
}

// --- Types ---
interface Employee {
  id: string;
  nameZh: string;
  nameEn: string;
  type: 'WC' | 'BC';
  password?: string;
  role: 'admin' | 'member' | 'purchaser' | 'pm' | 'approver' | 'warehouse';
  // 权限：可填写/可查看哪些人的工时（不填=仅本人；管理员默认全可见）
  visibleMode?: 'self' | 'all' | 'wc' | 'bc' | 'custom';
  visibleEmployeeIds?: string[]; // 当 visibleMode=custom 时生效
  // 权限：可代填/可修改哪些人的工时（不填=仅本人；管理员默认全可编辑）
  editMode?: 'self' | 'all' | 'wc' | 'bc' | 'custom';
  editEmployeeIds?: string[]; // 当 editMode=custom 时生效
}

interface Allocation {
  id: string;
  nameZh: string;
  nameEn: string;
  type: 'WC' | 'BC';
  comment?: string;
}

interface Project {
  id: string;
  name: string;
  csOrder?: string; // CS Order（创建项目必填；若暂无可先用“项目号+00x”临时号，后续可修改）
  managerId?: string; // 项目经理 ID
  budgets: { [allocationId: string]: number }; // 分工级预算
  budgetComments?: { [allocationId: string]: string }; // 分工预算备注
  comment?: string; // 项目备注
  visible?: boolean; // 是否在看板显示
  sortOrder?: number; // 排序权重
  status?: 'active' | 'archived'; // 项目状态
}

interface Category2 {
  id: string;
  nameZh: string;
  nameEn: string;
}

interface Category3 {
  id: string;
  parentId: string; // Category2 ID
  nameZh: string;
  nameEn: string;
}

interface UserMapping {
  id: string;
  userId?: string; // Optional: if empty, it's a global mapping for all users
  category2Ids: string[];
  category3Ids: string[];
  allocationId: string; // The "Project Allocation" it maps to
}

interface TimesheetEntry {
  id: string;
  employeeId: string;
  projectId: string; // Category 1
  category2Id: string;
  category3Id: string;
  allocationId: string; // 记录当时的项目分工
  date: string;
  hours: number;
  comment: string;
  submittedAt?: string; // NEW: Submission timestamp
}

interface DailyAdjustment {
  id: string;
  employeeId: string;
  date: string;
  overtime: number;
  leave: number;
}

interface LogEntry {
  id: string;
  userId: string;
  userName: string;
  action: string;
  details: string;
  timestamp: string;
}

// --- Procurement Types ---

type MaterialType = 'Mechanical' | 'Electrical' | 'Standard' | 'Spare';
type MaterialUrgency = 'Low' | 'Medium' | 'High' | 'Urgent';
type MaterialStatus = 
  | 'Draft' 
  | 'Pending Review' 
  | 'Approved' 
  | 'Rejected' 
  | 'Bidding' 
  | 'Bidding Complete' 
  | 'PR Pending' 
  | 'PR Approved' 
  | 'PR Rejected'
  | 'PO Issued' 
  | 'In Production' 
  | 'In Transit' 
  | 'Delivered';

interface ApprovalStep {
  role: 'admin' | 'member' | 'purchaser' | 'pm' | 'approver' | 'warehouse';
  userId?: string;
  status: 'Pending' | 'Approved' | 'Rejected';
  comment?: string;
  timestamp?: string;
}

interface MaterialRequirement {
  id: string;
  projectId: string;
  // 同一项目下可能有多个 BOM（机械/电气/备件等），用 bomId 区分
  bomId?: string;
  // 成本归属：优先使用 CS Order；若为非项目采购，可填写成本中心（Cost Center/成本中心）
  // 规则：CS Order 与 Cost Center 至少填一个；两者都填时以 CS Order 为准
  csOrder?: string;
  costCenter?: string;
  creatorId: string;
  type: MaterialType;
  // 需求时间（项目经理给采购的 need-by date），用于缺料/延期判断（以此为基准）
  needByDate?: string;
  // 申报时间：每次点击“保存”并对该行有修改时，自动写入当天（用于追溯 PR/PO 更新延迟）
  reportedAt?: string;
  // 机械加工件：用图号+版本号做唯一识别；其他物料：用型号/零件代号+版本号
  drawingNo?: string;
  revision?: string;
  investmentOrder: string;
  code: string;
  name: string;
  model: string;
  brand: string;
  specs: string;
  quantity: number;
  unit: string;
  unitPrice: number;
  totalPrice: number;
  urgency: MaterialUrgency;
  prNumber?: string;
  prCreatedAt?: string;
  poNumber?: string;
  poCreatedAt?: string;
  supplier: string;
  actualSupplier?: string;
  prStatus?: string;
  expectedArrival: string;
  deliveryTime?: string;
  receivedAt?: string;
  warehouseStatus?: string;
  repairArrivalAt?: string;
  inboundQuantity?: number;
  outboundQuantity?: number;
  inventoryQuantity?: number;
  userName?: string;
  useDate?: string;
  qualityFeedback?: string;
  customFields?: { [key: string]: any };
  technicalDocs?: string;
  changeReason?: string;
  changeReasonNote?: string;
  sourceTemplateVersion?: string;
  status: MaterialStatus;
  comments: string;
  version: number;
  updatedAt: string;
  createdAt: string;
}

interface BiddingTask {
  id: string;
  requirementIds: string[];
  status: 'Bidding' | 'Complete';
  fileUrl?: string; // Bidding document
  suppliers: Array<{
    id: string;
    name: string;
    price: number;
    isSelected: boolean;
    comment?: string;
  }>;
  resultComment?: string;
  createdAt: string;
}

interface PurchaseRequest {
  id: string;
  biddingTaskId?: string;
  requirementIds: string[];
  supplierId: string;
  totalPrice: number;
  fileUrl?: string; // PR documents/contracts
  status: 'Pending' | 'Approved' | 'Rejected';
  approvalFlow: ApprovalStep[];
  createdAt: string;
}

interface PurchaseOrder {
  id: string;
  prId: string;
  poNumber: string;
  status: 'Issued' | 'Production' | 'Shipping' | 'Received';
  fileUrl?: string;
  supplierId: string;
  estimatedDelivery?: string;
  actualDelivery?: string;
  createdAt: string;
}

interface MaterialBOM {
  id: string;
  projectId: string;
  type: MaterialType;
  name: string;
  createdAt: string;
  // Optional: 来源母版BOM信息（用于在采购页显示“这个BOM是从哪个设备BOM版本导入的”）
  importedFromDeviceModelId?: string;
  importedFromTemplateId?: string;
  importedFromTemplateVersion?: string;
  importedFromTemplateName?: string;
  importedAt?: string;
}

type MaterialColumnDataType = 'text' | 'numeric' | 'date';
type MaterialColumnGroup = 'engineering' | 'purchasing' | 'warehouse' | 'quality' | 'custom';

interface MaterialTrackingTemplateColumn {
  id: string;
  key: string;
  titleZh: string;
  titleEn: string;
  dataType: MaterialColumnDataType;
  width: number;
  group: MaterialColumnGroup;
  assigneeIds: string[];
  enabled: boolean;
  builtIn?: boolean;
}

interface AssemblyBOMItem {
  id: string;
  projectId: string;
  equipmentId: string;
  equipmentName: string;
  bomType: MaterialType;
  partCode: string;
  partName: string;
  quantity: number;
  assemblyHours: number;
  isCritical: boolean;
  priority: number; // lower means earlier scheduling priority
  targetDate?: string;
}

interface MaterialETAPlan {
  id: string;
  projectId: string;
  partCode: string;
  etaDate: string;
  status: 'OnTrack' | 'Risky' | 'Delayed';
}

interface ProductionCapacityPlan {
  id: string;
  bomType: MaterialType;
  lineName: string;
  weeklyHours: number;
  efficiency: number; // 0~1
}

interface ProductionForecastItem {
  id: string;
  projectId: string;
  equipmentId: string;
  equipmentName: string;
  parentNodeCode?: string;
  level?: number;
  bomType: MaterialType;
  earliestStart: string;
  plannedStart: string;
  plannedEnd: string;
  requiredHours: number;
  capacityPerWeek: number;
  readiness: number;
  risk: 'Low' | 'Medium' | 'High';
  blockedParts: string[];
  assemblyStatus?: 'NotReady' | 'ReadyToAssemble' | 'Assembling' | 'Done' | 'Blocked';
  startedAt?: string;
  finishedAt?: string;
  blockerReason?: string;
  generatedAt: string;
}

interface WorkflowEvent {
  id: string;
  module: 'timesheet' | 'material' | 'forecast';
  projectId: string;
  entityType: 'assembly_node' | 'material_requirement' | 'timesheet_entry' | 'project';
  entityId: string;
  action: string;
  fromStatus?: string;
  toStatus?: string;
  message?: string;
  operatorId?: string;
  createdAt: string;
}

const getMaterialKey = (r: { model?: string; name?: string; code?: string; drawingNo?: string; revision?: string }) => {
  const drawing = String((r as any).drawingNo || '').trim();
  const revision = String((r as any).revision || '').trim();
  const base = drawing || String((r as any).model || (r as any).code || (r as any).name || '').trim();
  return base.toLowerCase() + '|' + revision.toLowerCase();
};

interface ChangeRecord {
  id: string;
  templateId: string;
  deviceModelId: string;
  bomType: MaterialType;
  templateKind?: 'assembly' | 'purchasing';
  oldVersion?: string;
  newVersion: string;
  diffType: 'new_part' | 'deleted_part' | 'qty_changed' | 'spec_changed' | 'revision_changed' | 'new_version' | 'manual_edit';
  partCode?: string;
  details: string;
  operatorId?: string;
  createdAt: string;
}

// --- Device BOM (Master BOM) ---
interface DeviceModel {
  id: string; // 设备型号编码（主键）
  name: string; // 设备型号名称/描述
  createdAt: string;
  updatedAt: string;
}

function formatDeviceModelLabel(dm: DeviceModel, all: DeviceModel[]) {
  const name = String(dm?.name || '').trim() || String(dm?.id || '').trim();
  const id = String(dm?.id || '').trim();
  const sameNameCount = (all || []).filter(x => (String((x as any)?.name || '').trim() || String((x as any)?.id || '').trim()) === name).length;
  // 默认仅显示名称；若名称重复才补充 (ID)，避免用户误以为“新名字+旧名字”
  return sameNameCount > 1 && id ? `${name} (${id})` : name;
}

interface DeviceBOMTemplate {
  id: string;
  deviceModelId: string;
  bomType: MaterialType;
  templateKind?: 'assembly' | 'purchasing'; // 机械：装配BOM(树) / 采购BOM(列表)
  version: string; // 版本号（如 V1.0 / 2026-01）
  name: string; // 显示名（可选）
  createdAt: string;
  updatedAt: string;
  // Optional metadata for traceability
  importedAt?: string; // 系统首次导入/创建时间（用于判断“后续改动”）
  sourceFileName?: string; // 来源文件名
  sourceHistoryCount?: number; // 来源Excel“历史”sheet记录条数
  sourceHistoryLastDate?: string; // 来源Excel最后改动日期
  sourceHistoryLastVersion?: string; // 来源Excel最后版本
  sourceHistoryLastBy?: string; // 来源Excel最后更改人
}

interface DeviceBOMLine {
  id: string;
  templateId: string;
  partCodeModel: string; // 零件代号/型号
  revision: string; // 版本号（空=NA）
  name: string;
  specs: string;
  quantityPerDevice: number; // 单台数量
  unit: string;
  brand: string;
  supplier: string;
  remark: string;
  // Optional fields for assembly-tree import (kept for traceability / later tree rendering)
  levelPath?: string; // e.g. 1.1.3
  parentLevelPath?: string; // e.g. 1.1
  nodeType?: string; // BOM 表结构（普通件/不可拆分件…）
  fileName?: string;
  meta?: Record<string, any>;
}

interface AppState {
  allocations: Allocation[];
  employees: Employee[];
  projects: Project[];
  categories2: Category2[];
  categories3: Category3[];
  entries: TimesheetEntry[];
  dailyAdjustments: DailyAdjustment[];
  currentUser: Employee | null;
  deletedEntries: TimesheetEntry[];
  deletedAdjustments: DailyAdjustment[];
  deletedEmployees: Employee[];
  deletedProjects: Project[];
  deletedAllocations: Allocation[];
  userMappings: UserMapping[];
  logs: LogEntry[];
  fileHandle?: FileSystemFileHandle;
  backupHandle?: FileSystemDirectoryHandle;
  backupConfig?: {
    frequency: 'daily' | 'weekly' | 'monthly' | 'none';
    lastBackup?: string;
    targetPath?: string;
  };
  userSettings?: { [userId: string]: any }; // NEW: User preferences
  materialRequirements: MaterialRequirement[];
  biddingTasks: BiddingTask[];
  purchaseRequests: PurchaseRequest[];
  purchaseOrders: PurchaseOrder[];
  materialBoms: MaterialBOM[];
  materialTrackingTemplate: MaterialTrackingTemplateColumn[];
  assemblyBoms: AssemblyBOMItem[];
  materialEtaPlans: MaterialETAPlan[];
  productionCapacities: ProductionCapacityPlan[];
  productionForecasts: ProductionForecastItem[];
  workflowEvents: WorkflowEvent[];
  deviceModels: DeviceModel[];
  deviceBomTemplates: DeviceBOMTemplate[];
  deviceBomLines: DeviceBOMLine[];
  changeRecords: ChangeRecord[];
  pageAccess?: Record<string, { mode: 'all' | 'restricted'; userIds: string[] }>; // NEW: 左侧菜单页面可见性（按人）
}

import { HotTable } from '@handsontable/react';
import Handsontable from 'handsontable';
import { registerAllModules } from 'handsontable/registry';
import 'handsontable/styles/handsontable.min.css';

// Register Handsontable modules
registerAllModules();
const CHINESE_ONLY_MODE = true;

// Normalize server base URL
const normalizeApiBase = (u: string) => String(u || '').replace(/\/+$/, '');

// 解决：某些容器使用 transform/motion 动画时，fixed 弹窗会被“裁切/偏移”
// 统一用 Portal 挂到 document.body，保证弹窗覆盖整个视口
const Portal = ({ children }: { children: React.ReactNode }) => {
  const [mounted, setMounted] = useState(false);
  useEffect(() => setMounted(true), []);
  if (!mounted) return null;
  return createPortal(children, document.body);
};

type RoleMappingRow = {
  employeeName: string; // 对应 Employee.nameEn 或 nameZh
  category2: string; // 对应 Category2.nameEn 或 nameZh
  category3List: string[]; // 对应 Category3.nameEn 或 nameZh（可能多个）
  allocationName: string; // 对应 Allocation.nameZh 或 nameEn
  note?: string;
};

// 来自你确认后的“员工多角色_项目分工映射_按指定格式_v3.xlsx”
// 用于：一键写入系统 userMappings（每个人打开都能看到/生效）
const EMBEDDED_ROLE_MAPPING_ROWS = [
  {
    "employeeName": "Darran Chen",
    "category2": "Project management",
    "category3List": [
      "Procedure Monitor and Control",
      "Plan & task definition",
      "Meeting",
      "Procurement",
      "Project Meeting",
      "Shopfloor Meeting",
      "Coordinate",
      "Quotation Management"
    ],
    "allocationName": "项目经理",
    "note": "岗位/人工指定；系统暂无该分工，建议新增 Allocation=项目经理(WC)"
  },
  {
    "employeeName": "Darran Chen",
    "category2": "Qualification and validation",
    "category3List": [
      "Documentation & Report"
    ],
    "allocationName": "确认和验证",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Darran Chen",
    "category2": "Supply chain management",
    "category3List": [
      "Supplier Management",
      "Quotation Management",
      "Incoming material check & support",
      "Procurement"
    ],
    "allocationName": "项目经理",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Darran Chen",
    "category2": "Assembly",
    "category3List": [
      "Meeting"
    ],
    "allocationName": "生产经理管理",
    "note": "按Assembly(机械)推断"
  },
  {
    "employeeName": "Emma Ruan",
    "category2": "Qualification and validation",
    "category3List": [
      "Documentation & Report",
      "Coordinate",
      "Meeting",
      "Project Meeting"
    ],
    "allocationName": "确认和验证",
    "note": "按Q&V推断"
  },
  {
    "employeeName": "Emma Ruan",
    "category2": "Project management",
    "category3List": [
      "Coordinate",
      "Meeting"
    ],
    "allocationName": "确认和验证",
    "note": "按Project management推断（建议新增分工）；系统暂无该分工，建议新增 Allocation=项目经理(WC)"
  },
  {
    "employeeName": "Gangfei Chen",
    "category2": "Assembly",
    "category3List": [
      "Assembly mechanical",
      "Shopfloor Meeting",
      "Project Meeting"
    ],
    "allocationName": "电气装配",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Jianchao Ding",
    "category2": "Project management",
    "category3List": [
      "Documentation & Report",
      "Project Meeting",
      "Plan & task definition",
      "Project-related Training",
      "Procedure Monitor and Control",
      "Coordinate",
      "Material quality issue solve"
    ],
    "allocationName": "项目经理",
    "note": "岗位/人工指定；系统暂无该分工，建议新增 Allocation=项目经理(WC)"
  },
  {
    "employeeName": "Jianchao Ding",
    "category2": "Administration",
    "category3List": [
      "Paid Leave",
      "Part management",
      "Meeting",
      "Training",
      "Documentation & Report",
      "Factory activities"
    ],
    "allocationName": "技术与支持",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Jianchao Ding",
    "category2": "Engineering",
    "category3List": [
      "Ticket System Administration",
      "Incident Management",
      "Documentation & Report"
    ],
    "allocationName": "技术与支持",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Jianchao Ding",
    "category2": "Assembly",
    "category3List": [
      "Assembly mechanical"
    ],
    "allocationName": "技术与支持",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Jianchao Ding",
    "category2": "Design",
    "category3List": [
      "Design task"
    ],
    "allocationName": "技术与支持",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Jianhui Lv",
    "category2": "Assembly",
    "category3List": [
      "Assembly electrical",
      "Commissioning"
    ],
    "allocationName": "电气设计",
    "note": "按Assembly(电气/调试)推断"
  },
  {
    "employeeName": "Jianhui Lv",
    "category2": "Design",
    "category3List": [
      "Design task"
    ],
    "allocationName": "电气设计",
    "note": "按Design默认推断（如需电气请改）"
  },
  {
    "employeeName": "Jianhui Lv",
    "category2": "Engineering",
    "category3List": [
      "Ticket System Administration",
      "Meeting",
      "Training",
      "Design task",
      "nan",
      "Incoming material check & support"
    ],
    "allocationName": "技术与支持",
    "note": "按Engineering推断"
  },
  {
    "employeeName": "Jianhui Lv",
    "category2": "Administration",
    "category3List": [
      "Paid Leave",
      "Factory activities"
    ],
    "allocationName": "技术与支持",
    "note": "按行政/培训推断（可改为管理类）"
  },
  {
    "employeeName": "Jianhui Lv",
    "category2": "Supply chain management",
    "category3List": [
      "Purchasing"
    ],
    "allocationName": "技术与支持",
    "note": "按SCM默认推断"
  },
  {
    "employeeName": "Jianning Chen",
    "category2": "Design",
    "category3List": [
      "Design task",
      "Meeting",
      "Design"
    ],
    "allocationName": "机械设计",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Jianning Chen",
    "category2": "Supply chain management",
    "category3List": [
      "Purchasing"
    ],
    "allocationName": "采购",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Jianning Chen",
    "category2": "Assembly",
    "category3List": [
      "Supplier Management",
      "Incoming material check & support"
    ],
    "allocationName": "采购",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Jianning Chen",
    "category2": "Engineering",
    "category3List": [
      "Ticket System Administration"
    ],
    "allocationName": "技术与支持",
    "note": "按Engineering推断"
  },
  {
    "employeeName": "Junxiong Zhu",
    "category2": "Assembly",
    "category3List": [
      "Assembly mechanical",
      "Shopfloor Meeting"
    ],
    "allocationName": "机械装配",
    "note": "按Assembly(机械)推断"
  },
  {
    "employeeName": "Licui Lou",
    "category2": "Assembly",
    "category3List": [
      "Assembly mechanical",
      "Shopfloor Meeting",
      "Incoming material check & support",
      "Documentation & Report"
    ],
    "allocationName": "机械装配",
    "note": "按Assembly(机械)推断"
  },
  {
    "employeeName": "Qilu Liu",
    "category2": "Assembly",
    "category3List": [
      "Assembly mechanical",
      "Shopfloor Meeting"
    ],
    "allocationName": "电气装配",
    "note": "按Assembly(机械)推断"
  },
  {
    "employeeName": "Weijun Ye",
    "category2": "Design",
    "category3List": [
      "Design task",
      "Meeting"
    ],
    "allocationName": "机械设计",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Weijun Ye",
    "category2": "Engineering",
    "category3List": [
      "Design task",
      "Supplier Management",
      "Meeting",
      "Incoming material check & support"
    ],
    "allocationName": "机械设计",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Weijun Ye",
    "category2": "Supply chain management",
    "category3List": [
      "Purchasing"
    ],
    "allocationName": "采购",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Weijun Ye",
    "category2": "Assembly",
    "category3List": [
      "Incoming material check & support"
    ],
    "allocationName": "技术与支持",
    "note": "按Assembly(机械)推断"
  },
  {
    "employeeName": "Weiyang Ma",
    "category2": "Supply chain management",
    "category3List": [
      "Part management",
      "Import and export",
      "Shopfloor Meeting"
    ],
    "allocationName": "仓库物料管理",
    "note": "按SCM/Import&Export推断"
  },
  {
    "employeeName": "Yihong Li",
    "category2": "Supply chain management",
    "category3List": [
      "Supplier Management",
      "Purchasing",
      "Coordinate"
    ],
    "allocationName": "采购",
    "note": "岗位/人工指定"
  },
  {
    "employeeName": "Yihong Li",
    "category2": "Project management",
    "category3List": [
      "Meeting"
    ],
    "allocationName": "采购",
    "note": "岗位/人工指定；系统暂无该分工，建议新增 Allocation=项目经理(WC)"
  },
  {
    "employeeName": "Yijun Ma",
    "category2": "Assembly",
    "category3List": [
      "Assembly mechanical",
      "Incoming material check & support",
      "Material quality issue solve",
      "Shopfloor Meeting",
      "Commissioning",
      "Project Meeting",
      "Supplier Management",
      "Plan & task definition",
      "Purchasing"
    ],
    "allocationName": "装配主管",
    "note": "按Assembly(机械)推断"
  },
  {
    "employeeName": "Yijun Ma",
    "category2": "Supply chain management",
    "category3List": [
      "Supplier Management",
      "Part management"
    ],
    "allocationName": "装配主管",
    "note": "按SCM/Part management推断"
  },
  {
    "employeeName": "Yueming Lu",
    "category2": "Assembly",
    "category3List": [
      "Commissioning",
      "Assembly electrical",
      "Incoming material check & support"
    ],
    "allocationName": "电气设计",
    "note": "按Assembly(电气/调试)推断"
  },
  {
    "employeeName": "Yueming Lu",
    "category2": "Engineering",
    "category3List": [
      "Ticket System Administration",
      "Meeting",
      "Project Meeting",
      "Design task"
    ],
    "allocationName": "技术与支持",
    "note": "按Engineering推断"
  },
  {
    "employeeName": "Yueming Lu",
    "category2": "Project management",
    "category3List": [
      "Project Meeting",
      "Procurement"
    ],
    "allocationName": "技术与支持",
    "note": "按Project management推断（建议新增分工）；系统暂无该分工，建议新增 Allocation=项目经理(WC)"
  },
  {
    "employeeName": "Yueming Lu",
    "category2": "Training",
    "category3List": [
      "Common Training"
    ],
    "allocationName": "技术与支持",
    "note": "按行政/培训推断（可改为管理类）"
  },
  {
    "employeeName": "Yueming Lu",
    "category2": "Design",
    "category3List": [
      "Design task"
    ],
    "allocationName": "电气设计",
    "note": "按Design默认推断（如需电气请改）"
  }
] as const satisfies ReadonlyArray<RoleMappingRow>;

const normalizeLoose = (s: any) => String(s ?? '').trim();

const buildUserMappingsFromRoleRows = (rows: ReadonlyArray<RoleMappingRow>, state: AppState) => {
  const warnings: string[] = [];
  const now = Date.now();
  const findEmployee = (name: string) =>
    state.employees.find(e => e.nameEn === name || e.nameZh === name || e.id === name);
  const findCat2 = (name: string) =>
    state.categories2.find(c => c.nameEn === name || c.nameZh === name || c.id === name);
  const findCat3 = (name: string, parentId?: string) => {
    const exactInParent = parentId
      ? state.categories3.find(c => c.parentId === parentId && (c.nameEn === name || c.nameZh === name || c.id === name))
      : undefined;
    if (exactInParent) return exactInParent;
    // fallback: unique match across all
    const matches = state.categories3.filter(c => c.nameEn === name || c.nameZh === name || c.id === name);
    if (matches.length === 1) return matches[0];
    return undefined;
  };
  const findAlloc = (name: string) =>
    state.allocations.find(a => a.nameZh === name || a.nameEn === name || a.id === name);

  const out: UserMapping[] = [];
  rows.forEach((r, idx) => {
    const emp = findEmployee(normalizeLoose(r.employeeName));
    if (!emp) {
      warnings.push(`未识别员工：${r.employeeName}`);
      return;
    }
    const cat2 = findCat2(normalizeLoose(r.category2));
    if (!cat2) {
      warnings.push(`未识别分类2：${r.employeeName} / ${r.category2}`);
      return;
    }
    const alloc = findAlloc(normalizeLoose(r.allocationName));
    if (!alloc) {
      warnings.push(`未识别项目分工：${r.employeeName} / ${r.allocationName}`);
      return;
    }

    const cat3Ids: string[] = [];
    (r.category3List || [])
      .map(x => normalizeLoose(x))
      .filter(x => x && x.toLowerCase() !== 'nan')
      .forEach(name => {
        const c3 = findCat3(name, cat2.id);
        if (c3) cat3Ids.push(c3.id);
        else warnings.push(`未识别分类3：${r.employeeName} / ${r.category2} / ${name}`);
      });
    const uniqCat3 = Array.from(new Set(cat3Ids));
    if (!uniqCat3.length) {
      warnings.push(`该行无可用分类3（已跳过）：${r.employeeName} / ${r.category2}`);
      return;
    }

    out.push({
      id: `um-imp-${now}-${idx}`,
      userId: emp.id,
      category2Ids: [cat2.id],
      category3Ids: uniqCat3,
      allocationId: alloc.id
    });
  });

  // merge duplicates: same user + cat2 + allocation
  const merged = new Map<string, UserMapping>();
  out.forEach(m => {
    const key = `${m.userId}|${m.category2Ids.join(',')}|${m.allocationId}`;
    if (!merged.has(key)) merged.set(key, { ...m, category3Ids: [...m.category3Ids] });
    else {
      const cur = merged.get(key)!;
      cur.category3Ids = Array.from(new Set([...(cur.category3Ids || []), ...(m.category3Ids || [])]));
    }
  });
  return { mappings: Array.from(merged.values()), warnings };
};

const DEFAULT_MATERIAL_TRACKING_TEMPLATE: MaterialTrackingTemplateColumn[] = [
  { id: 'mt-investmentOrder', key: 'investmentOrder', titleZh: '投资单号', titleEn: 'Investment Order', dataType: 'text', width: 120, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-csOrder', key: 'csOrder', titleZh: 'CS Order', titleEn: 'CS Order', dataType: 'text', width: 120, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-costCenter', key: 'costCenter', titleZh: '成本中心', titleEn: 'Cost Center', dataType: 'text', width: 120, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-quantity', key: 'quantity', titleZh: '数量', titleEn: 'Qty', dataType: 'numeric', width: 60, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-model', key: 'model', titleZh: '零件代号/型号', titleEn: 'Part Code/Model', dataType: 'text', width: 150, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-name', key: 'name', titleZh: '名称', titleEn: 'Name', dataType: 'text', width: 180, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-drawingNo', key: 'drawingNo', titleZh: '图号', titleEn: 'Drawing No.', dataType: 'text', width: 120, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-revision', key: 'revision', titleZh: '版本号', titleEn: 'Revision', dataType: 'text', width: 90, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-sourceTemplateVersion', key: 'sourceTemplateVersion', titleZh: '来源BOM版本', titleEn: 'Source BOM Version', dataType: 'text', width: 130, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-supplier', key: 'supplier', titleZh: '供应商', titleEn: 'Supplier', dataType: 'text', width: 120, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-prNumber', key: 'prNumber', titleZh: 'PR', titleEn: 'PR', dataType: 'text', width: 100, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-prCreatedAt', key: 'prCreatedAt', titleZh: 'PR 创建日期', titleEn: 'PR Date', dataType: 'date', width: 100, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-reportedAt', key: 'reportedAt', titleZh: '申报时间', titleEn: 'Reported At', dataType: 'date', width: 110, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-poNumber', key: 'poNumber', titleZh: 'PO', titleEn: 'PO', dataType: 'text', width: 100, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-poCreatedAt', key: 'poCreatedAt', titleZh: 'PO 创建日期', titleEn: 'PO Date', dataType: 'date', width: 100, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-actualSupplier', key: 'actualSupplier', titleZh: '实供方', titleEn: 'Actual Supplier', dataType: 'text', width: 120, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-prStatus', key: 'prStatus', titleZh: '状态', titleEn: 'State', dataType: 'text', width: 80, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-unitPrice', key: 'unitPrice', titleZh: '单价', titleEn: 'Unit Price', dataType: 'numeric', width: 80, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-totalPrice', key: 'totalPrice', titleZh: '总价', titleEn: 'Total Price', dataType: 'numeric', width: 80, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-unit', key: 'unit', titleZh: '单位', titleEn: 'Unit', dataType: 'text', width: 50, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-needByDate', key: 'needByDate', titleZh: '需求日期', titleEn: 'Need-by Date', dataType: 'date', width: 110, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-deliveryTime', key: 'deliveryTime', titleZh: '交付日期', titleEn: 'Delivery Date', dataType: 'date', width: 100, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-receivedAt', key: 'receivedAt', titleZh: '收货日期', titleEn: 'Received Date', dataType: 'date', width: 100, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-warehouseStatus', key: 'warehouseStatus', titleZh: '仓库状态', titleEn: 'Warehouse Status', dataType: 'text', width: 80, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-repairArrivalAt', key: 'repairArrivalAt', titleZh: '维修到货日期', titleEn: 'Repair Arrival', dataType: 'date', width: 100, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-inboundQuantity', key: 'inboundQuantity', titleZh: '入库数', titleEn: 'Inbound', dataType: 'numeric', width: 80, group: 'warehouse', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-outboundQuantity', key: 'outboundQuantity', titleZh: '出库数', titleEn: 'Outbound', dataType: 'numeric', width: 80, group: 'warehouse', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-inventoryQuantity', key: 'inventoryQuantity', titleZh: '库存数', titleEn: 'Inventory', dataType: 'numeric', width: 80, group: 'warehouse', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-userName', key: 'userName', titleZh: '领用人', titleEn: 'User', dataType: 'text', width: 100, group: 'warehouse', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-useDate', key: 'useDate', titleZh: '领用日期', titleEn: 'Use Date', dataType: 'date', width: 100, group: 'warehouse', assigneeIds: [], enabled: true, builtIn: true },
  { id: 'mt-qualityFeedback', key: 'qualityFeedback', titleZh: '质反馈', titleEn: 'Quality Feedback', dataType: 'text', width: 150, group: 'quality', assigneeIds: [], enabled: true, builtIn: true }
  ,{ id: 'mt-changeReason', key: 'changeReason', titleZh: '变更原因', titleEn: 'Change Reason', dataType: 'text', width: 120, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true }
  ,{ id: 'mt-changeReasonNote', key: 'changeReasonNote', titleZh: '变更说明', titleEn: 'Change Note', dataType: 'text', width: 180, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true }
];

const INITIAL_STATE: AppState = {
  allocations: [
    { id: 'alloc-1', nameZh: '机械设计', nameEn: 'Mechanical Design', type: 'WC' },
    { id: 'alloc-2', nameZh: '电气设计', nameEn: 'Electrical Design', type: 'WC' },
    { id: 'alloc-3', nameZh: '技术与支持', nameEn: 'Engineering & Support', type: 'WC' },
    { id: 'alloc-4', nameZh: '采购', nameEn: 'Procurement', type: 'WC' },
    { id: 'alloc-5', nameZh: '进出口', nameEn: 'Import & Export', type: 'WC' },
    { id: 'alloc-6', nameZh: '生产经理管理', nameEn: 'Production Management', type: 'WC' },
    { id: 'alloc-12', nameZh: '项目经理', nameEn: 'Project Manager', type: 'WC' },
    { id: 'alloc-7', nameZh: '装配主管', nameEn: 'Supervisor Assembly', type: 'WC' },
    { id: 'alloc-8', nameZh: '仓库物料管理', nameEn: 'Warehouse Management', type: 'WC' },
    { id: 'alloc-9', nameZh: '电气装配', nameEn: 'Electrical Assembly', type: 'BC' },
    { id: 'alloc-10', nameZh: '机械装配', nameEn: 'Mechanical Assembly', type: 'BC' },
    { id: 'alloc-11', nameZh: '确认和验证', nameEn: 'Qualification & Validation', type: 'WC' }
  ],
  employees: [
    { id: 'admin-1', nameZh: '管理员', nameEn: 'Admin', type: 'BC', password: '1', role: 'admin' },
    { id: 'e-1', nameZh: '刘凯斌', nameEn: 'Kaibin Liu', type: 'BC', password: '1', role: 'admin' },
    { id: 'e-2', nameZh: '李易红', nameEn: 'Yihong Li', type: 'BC', password: '1', role: 'member' },
    { id: 'e-3', nameZh: '丁箭超', nameEn: 'Jianchao Ding', type: 'BC', password: '1', role: 'admin' },
    { id: 'e-4', nameZh: '陈德龙', nameEn: 'Darran Chen', type: 'BC', password: '1', role: 'admin' },
    { id: 'e-5', nameZh: '阮微妙', nameEn: 'Emma Ruan', type: 'BC', password: '1', role: 'member' },
    { id: 'e-6', nameZh: '叶炜俊', nameEn: 'Weijun Ye', type: 'BC', password: '1', role: 'member' },
    { id: 'e-7', nameZh: '陈健宁', nameEn: 'Jianning Chen', type: 'BC', password: '1', role: 'member' },
    { id: 'e-8', nameZh: '吕建辉', nameEn: 'Jianhui Lv', type: 'BC', password: '1', role: 'member' },
    { id: 'e-9', nameZh: '卢岳明', nameEn: 'Yueming Lu', type: 'BC', password: '1', role: 'member' },
    { id: 'e-10', nameZh: '麻义俊', nameEn: 'Yijun Ma', type: 'BC', password: '1', role: 'admin' },
    { id: 'e-11', nameZh: '马伟央', nameEn: 'Weiyang Ma', type: 'BC', password: '1', role: 'member' },
    { id: 'e-12', nameZh: '陈钢飞', nameEn: 'Gangfei Chen', type: 'BC', password: '1', role: 'member' },
    { id: 'e-13', nameZh: '柳齐禄', nameEn: 'Qilu Liu', type: 'BC', password: '1', role: 'member' },
    { id: 'e-14', nameZh: '朱君雄', nameEn: 'Junxiong Zhu', type: 'BC', password: '1', role: 'member' },
    { id: 'e-15', nameZh: '楼利崔', nameEn: 'Licui Lou', type: 'BC', password: '1', role: 'member' }
  ],
  projects: [
    { id: 'p-general', name: 'General time', budgets: {}, managerId: 'admin-1' },
    { id: 'p-2024-001', name: 'Project A - Mechanical Line', budgets: {}, managerId: 'e-3', status: 'active' },
    { id: 'p-2024-002', name: 'Project B - Electrical Cabinet', budgets: {}, managerId: 'e-10', status: 'active' }
  ],
  categories2: [
    { id: 'c2-1', nameZh: '组装', nameEn: 'Assembly' },
    { id: 'c2-2', nameZh: '工程', nameEn: 'Engineering' },
    { id: 'c2-3', nameZh: '设计', nameEn: 'Design' },
    { id: 'c2-4', nameZh: '供应链管理', nameEn: 'Supply chain management' },
    { id: 'c2-5', nameZh: '资质与验证', nameEn: 'Qualification and validation' },
    { id: 'c2-6', nameZh: '项目管理', nameEn: 'Project management' },
    { id: 'c2-7', nameZh: '行政', nameEn: 'Administration' },
    { id: 'c2-8', nameZh: '培训', nameEn: 'Training' }
  ],
  categories3: [
    // Assembly
    { id: 'c3-1', parentId: 'c2-1', nameZh: '来料检查与支持', nameEn: 'Incoming material check & support' },
    { id: 'c3-2', parentId: 'c2-1', nameZh: '材料质量问题解决', nameEn: 'Material quality issue solve' },
    { id: 'c3-3', parentId: 'c2-1', nameZh: '机械组装', nameEn: 'Assembly mechanical' },
    { id: 'c3-4', parentId: 'c2-1', nameZh: '电气组装', nameEn: 'Assembly electrical' },
    { id: 'c3-5', parentId: 'c2-1', nameZh: '车间会议', nameEn: 'Shopfloor Meeting' },
    // Engineering
    { id: 'c3-6', parentId: 'c2-2', nameZh: '调试', nameEn: 'Commissioning' },
    { id: 'c3-7', parentId: 'c2-2', nameZh: '工单系统管理', nameEn: 'Ticket System Administration' },
    { id: 'c3-8', parentId: 'c2-2', nameZh: '设置', nameEn: 'Setup' },
    // Design
    { id: 'c3-9', parentId: 'c2-3', nameZh: '变更管理-行政', nameEn: 'Change Management Administrative' },
    { id: 'c3-10', parentId: 'c2-3', nameZh: '变更管理-操作', nameEn: 'Change Management Operational' },
    { id: 'c3-11', parentId: 'c2-3', nameZh: '事故管理', nameEn: 'Incident Management' },
    { id: 'c3-12', parentId: 'c2-3', nameZh: '采购', nameEn: 'Procurement' },
    { id: 'c3-13', parentId: 'c2-3', nameZh: '设计任务', nameEn: 'Design task' },
    { id: 'c3-14', parentId: 'c2-3', nameZh: '会议', nameEn: 'Meeting' },
    // Supply chain management
    { id: 'c3-15', parentId: 'c2-4', nameZh: '采购', nameEn: 'Purchasing' },
    { id: 'c3-16', parentId: 'c2-4', nameZh: '零件管理', nameEn: 'Part management' },
    { id: 'c3-17', parentId: 'c2-4', nameZh: '材料质量问题解决', nameEn: 'Material quality issue solve' },
    { id: 'c3-18', parentId: 'c2-4', nameZh: '供应商管理', nameEn: 'Supplier Management' },
    { id: 'c3-19', parentId: 'c2-4', nameZh: '进出口', nameEn: 'Import and export' },
    { id: 'c3-20', parentId: 'c2-4', nameZh: '库存', nameEn: 'Inventory' },
    // Qualification and validation
    { id: 'c3-21', parentId: 'c2-5', nameZh: '文档与报告', nameEn: 'Documentation & Report' },
    { id: 'c3-22', parentId: 'c2-5', nameZh: '会议', nameEn: 'Meeting' },
    { id: 'c3-23', parentId: 'c2-5', nameZh: '协调', nameEn: 'Coordinate' },
    // Project management
    { id: 'c3-24', parentId: 'c2-6', nameZh: '计划与任务定义', nameEn: 'Plan & task definition' },
    { id: 'c3-25', parentId: 'c2-6', nameZh: '报价管理', nameEn: 'Quotation Management' },
    { id: 'c3-26', parentId: 'c2-6', nameZh: '项目会议', nameEn: 'Project Meeting' },
    { id: 'c3-27', parentId: 'c2-6', nameZh: '程序监控与控制', nameEn: 'Procedure Monitor and Control' },
    { id: 'c3-28', parentId: 'c2-6', nameZh: '协调', nameEn: 'Coordinate' },
    // Administration
    { id: 'c3-29', parentId: 'c2-7', nameZh: '会议', nameEn: 'Meeting' },
    { id: 'c3-30', parentId: 'c2-7', nameZh: '人力资源', nameEn: 'Human Resources' },
    { id: 'c3-31', parentId: 'c2-7', nameZh: '5S, EHS', nameEn: '5S, EHS' },
    { id: 'c3-32', parentId: 'c2-7', nameZh: '现场管理', nameEn: 'Site Management' },
    { id: 'c3-33', parentId: 'c2-7', nameZh: '培训', nameEn: 'Training' },
    { id: 'c3-34', parentId: 'c2-7', nameZh: '考勤维护', nameEn: 'Attendance Maintenance' },
    { id: 'c3-35', parentId: 'c2-7', nameZh: '带薪休假', nameEn: 'Paid Leave' },
    { id: 'c3-36', parentId: 'c2-7', nameZh: '助理任务', nameEn: 'Assistant tasks' },
    { id: 'c3-37', parentId: 'c2-7', nameZh: '工厂活动', nameEn: 'Factory activities' },
    // Training
    { id: 'c3-38', parentId: 'c2-8', nameZh: '通用培训', nameEn: 'Common Training' },
    { id: 'c3-39', parentId: 'c2-8', nameZh: '项目相关培训', nameEn: 'Project-related Training' }
  ],
  entries: [],
  dailyAdjustments: [],
  currentUser: null,
  fileHandle: null,
  deletedEntries: [],
  deletedAdjustments: [],
  deletedEmployees: [],
  deletedProjects: [],
  deletedAllocations: [],
  userMappings: [],
  logs: [],
  backupConfig: {
    frequency: 'none',
    lastBackup: undefined,
    targetPath: 'C:/Backups/Timesheet'
  },
  materialRequirements: [
    {
      id: 'req-1',
      projectId: 'p1',
      creatorId: 'e-1',
      type: 'Mechanical',
      code: 'M-5501',
      name: 'Guide Rail X-Axis',
      model: 'HG20-2000',
      brand: 'HIWIN',
      specs: '2000mm, Standard Precision',
      quantity: 4,
      unit: 'pcs',
      unitPrice: 200,
      totalPrice: 800,
      supplier: 'Supplier A',
      investmentOrder: 'INV-001',
      urgency: 'High',
      expectedArrival: '2024-06-15',
      status: 'Pending Review',
      comments: 'Need urgent for assembly phase 2',
      version: 1,
      createdAt: '2024-05-10T10:00:00Z',
      updatedAt: '2024-05-10T10:00:00Z'
    },
    {
      id: 'req-2',
      projectId: 'p2',
      creatorId: 'e-2',
      type: 'Electrical',
      code: 'E-2204',
      name: 'PLC Controller S7-1500',
      model: '1515-2 PN',
      brand: 'Siemens',
      specs: 'CPU 1515-2 PN',
      quantity: 1,
      unit: 'pcs',
      unitPrice: 1500,
      totalPrice: 1500,
      supplier: 'Supplier B',
      investmentOrder: 'INV-002',
      urgency: 'Medium',
      expectedArrival: '2024-07-20',
      status: 'PO Issued',
      comments: 'Standard component',
      version: 1,
      createdAt: '2024-04-15T09:00:00Z',
      updatedAt: '2024-05-01T14:00:00Z'
    },
    {
      id: 'req-3',
      projectId: 'p1',
      creatorId: 'e-1',
      type: 'Mechanical',
      code: 'M-9901',
      name: 'Pneumatic Cylinder',
      model: 'CQ2A',
      brand: 'SMC',
      specs: 'Stroke 50mm',
      quantity: 10,
      unit: 'pcs',
      unitPrice: 80,
      totalPrice: 800,
      supplier: 'Supplier C',
      investmentOrder: 'INV-003',
      urgency: 'Low',
      expectedArrival: '2024-08-01',
      status: 'PR Approved',
      comments: 'Batch for stock',
      version: 1,
      createdAt: '2024-05-15T11:00:00Z',
      updatedAt: '2024-05-15T11:00:00Z'
    },
    {
      id: 'req-4',
      projectId: 'p2',
      creatorId: 'e-2',
      type: 'Electrical',
      code: 'E-4401',
      name: 'Inverter FR-A800',
      model: 'FR-A820-0.4K',
      brand: 'Mitsubishi',
      specs: '0.4KW, 200V',
      quantity: 2,
      unit: 'pcs',
      unitPrice: 400,
      totalPrice: 800,
      supplier: 'Supplier D',
      investmentOrder: 'INV-004',
      urgency: 'High',
      expectedArrival: '2024-09-10',
      status: 'Pending Review',
      comments: 'Need for motor control',
      version: 1,
      createdAt: '2024-06-01T10:00:00Z',
      updatedAt: '2024-06-01T10:00:00Z'
    }
  ],
  biddingTasks: [
    {
      id: 'bid-1',
      requirementIds: ['req-1'],
      status: 'Bidding',
      suppliers: [
        { id: 's1', name: 'Supplier A', price: 1200, isSelected: false },
        { id: 's2', name: 'Supplier B', price: 1150, isSelected: true, comment: 'Best price' }
      ],
      createdAt: '2024-05-12T08:00:00Z'
    }
  ],
  purchaseRequests: [],
  purchaseOrders: [
      {
        id: 'po-1',
        prId: 'pr-demo',
        poNumber: 'PO-2024-0001',
        status: 'Production',
        supplierId: 'Global Motors Ltd',
        estimatedDelivery: '2024-07-20',
        createdAt: '2024-05-15T12:00:00Z'
      }
    ],
  materialBoms: [],
  materialTrackingTemplate: DEFAULT_MATERIAL_TRACKING_TEMPLATE,
  assemblyBoms: [
    { id: 'ab-1', projectId: 'p-2024-001', equipmentId: 'EQ-A01', equipmentName: '装配设备A01', bomType: 'Mechanical', partCode: 'RAIL-HG20', partName: '导轨HG20', quantity: 4, assemblyHours: 30, isCritical: true, priority: 10, targetDate: format(addWeeks(new Date(), 6), 'yyyy-MM-dd') },
    { id: 'ab-2', projectId: 'p-2024-001', equipmentId: 'EQ-A01', equipmentName: '装配设备A01', bomType: 'Mechanical', partCode: 'SERVO-750W', partName: '伺服电机750W', quantity: 2, assemblyHours: 18, isCritical: true, priority: 10, targetDate: format(addWeeks(new Date(), 6), 'yyyy-MM-dd') },
    { id: 'ab-3', projectId: 'p-2024-002', equipmentId: 'EQ-B02', equipmentName: '装配设备B02', bomType: 'Electrical', partCode: 'PLC-1515', partName: 'PLC-1515', quantity: 1, assemblyHours: 24, isCritical: true, priority: 20, targetDate: format(addWeeks(new Date(), 8), 'yyyy-MM-dd') }
  ],
  materialEtaPlans: [
    { id: 'eta-1', projectId: 'p-2024-001', partCode: 'RAIL-HG20', etaDate: format(addWeeks(new Date(), 1), 'yyyy-MM-dd'), status: 'OnTrack' },
    { id: 'eta-2', projectId: 'p-2024-001', partCode: 'SERVO-750W', etaDate: format(addWeeks(new Date(), 2), 'yyyy-MM-dd'), status: 'Risky' },
    { id: 'eta-3', projectId: 'p-2024-002', partCode: 'PLC-1515', etaDate: format(addWeeks(new Date(), 3), 'yyyy-MM-dd'), status: 'OnTrack' }
  ],
  productionCapacities: [
    { id: 'cap-1', bomType: 'Mechanical', lineName: '机械装配线-1', weeklyHours: 320, efficiency: 0.85 },
    { id: 'cap-2', bomType: 'Electrical', lineName: '电气装配线-1', weeklyHours: 280, efficiency: 0.8 },
    { id: 'cap-3', bomType: 'Standard', lineName: '标准件预装线', weeklyHours: 160, efficiency: 0.9 }
  ],
  productionForecasts: [],
  workflowEvents: [],
  userSettings: {},
  pageAccess: {},
  deviceModels: [
    {
      id: 'DM-100',
      name: '示例设备型号（可删除）',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    }
  ],
  deviceBomTemplates: [
    {
      id: 'dbt-1',
      deviceModelId: 'DM-100',
      bomType: 'Mechanical',
      version: 'V1.0',
      name: '机械BOM V1.0',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    }
  ],
  deviceBomLines: [
    {
      id: 'dbl-1',
      templateId: 'dbt-1',
      partCodeModel: 'RAIL-HG20',
      revision: 'A',
      name: '导轨HG20',
      specs: '',
      quantityPerDevice: 4,
      unit: 'PCS',
      brand: '',
      supplier: '',
      remark: ''
    }
  ],
  changeRecords: []
  };

// --- Context ---
export type AppView =
  | 'entry'
  | 'pm_view'
  | 'setup'
  | 'page_access'
  | 'login'
  | 'register'
  | 'database'
  | 'admin_db'
  | 'procurement'
  | 'device_bom'
  | 'production_forecast'
  | 'kit_readiness';

export const AppContext = createContext<{
  state: AppState;
  setState: React.Dispatch<React.SetStateAction<AppState>>;
  view: AppView;
  language: 'zh' | 'en';
  setLanguage: (l: 'zh' | 'en') => void;
  handleExportFullSystem: () => void;
  handleBatchImportEntries: (files: FileList) => void;
  handleConnectDatabase: () => Promise<void>;
  handleSetBackupFolder: () => Promise<void>;
  loadDatabaseFromHandle: (handle: any) => Promise<void>;
  handleSaveToDatabase: (newState: AppState) => Promise<void>;
  handleLogout: () => void;
  updateUserSetting: (key: string, value: any) => void;
  setView: React.Dispatch<React.SetStateAction<AppView>>;
  addLog: (action: string, details: string) => void;
  isConnected: boolean;
  connectionType: 'none' | 'server' | 'file';
  apiBase: string;
  permissionError: boolean;
  handleReauthorize: () => Promise<void>;
} | null>(null);

// --- Data Audit View ---
function DataAuditView({ state, language }: { state: AppState, language: 'zh' | 'en' }) {
  const [selectedAnomaly, setSelectedAnomaly] = useState<any>(null);

  // 1. Check daily totals
  const dailyTotals = useMemo(() => {
    const totals: { [key: string]: { [date: string]: number } } = {};
    state.entries.forEach(e => {
      if (!totals[e.employeeId]) totals[e.employeeId] = {};
      totals[e.employeeId][e.date] = (totals[e.employeeId][e.date] || 0) + e.hours;
    });
    return totals;
  }, [state.entries]);

  const anomalies = useMemo(() => {
    const list: any[] = [];
    Object.entries(dailyTotals).forEach(([empId, dates]) => {
      const emp = state.employees.find(e => e.id === empId);
      Object.entries(dates).forEach(([date, total]) => {
        const adj = state.dailyAdjustments.find(a => a.employeeId === empId && a.date === date);
        const expected = 8 + (adj?.overtime || 0) - (adj?.leave || 0);
        if (Math.abs(total - expected) > 0.01) {
          list.push({
            type: 'hours',
            empId,
            empName: getEmpName(emp, language),
            date,
            actual: total,
            expected,
            diff: total - expected,
            adj
          });
        }
      });
    });
    return list;
  }, [dailyTotals, state.employees, state.dailyAdjustments, language]);

  // 2. Check budgets
  const budgetAnomalies = useMemo(() => {
    const list: any[] = [];
    state.projects.forEach(proj => {
      if (proj.id === 'p-general' || proj.id === 'p-leave') return;
      const projEntries = state.entries.filter(e => e.projectId === proj.id);
      const actualByAlloc: { [allocId: string]: number } = {};
      projEntries.forEach(e => {
        if (e.allocationId) {
          actualByAlloc[e.allocationId] = (actualByAlloc[e.allocationId] || 0) + e.hours;
        }
      });

      Object.entries(proj.budgets || {}).forEach(([allocId, budget]) => {
        const actual = actualByAlloc[allocId] || 0;
        if (actual > budget && budget > 0) {
          list.push({
            type: 'budget',
            projName: proj.name,
            allocName: state.allocations.find(a => a.id === allocId)?.nameZh || allocId,
            budget,
            actual
          });
        }
      });
    });
    return list;
  }, [state.projects, state.entries, state.allocations]);

  return (
    <div className="space-y-8">
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        {/* Hours Anomalies List */}
        <div className="lg:col-span-1 bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden flex flex-col h-[600px]">
          <div className="p-6 bg-slate-50 border-b border-slate-100 flex justify-between items-center">
            <h3 className="font-black text-slate-900 flex items-center gap-2">
              <AlertTriangle className="text-amber-500" size={18} />
              {language === 'zh' ? '异常列表' : 'Anomalies'}
            </h3>
            <span className="bg-amber-100 text-amber-600 px-3 py-1 rounded-full text-[10px] font-black">{anomalies.length}</span>
          </div>
          <div className="flex-1 overflow-y-auto">
            <div className="divide-y divide-slate-50">
              {anomalies.map((a, i) => (
                <button 
                  key={i} 
                  onClick={() => setSelectedAnomaly(a)}
                  className={cn(
                    "w-full text-left p-4 hover:bg-slate-50 transition-all flex justify-between items-center group",
                    selectedAnomaly === a ? "bg-blue-50 border-l-4 border-blue-600" : ""
                  )}
                >
                  <div>
                    <p className="text-xs font-black text-slate-900">{a.empName}</p>
                    <p className="text-[10px] font-bold text-slate-400">{a.date}</p>
                  </div>
                  <div className="text-right">
                    <p className={cn("text-xs font-black", a.diff > 0 ? "text-rose-600" : "text-blue-600")}>
                      {a.actual.toFixed(1)} <span className="text-[10px] text-slate-300">/</span> {a.expected.toFixed(1)}
                    </p>
                    <p className="text-[9px] font-bold text-slate-400 uppercase tracking-tighter">
                      {a.diff > 0 ? (language === 'zh' ? '多填' : 'Over') : (language === 'zh' ? '少填' : 'Under')} {Math.abs(a.diff).toFixed(1)}h
                    </p>
                  </div>
                </button>
              ))}
              {anomalies.length === 0 && (
                <div className="p-10 text-center text-slate-400 text-xs font-bold italic">
                  {language === 'zh' ? '暂无异常工时记录' : 'No hours anomalies found'}
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Detail Panel */}
        <div className="lg:col-span-2 bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden flex flex-col h-[600px]">
          {selectedAnomaly ? (
            <>
              <div className="p-6 bg-slate-50 border-b border-slate-100 flex justify-between items-center">
                <div>
                  <h3 className="font-black text-slate-900 text-lg">{selectedAnomaly.empName}</h3>
                  <p className="text-xs font-bold text-slate-400">{selectedAnomaly.date} {language === 'zh' ? '数据核对' : 'Data Verification'}</p>
                </div>
                <button 
                  onClick={() => setSelectedAnomaly(null)}
                  className="p-2 hover:bg-slate-200 rounded-full transition-colors"
                >
                  <X size={20} />
                </button>
              </div>
              <div className="flex-1 overflow-y-auto p-8 space-y-8">
                {/* 1. Raw Entries */}
                <section className="space-y-4">
                  <h4 className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2">
                    <FileText size={14} />
                    {language === 'zh' ? '原始填报记录' : 'Raw Entries'}
                  </h4>
                  <div className="bg-slate-50 rounded-2xl border border-slate-100 overflow-hidden">
                    <table className="w-full text-left text-xs">
                      <thead>
                        <tr className="bg-slate-100/50 text-slate-500">
                          <th className="px-4 py-2 font-black">{language === 'zh' ? '项目' : 'Project'}</th>
                          <th className="px-4 py-2 font-black">{language === 'zh' ? '分类/分工' : 'Category/Alloc'}</th>
                          <th className="px-4 py-2 font-black text-right">{language === 'zh' ? '工时' : 'Hours'}</th>
                          <th className="px-4 py-2 font-black">{language === 'zh' ? '备注' : 'Comment'}</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {state.entries.filter(e => e.employeeId === selectedAnomaly.empId && e.date === selectedAnomaly.date).map((e, idx) => {
                          const proj = state.projects.find(p => p.id === e.projectId);
                          const alloc = state.allocations.find(a => a.id === e.allocationId);
                          return (
                            <tr key={idx}>
                              <td className="px-4 py-3 font-bold text-slate-700">{proj?.name}</td>
                              <td className="px-4 py-3 text-slate-500">
                                {(() => {
                                  const alloc = state.allocations.find(a => a.id === e.allocationId);
                                  return language === 'zh' ? (alloc?.nameZh || '-') : (alloc?.nameEn || alloc?.nameZh || '-');
                                })()}
                              </td>
                              <td className="px-4 py-3 font-black text-right text-blue-600">{e.hours.toFixed(1)}</td>
                              <td className="px-4 py-3 text-xs text-slate-400 italic max-w-[150px] truncate" title={e.comment}>{e.comment || '-'}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                      <tfoot className="bg-blue-50/50">
                        <tr>
                          <td colSpan={3} className="px-4 py-3 font-black text-slate-900">{language === 'zh' ? '填报总计' : 'Total Filled'}</td>
                          <td className="px-4 py-3 font-black text-right text-blue-700 text-sm">{selectedAnomaly.actual.toFixed(1)}h</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                </section>

                {/* 2. Calculation Verification */}
                <section className="space-y-4">
                  <h4 className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2">
                    <Activity size={14} />
                    {language === 'zh' ? '计算逻辑核对' : 'Calculation Verification'}
                  </h4>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="p-5 bg-indigo-50 rounded-2xl border border-indigo-100">
                      <p className="text-[10px] font-black text-indigo-400 uppercase tracking-widest mb-2">{language === 'zh' ? '应填工时计算' : 'Expected Hours Calculation'}</p>
                      <div className="flex items-end gap-2">
                        <span className="text-2xl font-black text-indigo-700">8.0</span>
                        <span className="text-indigo-300 font-bold mb-1">+</span>
                        <span className="text-xl font-black text-emerald-600">{selectedAnomaly.adj?.overtime || 0}</span>
                        <span className="text-indigo-300 font-bold mb-1">-</span>
                        <span className="text-xl font-black text-rose-500">{selectedAnomaly.adj?.leave || 0}</span>
                        <span className="text-indigo-300 font-bold mb-1">=</span>
                        <span className="text-2xl font-black text-indigo-900">{selectedAnomaly.expected.toFixed(1)}</span>
                      </div>
                      <p className="text-[9px] font-bold text-indigo-400 mt-2 italic">
                        {language === 'zh' ? '(标准8h + 加班 - 请假)' : '(Standard 8h + Overtime - Leave)'}
                      </p>
                    </div>
                    <div className="p-5 bg-blue-50 rounded-2xl border border-blue-100">
                      <p className="text-[10px] font-black text-blue-400 uppercase tracking-widest mb-2">{language === 'zh' ? '效率占比结果' : 'Efficiency Results'}</p>
                      <div className="space-y-2">
                        <div className="flex justify-between items-center">
                          <span className="text-[10px] font-bold text-slate-500">{language === 'zh' ? '工投 (Efficiency)' : 'Efficiency'}</span>
                          <span className="text-xs font-black text-blue-700">{(selectedAnomaly.actual / selectedAnomaly.expected * 100).toFixed(1)}%</span>
                        </div>
                        <div className="w-full bg-blue-200 h-1 rounded-full overflow-hidden">
                          <div className="bg-blue-600 h-full" style={{ width: `${Math.min(100, (selectedAnomaly.actual / selectedAnomaly.expected * 100))}%` }} />
                        </div>
                      </div>
                    </div>
                  </div>
                </section>

                {/* 3. Project Allocation Check */}
                <section className="space-y-4">
                  <h4 className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2">
                    <Target size={14} />
                    {language === 'zh' ? '项目分配核对' : 'Project Allocation Check'}
                  </h4>
                  <div className="p-5 bg-slate-50 rounded-2xl border border-slate-100">
                    <div className="flex justify-between items-center mb-4">
                      <span className="text-xs font-bold text-slate-600">{language === 'zh' ? '项目投入占比 (Xiang Tou)' : 'Project Investment'}</span>
                      <span className="text-sm font-black text-slate-900">
                        {(() => {
                          const projectHours = state.entries
                            .filter(e => e.employeeId === selectedAnomaly.empId && e.date === selectedAnomaly.date && e.projectId !== 'p-general' && e.projectId !== 'p-leave')
                            .reduce((acc, e) => acc + e.hours, 0);
                          return ((projectHours / selectedAnomaly.actual) * 100).toFixed(1);
                        })()}%
                      </span>
                    </div>
                    <div className="space-y-2">
                      {state.entries
                        .filter(e => e.employeeId === selectedAnomaly.empId && e.date === selectedAnomaly.date)
                        .map((e, idx) => {
                          const proj = state.projects.find(p => p.id === e.projectId);
                          const percentage = (e.hours / selectedAnomaly.actual * 100).toFixed(1);
                          return (
                            <div key={idx} className="flex items-center gap-3">
                              <div className="flex-1">
                                <div className="flex justify-between text-[9px] font-bold mb-1">
                                  <span className="text-slate-500">{proj?.name}</span>
                                  <span className="text-slate-900">{percentage}%</span>
                                </div>
                                <div className="w-full bg-slate-200 h-1 rounded-full overflow-hidden">
                                  <div className="bg-indigo-500 h-full" style={{ width: `${percentage}%` }} />
                                </div>
                              </div>
                            </div>
                          );
                        })}
                    </div>
                  </div>
                </section>
              </div>
            </>
          ) : (
            <div className="flex-1 flex flex-col items-center justify-center p-12 text-center space-y-4">
              <div className="w-20 h-20 bg-slate-50 rounded-[2.5rem] flex items-center justify-center text-slate-200">
                <ShieldCheck size={40} />
              </div>
              <div>
                <h3 className="text-lg font-black text-slate-900">{language === 'zh' ? '请选择一条异常记录' : 'Select an Anomaly'}</h3>
                <p className="text-xs font-medium text-slate-400 max-w-xs mx-auto">
                  {language === 'zh' ? '点击左侧列表中的记录，即可查看详细的原始数据、计算逻辑以及项目分配占比，快速定位问题。' : 'Click a record on the left to view detailed raw data, calculation logic, and project allocation breakdown.'}
                </p>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Budget Anomalies (Simplified) */}
      <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden">
        <div className="p-6 bg-slate-50 border-b border-slate-100 flex justify-between items-center">
          <h3 className="font-black text-slate-900 flex items-center gap-2">
            <Target className="text-rose-500" size={18} />
            {language === 'zh' ? '预算超支预警' : 'Budget Overruns'}
          </h3>
          <span className="bg-rose-100 text-rose-600 px-3 py-1 rounded-full text-[10px] font-black">{budgetAnomalies.length}</span>
        </div>
        <div className="max-h-[300px] overflow-y-auto">
          <table className="w-full text-left border-collapse">
            <thead className="sticky top-0 bg-white z-10">
              <tr className="border-b border-slate-100">
                <th className="px-6 py-3 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '项目' : 'Project'}</th>
                <th className="px-6 py-3 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '分工' : 'Allocation'}</th>
                <th className="px-6 py-3 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">{language === 'zh' ? '预算/实际' : 'Bud/Act'}</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {budgetAnomalies.map((a, i) => (
                <tr key={i} className="hover:bg-slate-50 transition-colors">
                  <td className="px-6 py-4 text-xs font-bold text-slate-700 truncate max-w-[120px]">{a.projName}</td>
                  <td className="px-6 py-4 text-xs text-slate-500">{a.allocName}</td>
                  <td className="px-6 py-4 text-xs font-black text-right">
                    <span className="text-slate-500">{a.budget.toFixed(0)}</span>
                    <span className="text-slate-300 mx-1">/</span>
                    <span className="text-rose-600">{a.actual.toFixed(1)}</span>
                  </td>
                </tr>
              ))}
              {budgetAnomalies.length === 0 && (
                <tr>
                  <td colSpan={3} className="px-6 py-10 text-center text-slate-400 text-xs font-bold italic">
                    {language === 'zh' ? '暂无预算超支记录' : 'No budget overruns found'}
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// --- 5. Database Admin View (User-Friendly) ---
export function DatabaseAdminView() {
  const context = useContext(AppContext);
  const [activeTab, setActiveTab] = useState<'employees' | 'projects' | 'entries' | 'categories2' | 'categories3' | 'logs' | 'audit' | 'super_admin'>('entries');
  const [searchTerm, setSearchTerm] = useState('');
  const [editingItem, setEditingItem] = useState<any>(null);
  const [isAdding, setIsAdding] = useState(false);

  if (!context) return null;
  const { state, setState, language, handleSaveToDatabase, addLog } = context;

  const isSuperAdmin = state.currentUser?.role === 'admin';

  const filteredEmployees = getVisibleEmployees(state.currentUser, state.employees).filter(e => isSuperAdmin || e.id !== 'admin-1');

  const handleDeleteEntry = async (id: string) => {
    if (!isSuperAdmin) return;
    const entryToDelete = state.entries.find(e => e.id === id);
    if (!entryToDelete) return;
    if (!window.confirm(language === 'zh' ? '确定要删除这条记录吗？' : 'Are you sure you want to delete this entry?')) return;
    
    const newState = {
      ...state,
      entries: state.entries.filter(e => e.id !== id),
      deletedEntries: [entryToDelete, ...state.deletedEntries].slice(0, 100)
    };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '删除工时记录' : 'Remove Entry', `${entryToDelete.date} - ${getEmpName(state.employees.find(emp => emp.id === entryToDelete.employeeId), language)}`);
  };

  const handleBulkDeleteEntries = () => {
    if (!isSuperAdmin) return;
    const idsToDelete = state.entries
      .filter(e => {
        const emp = state.employees.find(emp => emp.id === e.employeeId);
        const empName = getEmpName(emp, language);
        const proj = state.projects.find(p => p.id === e.projectId);
        const projName = proj?.name || '';
        const search = searchTerm.toLowerCase();
        return e.date.includes(search) || empName.toLowerCase().includes(search) || projName.toLowerCase().includes(search) || (e.comment || '').toLowerCase().includes(search);
      })
      .map(e => e.id);

    if (idsToDelete.length === 0) return;
    if (!window.confirm(language === 'zh' ? `确定要删除当前筛选出的 ${idsToDelete.length} 条记录吗？此操作不可撤销。` : `Are you sure you want to delete ${idsToDelete.length} filtered entries? This cannot be undone.`)) return;

    const newState = {
      ...state,
      entries: state.entries.filter(e => !idsToDelete.includes(e.id))
    };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '批量删除工时记录' : 'Bulk Remove Entries', `Count: ${idsToDelete.length}`);
    alert(language === 'zh' ? `已删除 ${idsToDelete.length} 条记录` : `Deleted ${idsToDelete.length} entries`);
  };

  const handleDeleteEmployee = (id: string) => {
    if (!isSuperAdmin) return;
    const empToDelete = state.employees.find(e => e.id === id);
    if (!empToDelete) return;
    if (id === 'admin-1') {
      alert(language === 'zh' ? '不能删除超级管理员' : 'Cannot delete super admin');
      return;
    }
    if (!window.confirm(language === 'zh' ? '确定要删除该员工吗？' : 'Are you sure you want to delete this employee?')) return;
    const newState = { 
      ...state, 
      employees: state.employees.filter(e => e.id !== id),
      deletedEmployees: [empToDelete, ...state.deletedEmployees].slice(0, 50)
    };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '删除员工' : 'Remove Employee', getEmpName(empToDelete, language));
  };

  const handleDeleteProject = (id: string) => {
    if (!isSuperAdmin) return;
    const projToDelete = state.projects.find(p => p.id === id);
    if (!projToDelete) return;
    if (id === 'p-general' || id === 'p-leave') {
      alert(language === 'zh' ? '不能删除系统默认项目' : 'Cannot delete system projects');
      return;
    }
    if (!window.confirm(language === 'zh' ? '确定要删除该项目吗？' : 'Are you sure you want to delete this project?')) return;
    const newState = { 
      ...state, 
      projects: state.projects.filter(p => p.id !== id),
      deletedProjects: [projToDelete, ...state.deletedProjects].slice(0, 50)
    };
    setState(newState);
    handleSaveToDatabase(newState);
  };

  const handleUndoDatabaseDelete = () => {
    const hasDeleted = state.deletedEmployees.length > 0 || state.deletedProjects.length > 0 || state.deletedEntries.length > 0;
    if (!hasDeleted) return;

    let newState = { ...state };
    if (state.deletedEmployees.length > 0 && activeTab === 'employees') {
      const [item, ...remaining] = state.deletedEmployees;
      newState = { ...newState, employees: [item, ...state.employees], deletedEmployees: remaining };
    } else if (state.deletedProjects.length > 0 && activeTab === 'projects') {
      const [item, ...remaining] = state.deletedProjects;
      newState = { ...newState, projects: [item, ...state.projects], deletedProjects: remaining };
    } else if (state.deletedEntries.length > 0 && (activeTab as string) === 'entries') {
      const [item, ...remaining] = state.deletedEntries;
      newState = { ...newState, entries: [item, ...state.entries], deletedEntries: remaining };
    }

    setState(newState);
    handleSaveToDatabase(newState);
  };

  const handleSaveEdit = (updatedItem: any) => {
    let newState = { ...state };
    if ((activeTab as string) === 'entries') {
      newState.entries = isAdding 
        ? [...state.entries, { ...updatedItem, id: 'entry-' + Date.now() }]
        : state.entries.map(e => e.id === updatedItem.id ? updatedItem : e);
    } else if ((activeTab as string) === 'employees') {
      if (isAdding) {
        newState.employees = [...state.employees, { ...updatedItem, id: 'e-' + Date.now() }];
      } else {
        newState.employees = state.employees.map(e => {
          if (e.id === updatedItem.id) {
            // If password is empty in the form, keep the old one
            const password = updatedItem.password || e.password;
            return { ...updatedItem, password };
          }
          return e;
        });
      }
    } else if ((activeTab as string) === 'projects') {
      newState.projects = isAdding
        ? [...state.projects, { ...updatedItem, id: 'p-' + Date.now() }]
        : state.projects.map(p => p.id === updatedItem.id ? updatedItem : p);
    } else if ((activeTab as string) === 'categories2') {
      newState.categories2 = isAdding
        ? [...state.categories2, { ...updatedItem, id: 'c2-' + Date.now() }]
        : state.categories2.map(c => c.id === updatedItem.id ? updatedItem : c);
    } else if ((activeTab as string) === 'categories3') {
      newState.categories3 = isAdding
        ? [...state.categories3, { ...updatedItem, id: 'c3-' + Date.now() }]
        : state.categories3.map(c => c.id === updatedItem.id ? updatedItem : c);
    }
    setState(newState);
    handleSaveToDatabase(newState);
    
    let detail = `${activeTab}: ${updatedItem.id || 'New'}`;
    if (activeTab === 'projects') {
      const oldProj = state.projects.find(p => p.id === updatedItem.id);
      let budgetChanges: string[] = [];
      if (oldProj) {
        for (const allocId in updatedItem.budgets) {
          const oldVal = oldProj.budgets[allocId] || 0;
          const newVal = updatedItem.budgets[allocId] || 0;
          if (oldVal !== newVal) {
            const alloc = state.allocations.find(a => a.id === allocId);
            const allocName = alloc ? (language === 'zh' ? alloc.nameZh : alloc.nameEn) : allocId;
            budgetChanges.push(`${allocName}: ${oldVal}h -> ${newVal}h`);
          }
        }
      }
      detail += ` (${updatedItem.name})`;
      if (budgetChanges.length > 0) {
        detail += ` (Budget changes: ${budgetChanges.join(', ')})`;
      }
    }
    if (activeTab === 'employees') detail += ` (${updatedItem.nameEn || updatedItem.nameZh})`;
    if (activeTab === 'categories2' || activeTab === 'categories3') detail += ` (${getCatName(updatedItem, language)})`;
    if (activeTab === 'entries') detail += ` (${updatedItem.date}, ${updatedItem.hours}h)`;

    addLog(language === 'zh' ? '保存/修改数据' : 'Save/Edit Data', detail);
    setEditingItem(null);
    setIsAdding(false);
  };

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-8">
      <div className="flex gap-4">
        {isSuperAdmin && (
          <div className="flex gap-4">
            {activeTab === 'entries' && searchTerm && (
              <button 
                onClick={handleBulkDeleteEntries}
                className="bg-rose-100 text-rose-600 px-6 py-3 rounded-2xl font-black text-sm flex items-center gap-2 hover:bg-rose-200 transition-all"
              >
                <Trash2 size={18} />
                {language === 'zh' ? '批量删除筛选结果' : 'Bulk Delete Filtered'}
              </button>
            )}
            {(state.deletedEmployees.length > 0 || state.deletedProjects.length > 0 || state.deletedEntries.length > 0) && (
              <button 
                onClick={handleUndoDatabaseDelete}
                className="bg-slate-100 text-slate-600 px-6 py-3 rounded-2xl font-black text-sm flex items-center gap-2 hover:bg-slate-200 transition-all"
              >
                <RotateCcw size={18} />
                {language === 'zh' ? '恢复删除' : 'Undo Delete'}
              </button>
            )}
            <button 
              onClick={() => {
                setIsAdding(true);
                if ((activeTab as string) === 'entries') setEditingItem({ date: format(new Date(), 'yyyy-MM-dd'), employeeId: '', projectId: '', category2Id: '', category3Id: '', allocationId: '', hours: 0, comment: '' });
                if ((activeTab as string) === 'employees') setEditingItem({ nameZh: '', nameEn: '', type: 'WC', role: 'member', password: '' });
                if ((activeTab as string) === 'projects') setEditingItem({ name: '', managerId: '', budgets: {}, comment: '' });
                if ((activeTab as string) === 'categories2') setEditingItem({ nameZh: '', nameEn: '' });
                if ((activeTab as string) === 'categories3') setEditingItem({ parentId: '', nameZh: '', nameEn: '' });
              }}
              className="bg-blue-600 text-white px-6 py-3 rounded-2xl font-black text-sm flex items-center gap-2 shadow-lg hover:bg-blue-500 transition-all"
            >
              <Plus size={18} />
              {language === 'zh' ? '新增记录' : 'Add New'}
            </button>
          </div>
        )}
      </div>

      <div className="bg-white rounded-[40px] shadow-sm border border-slate-200 overflow-hidden">
        <div className="flex border-b border-slate-100 overflow-x-auto">
          {(['entries', 'employees', 'projects', 'categories2', 'categories3', 'logs', 'audit', 'super_admin'] as const)
            .filter(tab => tab !== 'super_admin' || isSuperAdmin)
            .map(tab => (
            <button
              key={tab}
              onClick={() => setActiveTab(tab as any)}
              className={cn(
                "px-8 py-6 font-black text-sm uppercase tracking-widest transition-all whitespace-nowrap",
                (activeTab as string) === tab ? "text-blue-600 border-b-4 border-blue-600 bg-blue-50/30" : "text-slate-400 hover:text-slate-600"
              )}
            >
              {language === 'zh' 
                ? (tab === 'entries' ? '工时记录' : tab === 'employees' ? '员工列表' : tab === 'projects' ? '项目列表' : tab === 'categories2' ? '分类2' : tab === 'categories3' ? '分类3' : tab === 'logs' ? '操作日志' : tab === 'audit' ? '数据核对' : '超级管理')
                : (tab === 'super_admin' ? 'SUPER ADMIN' : tab.toUpperCase())}
            </button>
          ))}
        </div>

        <div className="p-8">
          <div className="mb-6">
            <input 
              type="text"
              placeholder={language === 'zh' ? '搜索...' : 'Search...'}
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
              className="w-full max-w-md bg-slate-50 border border-slate-200 rounded-2xl px-6 py-3 font-bold text-sm outline-none focus:ring-4 focus:ring-blue-500/10"
            />
          </div>

          <div className="overflow-x-auto">
            {activeTab === 'entries' && (
              <table className="w-full">
                <thead>
                  <tr className="text-left border-b border-slate-100">
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '日期' : 'Date'}</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '提报时间' : 'Submit Time'}</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '员工' : 'Employee'}</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '分工' : 'Allocation'}</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Category 1</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Category 2</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Category 3</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '工时' : 'Hours'}</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '备注' : 'Comment'}</th>
                    {isSuperAdmin && <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">{language === 'zh' ? '操作' : 'Action'}</th>}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {state.entries
                    .filter(e => 
                      e.date.includes(searchTerm) || 
                      getEmpName(state.employees.find(emp => emp.id === e.employeeId) || null, language).toLowerCase().includes(searchTerm.toLowerCase()) ||
                      state.projects.find(p => p.id === e.projectId)?.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
                      getCatName(state.categories2.find(c => c.id === e.category2Id), language).toLowerCase().includes(searchTerm.toLowerCase()) ||
                      getCatName(state.categories3.find(c => c.id === e.category3Id), language).toLowerCase().includes(searchTerm.toLowerCase())
                    )
                    .slice(0, 100)
                    .map(entry => (
                    <tr key={entry.id} className="hover:bg-slate-50/50 transition-colors">
                      <td className="px-6 py-4 text-sm font-bold text-slate-600">{entry.date}</td>
                      <td className="px-6 py-4 text-[10px] font-bold text-slate-400 font-mono">
                        {entry.submittedAt ? format(parseISO(entry.submittedAt), 'MM-dd HH:mm') : '-'}
                      </td>
                      <td className="px-6 py-4 text-sm font-black text-slate-900">{getEmpName(state.employees.find(emp => emp.id === entry.employeeId) || null, language)}</td>
                      <td className="px-6 py-4 text-xs font-bold text-slate-500">
                        {(() => {
                          const alloc = state.allocations.find(a => a.id === entry.allocationId);
                          return alloc ? (language === 'zh' ? alloc.nameZh : alloc.nameEn) : '-';
                        })()}
                      </td>
                      <td className="px-6 py-4 text-sm font-bold text-blue-600">{state.projects.find(p => p.id === entry.projectId)?.name}</td>
                      <td className="px-6 py-4 text-xs font-bold text-slate-500">{getCatName(state.categories2.find(c => c.id === entry.category2Id), language)}</td>
                      <td className="px-6 py-4 text-xs font-bold text-slate-500">{getCatName(state.categories3.find(c => c.id === entry.category3Id), language)}</td>
                      <td className="px-6 py-4 text-sm font-black text-slate-900">{entry.hours}h</td>
                      <td className="px-6 py-4 text-xs text-slate-400 italic truncate max-w-[150px]">{entry.comment || '-'}</td>
                      {isSuperAdmin && (
                        <td className="px-6 py-4 text-right space-x-2">
                          <button onClick={() => setEditingItem(entry)} className="text-slate-300 hover:text-blue-600 transition-colors">
                            <Pencil size={18} />
                          </button>
                          <button onClick={() => handleDeleteEntry(entry.id)} className="text-slate-300 hover:text-rose-500 transition-colors">
                            <Trash2 size={18} />
                          </button>
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            {activeTab === 'employees' && (
              <table className="w-full">
                <thead>
                  <tr className="text-left border-b border-slate-100">
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '姓名' : 'Name'}</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '类型' : 'Type'}</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '角色' : 'Role'}</th>
                    {isSuperAdmin && <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">{language === 'zh' ? '操作' : 'Action'}</th>}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {filteredEmployees
                    .filter(e => getEmpName(e, language).toLowerCase().includes(searchTerm.toLowerCase()))
                    .map(emp => (
                    <tr key={emp.id} className="hover:bg-slate-50/50 transition-colors">
                      <td className="px-6 py-4 text-sm font-black text-slate-900">{getEmpName(emp, language)}</td>
                      <td className="px-6 py-4 text-xs font-bold text-slate-500">{emp.type}</td>
                      <td className="px-6 py-4">
                        <span className={cn(
                          "px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest",
                          emp.role === 'admin' ? "bg-blue-100 text-blue-600" : "bg-slate-100 text-slate-500"
                        )}>
                          {emp.role}
                        </span>
                      </td>
                      {isSuperAdmin && (
                        <td className="px-6 py-4 text-right space-x-2">
                          <button onClick={() => setEditingItem(emp)} className="text-slate-300 hover:text-blue-600 transition-colors">
                            <Pencil size={18} />
                          </button>
                          <button onClick={() => handleDeleteEmployee(emp.id)} className="text-slate-300 hover:text-rose-500 transition-colors">
                            <Trash2 size={18} />
                          </button>
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            {activeTab === 'projects' && (
              <div className="overflow-x-auto rounded-3xl border border-slate-100">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-100">
                      <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Project Name</th>
                      <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Manager</th>
                      <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Allocations</th>
                      <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Comment</th>
                      {isSuperAdmin && <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">Action</th>}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {state.projects
                      .filter(p => p.id !== 'p-leave')
                      .filter(p => p.name.toLowerCase().includes(searchTerm.toLowerCase()))
                      .map(proj => (
                      <tr key={proj.id} className="hover:bg-slate-50/50 transition-colors group">
                        <td className="px-6 py-4">
                          <div className="flex items-center gap-3">
                            <div className="w-8 h-8 bg-blue-50 rounded-lg flex items-center justify-center text-blue-600 font-black text-xs">
                              {proj.name.charAt(0)}
                            </div>
                            <span className="text-sm font-black text-slate-900">{proj.name}</span>
                          </div>
                        </td>
                        <td className="px-6 py-4 text-sm font-bold text-slate-600">
                          {state.employees.find(e => e.id === proj.managerId) ? getEmpName(state.employees.find(e => e.id === proj.managerId), language) : '-'}
                        </td>
                        <td className="px-6 py-4">
                          <span className="bg-slate-100 text-slate-500 text-[10px] font-black px-2 py-1 rounded-full uppercase tracking-tighter">
                            {Object.keys(proj.budgets).length} Allocations
                          </span>
                        </td>
                        <td className="px-6 py-4 text-xs font-medium text-slate-400 truncate max-w-[200px]">{proj.comment || '-'}</td>
                        {isSuperAdmin && (
                          <td className="px-6 py-4 text-right">
                            <div className="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                              <button onClick={() => setEditingItem(proj)} className="p-2 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all">
                                <Pencil size={16} />
                              </button>
                              <button onClick={() => handleDeleteProject(proj.id)} className="p-2 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all">
                                <Trash2 size={16} />
                              </button>
                            </div>
                          </td>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {activeTab === 'categories2' && (
              <table className="w-full">
                <thead>
                  <tr className="text-left border-b border-slate-100">
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">ID</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Name</th>
                    {isSuperAdmin && <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">{language === 'zh' ? '操作' : 'Action'}</th>}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {state.categories2
                    .filter(c => getCatName(c, language).toLowerCase().includes(searchTerm.toLowerCase()))
                    .map(cat => (
                    <tr key={cat.id} className="hover:bg-slate-50/50 transition-colors">
                      <td className="px-6 py-4 text-sm font-bold text-slate-400">{cat.id}</td>
                      <td className="px-6 py-4 text-sm font-black text-slate-900">{cat.nameZh}</td>
                      <td className="px-6 py-4 text-sm font-black text-slate-900">{cat.nameEn}</td>
                      {isSuperAdmin && (
                        <td className="px-6 py-4 text-right space-x-2">
                          <button onClick={() => setEditingItem(cat)} className="text-slate-300 hover:text-blue-600 transition-colors">
                            <Pencil size={18} />
                          </button>
                          <button onClick={() => {
                            const newState = { ...state, categories2: state.categories2.filter(c => c.id !== cat.id) };
                            setState(newState);
                            handleSaveToDatabase(newState);
                          }} className="text-slate-300 hover:text-rose-500 transition-colors">
                            <Trash2 size={18} />
                          </button>
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            {activeTab === 'categories3' && (
              <table className="w-full">
                <thead>
                  <tr className="text-left border-b border-slate-100">
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">ID</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Parent (Cat 2)</th>
                    <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Name</th>
                    {isSuperAdmin && <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">{language === 'zh' ? '操作' : 'Action'}</th>}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {state.categories3
                    .filter(c => getCatName(c, language).toLowerCase().includes(searchTerm.toLowerCase()))
                    .map(cat => (
                    <tr key={cat.id} className="hover:bg-slate-50/50 transition-colors">
                      <td className="px-6 py-4 text-sm font-bold text-slate-400">{cat.id}</td>
                      <td className="px-6 py-4 text-xs font-black text-indigo-600 uppercase tracking-widest">
                        {getCatName(state.categories2.find(c => c.id === cat.parentId), language)}
                      </td>
                      <td className="px-6 py-4 text-sm font-black text-slate-900">{cat.nameZh}</td>
                      <td className="px-6 py-4 text-sm font-black text-slate-900">{cat.nameEn}</td>
                      {isSuperAdmin && (
                        <td className="px-6 py-4 text-right space-x-2">
                          <button onClick={() => setEditingItem(cat)} className="text-slate-300 hover:text-blue-600 transition-colors">
                            <Pencil size={18} />
                          </button>
                          <button onClick={() => {
                            const newState = { ...state, categories3: state.categories3.filter(c => c.id !== cat.id) };
                            setState(newState);
                            handleSaveToDatabase(newState);
                          }} className="text-slate-300 hover:text-rose-500 transition-colors">
                            <Trash2 size={18} />
                          </button>
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            {activeTab === 'audit' && <DataAuditView state={state} language={language} />}

            {activeTab === 'logs' && (
              <div className="space-y-4 max-h-[600px] overflow-y-auto pr-2 custom-scrollbar">
                {(state.logs || []).map(log => (
                  <div key={log.id} className="p-4 bg-white rounded-2xl border border-slate-100 shadow-sm space-y-2">
                    <div className="flex justify-between items-start">
                      <span className="text-xs font-black text-indigo-600 uppercase tracking-widest">{log.action}</span>
                      <span className="text-[10px] font-bold text-slate-400">{format(parseISO(log.timestamp), 'yyyy-MM-dd HH:mm:ss')}</span>
                    </div>
                    <p className="text-sm font-bold text-slate-700">{log.details}</p>
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">By: {log.userName} ({log.userId})</p>
                  </div>
                ))}
              </div>
            )}

            {activeTab === 'super_admin' && <SuperAdminView />}
          </div>
        </div>
      </div>

      {/* Edit/Add Modal */}
      <AnimatePresence>
        {editingItem && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-[40px] shadow-2xl w-full max-w-lg overflow-hidden"
            >
              <div className="p-8 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                <h3 className="text-2xl font-black text-slate-900">
                  {isAdding ? (language === 'zh' ? '新增' : 'Add New') : (language === 'zh' ? '编辑' : 'Edit')}
                  {activeTab === 'entries' ? (language === 'zh' ? '记录' : ' Entry') : activeTab === 'employees' ? (language === 'zh' ? '员工' : ' Employee') : activeTab === 'projects' ? (language === 'zh' ? '项目' : ' Project') : (activeTab as string) === 'categories2' ? ' Category 2' : ' Category 3'}
                </h3>
                <button onClick={() => { setEditingItem(null); setIsAdding(false); }} className="text-slate-400 hover:text-slate-600 transition-colors">
                  <X size={24} />
                </button>
              </div>
              <div className="p-8 space-y-6 max-h-[70vh] overflow-y-auto custom-scrollbar">
                {activeTab === 'entries' && (
                  <>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '日期' : 'Date'}</label>
                      <input type="date" value={editingItem.date} onChange={e => setEditingItem({...editingItem, date: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold" />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '员工' : 'Employee'}</label>
                      <select value={editingItem.employeeId} onChange={e => {
                        setEditingItem({...editingItem, employeeId: e.target.value});
                      }} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold">
                        <option value="">{language === 'zh' ? '选择员工' : 'Select Employee'}</option>
                        {filteredEmployees.map(e => <option key={e.id} value={e.id}>{getEmpName(e, language)}</option>)}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '项目' : 'Project'}</label>
                      <select value={editingItem.projectId} onChange={e => setEditingItem({...editingItem, projectId: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold">
                        <option value="">{language === 'zh' ? '选择项目' : 'Select Project'}</option>
                        {state.projects.filter(p => p.id !== 'p-leave').map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Category 2</label>
                      <select value={editingItem.category2Id} onChange={e => setEditingItem({...editingItem, category2Id: e.target.value, category3Id: ''})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold">
                        <option value="">Select Category 2</option>
                        {state.categories2.map(c => <option key={c.id} value={c.id}>{getCatName(c, language)}</option>)}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Category 3</label>
                      <select value={editingItem.category3Id} onChange={e => setEditingItem({...editingItem, category3Id: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold">
                        <option value="">Select Category 3</option>
                        {state.categories3.filter(c => c.parentId === editingItem.category2Id).map(c => <option key={c.id} value={c.id}>{getCatName(c, language)}</option>)}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '工时' : 'Hours'}</label>
                      <input type="number" step="0.1" value={editingItem.hours} onChange={e => setEditingItem({...editingItem, hours: parseFloat(e.target.value)})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold" />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '备注' : 'Comment'}</label>
                      <textarea value={editingItem.comment} onChange={e => setEditingItem({...editingItem, comment: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold min-h-[80px]" />
                    </div>
                  </>
                )}
                {activeTab === 'employees' && (
                  <>
                    <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '中文姓名' : 'Chinese Name'}</label>
                        <input type="text" value={editingItem.nameZh || ''} onChange={e => setEditingItem({...editingItem, nameZh: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold" />
                      </div>
                      <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '英文姓名' : 'English Name'}</label>
                        <input type="text" value={editingItem.nameEn || ''} onChange={e => setEditingItem({...editingItem, nameEn: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold" />
                      </div>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '类型' : 'Type'}</label>
                      <select value={editingItem.type} onChange={e => setEditingItem({...editingItem, type: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold">
                        <option value="WC">WC</option>
                        <option value="BC">BC</option>
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '角色' : 'Role'}</label>
                      <select value={editingItem.role} onChange={e => setEditingItem({...editingItem, role: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold">
                        <option value="member">Member</option>
                        <option value="admin">Admin</option>
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                        {language === 'zh' ? '可查看他人范围' : 'View Others Scope'}
                      </label>
                      <select
                        value={editingItem.visibleMode || 'self'}
                        onChange={e => {
                          const v = e.target.value;
                          setEditingItem({
                            ...editingItem,
                            visibleMode: v,
                            visibleEmployeeIds: v === 'custom' ? (editingItem.visibleEmployeeIds || []) : []
                          });
                        }}
                        className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold"
                      >
                        <option value="self">{language === 'zh' ? '仅本人' : 'Self only'}</option>
                        <option value="all">{language === 'zh' ? '所有人' : 'All employees'}</option>
                        <option value="wc">{language === 'zh' ? '所有白领(WC)' : 'All WC'}</option>
                        <option value="bc">{language === 'zh' ? '所有蓝领(BC)' : 'All BC'}</option>
                        <option value="custom">{language === 'zh' ? '指定人员' : 'Custom list'}</option>
                      </select>
                      {editingItem.visibleMode === 'custom' && (
                        <div className="bg-slate-50 border border-slate-200 rounded-2xl p-3 max-h-[160px] overflow-y-auto space-y-1">
                          {state.employees.filter(e => e.id !== 'admin-1').map(emp => (
                            <label key={emp.id} className="flex items-center gap-2 text-[10px] font-bold cursor-pointer hover:bg-white p-1 rounded-xl">
                              <input
                                type="checkbox"
                                checked={Array.isArray(editingItem.visibleEmployeeIds) ? editingItem.visibleEmployeeIds.includes(emp.id) : false}
                                onChange={(ev) => {
                                  const cur = Array.isArray(editingItem.visibleEmployeeIds) ? editingItem.visibleEmployeeIds : [];
                                  const next = ev.target.checked ? [...cur, emp.id] : cur.filter((x: string) => x !== emp.id);
                                  setEditingItem({ ...editingItem, visibleEmployeeIds: next });
                                }}
                              />
                              <span>{getEmpName(emp, language)} <span className="text-slate-400">({emp.type})</span></span>
                            </label>
                          ))}
                        </div>
                      )}
                      <div className="text-[10px] text-slate-400 font-bold">
                        {language === 'zh'
                          ? '说明：默认仅本人。此处只控制“能看到谁”。是否可代填/修改请看下一个配置。'
                          : 'Note: This controls who you can VIEW. Editing scope is configured below.'}
                      </div>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                        {language === 'zh' ? '可代填/可修改范围' : 'Edit Others Scope'}
                      </label>
                      <select
                        value={editingItem.editMode || 'self'}
                        onChange={e => {
                          const v = e.target.value;
                          setEditingItem({
                            ...editingItem,
                            editMode: v,
                            editEmployeeIds: v === 'custom' ? (editingItem.editEmployeeIds || []) : []
                          });
                        }}
                        className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold"
                      >
                        <option value="self">{language === 'zh' ? '仅本人（只读他人）' : 'Self only'}</option>
                        <option value="all">{language === 'zh' ? '所有人（可代填/修改）' : 'All employees'}</option>
                        <option value="wc">{language === 'zh' ? '所有白领(WC)' : 'All WC'}</option>
                        <option value="bc">{language === 'zh' ? '所有蓝领(BC)' : 'All BC'}</option>
                        <option value="custom">{language === 'zh' ? '指定人员' : 'Custom list'}</option>
                      </select>
                      {editingItem.editMode === 'custom' && (
                        <div className="bg-slate-50 border border-slate-200 rounded-2xl p-3 max-h-[160px] overflow-y-auto space-y-1">
                          {state.employees.filter(e => e.id !== 'admin-1').map(emp => (
                            <label key={emp.id} className="flex items-center gap-2 text-[10px] font-bold cursor-pointer hover:bg-white p-1 rounded-xl">
                              <input
                                type="checkbox"
                                checked={Array.isArray(editingItem.editEmployeeIds) ? editingItem.editEmployeeIds.includes(emp.id) : false}
                                onChange={(ev) => {
                                  const cur = Array.isArray(editingItem.editEmployeeIds) ? editingItem.editEmployeeIds : [];
                                  const next = ev.target.checked ? [...cur, emp.id] : cur.filter((x: string) => x !== emp.id);
                                  setEditingItem({ ...editingItem, editEmployeeIds: next });
                                }}
                              />
                              <span>{getEmpName(emp, language)} <span className="text-slate-400">({emp.type})</span></span>
                            </label>
                          ))}
                        </div>
                      )}
                      <div className="text-[10px] text-slate-400 font-bold">
                        {language === 'zh'
                          ? '说明：这里控制“是否能代填/修改他人”。设置为“仅本人”时，仍可查看（若上面可查看范围允许），但不可编辑。'
                          : 'Note: This controls whether you can edit other people.'}
                      </div>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '密码' : 'Password'}</label>
                      <input type="text" value={editingItem.password || ''} onChange={e => setEditingItem({...editingItem, password: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold" placeholder="Leave empty to keep current" />
                    </div>
                  </>
                )}
                {activeTab === 'projects' && (
                  <>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '项目名称' : 'Project Name'}</label>
                      <input type="text" value={editingItem.name} onChange={e => setEditingItem({...editingItem, name: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold" />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '项目经理' : 'Manager'}</label>
                      <select value={editingItem.managerId} onChange={e => setEditingItem({...editingItem, managerId: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold">
                        <option value="">{language === 'zh' ? '选择经理' : 'Select Manager'}</option>
                        {state.employees.filter(e => e.role === 'admin').map(e => <option key={e.id} value={e.id}>{getEmpName(e, language)}</option>)}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '备注' : 'Comment'}</label>
                      <input type="text" value={editingItem.comment || ''} onChange={e => setEditingItem({...editingItem, comment: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold" />
                    </div>
                    <div className="space-y-4 pt-4 border-t border-slate-100">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '分工预算' : 'Allocation Budgets'}</label>
                      {state.allocations.map(alloc => (
                        <div key={alloc.id} className="space-y-2">
                          <div className="flex items-center justify-between gap-4">
                            <span className="text-xs font-bold text-slate-600">{language === 'zh' ? alloc.nameZh : alloc.nameEn}</span>
                            <input 
                              type="number" 
                              value={editingItem.budgets[alloc.id] || 0} 
                              onChange={e => setEditingItem({
                                ...editingItem, 
                                budgets: { ...editingItem.budgets, [alloc.id]: parseFloat(e.target.value) }
                              })} 
                              className="w-24 bg-slate-50 border border-slate-200 rounded-xl px-3 py-1 text-right font-black"
                            />
                          </div>
                          <input 
                            type="text" 
                            placeholder={language === 'zh' ? '预算备注' : 'Budget Comment'}
                            value={(editingItem.budgetComments && editingItem.budgetComments[alloc.id]) || ''}
                            onChange={e => setEditingItem({
                              ...editingItem,
                              budgetComments: { ...(editingItem.budgetComments || {}), [alloc.id]: e.target.value }
                            })}
                            className="w-full bg-slate-50 border border-slate-100 rounded-lg px-3 py-1 text-[10px] font-bold text-slate-400 outline-none focus:border-blue-300"
                          />
                        </div>
                      ))}
                    </div>
                  </>
                )}
                {(activeTab as string) === 'categories2' && (
                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '中文名称' : 'Name (ZH)'}</label>
                      <input type="text" value={editingItem.nameZh} onChange={e => setEditingItem({...editingItem, nameZh: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold" />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '英文名称' : 'Name (EN)'}</label>
                      <input type="text" value={editingItem.nameEn} onChange={e => setEditingItem({...editingItem, nameEn: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold" />
                    </div>
                  </div>
                )}
                {(activeTab as string) === 'categories3' && (
                  <>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Parent (Cat 2)</label>
                      <select value={editingItem.parentId} onChange={e => setEditingItem({...editingItem, parentId: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold">
                        <option value="">Select Category 2</option>
                        {state.categories2.map(c => <option key={c.id} value={c.id}>{getCatName(c, language)}</option>)}
                      </select>
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '中文名称' : 'Name (ZH)'}</label>
                        <input type="text" value={editingItem.nameZh} onChange={e => setEditingItem({...editingItem, nameZh: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold" />
                      </div>
                      <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '英文名称' : 'Name (EN)'}</label>
                        <input type="text" value={editingItem.nameEn} onChange={e => setEditingItem({...editingItem, nameEn: e.target.value})} className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold" />
                      </div>
                    </div>
                  </>
                )}
              </div>
              <div className="p-8 bg-slate-50/50 border-t border-slate-100 flex gap-4">
                <button 
                  onClick={() => { setEditingItem(null); setIsAdding(false); }}
                  className="flex-1 bg-white border border-slate-200 py-4 rounded-2xl font-black text-slate-400 hover:text-slate-600 transition-all"
                >
                  {language === 'zh' ? '取消' : 'Cancel'}
                </button>
                <button 
                  onClick={() => handleSaveEdit(editingItem)}
                  className="flex-1 bg-blue-600 text-white py-4 rounded-2xl font-black shadow-xl shadow-blue-600/20 hover:bg-blue-500 transition-all"
                >
                  {language === 'zh' ? '保存' : 'Save'}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </motion.div>
  );
}

// --- 6. Super Admin View ---
function SuperAdminView() {
  const context = useContext(AppContext);
  if (!context) return null;
  const { state, setState, language, handleSaveToDatabase, addLog } = context;

  const stats = {
    totalEntries: state.entries.length,
    totalEmployees: state.employees.length,
    totalProjects: state.projects.length,
    totalAllocations: state.allocations.length,
    totalLogs: state.logs.length,
  };

  const handleBackup = async () => {
    try {
      const { fileHandle, backupHandle: bh, currentUser, ...dataToSave } = state;
      const dataStr = JSON.stringify(dataToSave, null, 2);
      const encrypted = obfuscate(dataStr);
      const timestamp = format(new Date(), 'yyyyMMdd_HHmmss');
      const fileName = `timesheet_backup_${timestamp}.tmsdb`;

      let savedToFolder = false;
      const backupHandle = state.backupHandle || await get('tms_backup_handle');

      if (backupHandle) {
        try {
          const permission = await (backupHandle as any).queryPermission({ mode: 'readwrite' });
          if (permission === 'granted') {
            const fileHandle = await backupHandle.getFileHandle(fileName, { create: true });
            const writable = await fileHandle.createWritable();
            await writable.write(encrypted);
            await writable.close();
            savedToFolder = true;
          }
        } catch (e) {
          console.warn('Failed to save to backup folder, falling back to download', e);
        }
      }

      if (!savedToFolder) {
        const blob = new Blob([encrypted], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }

      const newState = {
        ...state,
        backupConfig: {
          ...state.backupConfig!,
          lastBackup: new Date().toISOString()
        }
      };
      setState(newState);
      handleSaveToDatabase(newState);
      addLog(language === 'zh' ? '手动备份' : 'Manual Backup', savedToFolder ? 'Saved to folder' : 'Downloaded');
      if (savedToFolder) {
        alert(language === 'zh' ? '备份已保存到指定文件夹' : 'Backup saved to specified folder');
      }
    } catch (error) {
      console.error('Backup failed:', error);
      alert(language === 'zh' ? '备份失败' : 'Backup failed');
    }
  };

  const clearOldLogs = () => {
    if (!window.confirm(language === 'zh' ? '确定要清除所有日志吗？' : 'Are you sure you want to clear all logs?')) return;
    const newState = { ...state, logs: [] };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '清除日志' : 'Clear Logs', 'All logs cleared');
  };

  const runDataCleanup = () => {
    const projectIds = new Set(state.projects.map(p => p.id));
    const employeeIds = new Set(state.employees.map(e => e.id));
    
    const validEntries = state.entries.filter(e => projectIds.has(e.projectId) && employeeIds.has(e.employeeId));
    const orphanedCount = state.entries.length - validEntries.length;

    if (orphanedCount === 0) {
      alert(language === 'zh' ? '未发现脏数据。' : 'No orphaned data found.');
      return;
    }

    if (!window.confirm(language === 'zh' ? `发现 ${orphanedCount} 条无关联记录，确定要清理吗？` : `Found ${orphanedCount} orphaned records. Clean them?`)) return;

    const newState = { ...state, entries: validEntries };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '数据清理' : 'Data Cleanup', `Removed ${orphanedCount} orphaned entries`);
  };

  const forceInitializeData = async () => {
    if (!window.confirm(language === 'zh' ? '确定要重置系统初始数据吗？这会覆盖当前的员工、分工和分类设置（工时记录会保留）。' : 'Are you sure you want to initialize system data? This will overwrite current employees, allocations, and categories (entries will be kept).')) return;
    
    const newState = {
      ...state,
      employees: INITIAL_STATE.employees,
      allocations: INITIAL_STATE.allocations,
      categories2: INITIAL_STATE.categories2,
      categories3: INITIAL_STATE.categories3,
      projects: INITIAL_STATE.projects
    };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '强制初始化数据' : 'Force Initialize Data', 'System data reset to code defaults');
    alert(language === 'zh' ? '初始化成功！' : 'Initialization successful!');
  };

  const [initClickCount, setInitClickCount] = useState(0);
  const [logClickCount, setLogClickCount] = useState(0);
  const longPressTimer = useRef<any>(null);

  const handleLongPressStart = (type: 'init' | 'log') => {
    longPressTimer.current = setTimeout(() => {
      if (type === 'init') forceInitializeData();
      else clearOldLogs();
    }, 5000);
  };

  const handleLongPressEnd = () => {
    if (longPressTimer.current) {
      clearTimeout(longPressTimer.current);
      longPressTimer.current = null;
    }
  };

  const handleQuickClick = (type: 'init' | 'log') => {
    if (type === 'init') {
      const newCount = initClickCount + 1;
      if (newCount >= 10) {
        forceInitializeData();
        setInitClickCount(0);
      } else {
        setInitClickCount(newCount);
      }
    } else {
      const newCount = logClickCount + 1;
      if (newCount >= 10) {
        clearOldLogs();
        setLogClickCount(0);
      } else {
        setLogClickCount(newCount);
      }
    }
  };

  return (
    <div className="space-y-8">
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        {[
          { label: language === 'zh' ? '总工时记录' : 'Total Entries', val: stats.totalEntries, icon: Clock, color: 'text-blue-600' },
          { label: language === 'zh' ? '项目总数' : 'Total Projects', val: stats.totalProjects, icon: LayoutGrid, color: 'text-indigo-600' },
          { label: language === 'zh' ? '员工总数' : 'Total Employees', val: stats.totalEmployees, icon: Users, color: 'text-emerald-600' },
        ].map(s => (
          <div key={s.label} className="bg-white p-8 rounded-[40px] border border-slate-200 shadow-sm flex items-center gap-6">
            <div className={cn("p-4 rounded-3xl bg-slate-50", s.color)}>
              <s.icon size={32} />
            </div>
            <div>
              <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{s.label}</p>
              <p className="text-3xl font-black text-slate-900">{s.val}</p>
            </div>
          </div>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {/* Backup Settings */}
        <div className="bg-white p-8 rounded-[40px] border border-slate-200 shadow-sm space-y-6">
          <h3 className="text-xl font-black text-slate-900 flex items-center gap-3">
            <Database className="text-blue-600" />
            {language === 'zh' ? '自动备份设置' : 'Auto Backup Settings'}
          </h3>
          <div className="space-y-4">
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">{language === 'zh' ? '备份频率' : 'Backup Frequency'}</label>
              <select 
                value={state.backupConfig?.frequency || 'none'}
                onChange={e => {
                  const newState = { ...state, backupConfig: { ...state.backupConfig!, frequency: e.target.value as any } };
                  setState(newState);
                  handleSaveToDatabase(newState);
                }}
                className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold text-slate-700 outline-none focus:ring-4 focus:ring-blue-500/10 transition-all"
              >
                <option value="none">{language === 'zh' ? '不定期' : 'None'}</option>
                <option value="daily">{language === 'zh' ? '每日' : 'Daily'}</option>
                <option value="weekly">{language === 'zh' ? '每周' : 'Weekly'}</option>
                <option value="monthly">{language === 'zh' ? '每月' : 'Monthly'}</option>
              </select>
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">{language === 'zh' ? '备份路径' : 'Backup Path'}</label>
              <div className="flex gap-2">
                <input 
                  type="text"
                  readOnly
                  value={state.backupHandle?.name || (language === 'zh' ? '未设置' : 'Not Set')}
                  className="flex-1 bg-slate-50 border border-slate-200 rounded-2xl px-5 py-3 font-bold text-slate-400 outline-none"
                />
                <button 
                  onClick={context.handleSetBackupFolder}
                  className="px-4 py-3 bg-white border border-slate-200 rounded-2xl text-blue-600 hover:bg-slate-50 transition-all shadow-sm"
                >
                  <FolderOpen size={20} />
                </button>
              </div>
            </div>
            <div className="p-4 bg-blue-50 rounded-2xl border border-blue-100">
              <p className="text-xs font-bold text-blue-700">
                {language === 'zh' ? '上次备份时间：' : 'Last Backup: '}
                {state.backupConfig?.lastBackup ? format(parseISO(state.backupConfig.lastBackup), 'yyyy-MM-dd HH:mm:ss') : (language === 'zh' ? '从未备份' : 'Never')}
              </p>
            </div>
            <button onClick={handleBackup} className="btn-primary w-full">
              <Download size={20} />
              {language === 'zh' ? '立即执行备份' : 'Backup Now'}
            </button>
          </div>
        </div>

        {/* Advanced Tools */}
        <div className="bg-white p-8 rounded-[40px] border border-slate-200 shadow-sm space-y-6">
          <h3 className="text-xl font-black text-slate-900 flex items-center gap-3">
            <ShieldCheck className="text-indigo-600" />
            {language === 'zh' ? '高级管理工具' : 'Advanced Management'}
          </h3>
          <div className="grid grid-cols-1 gap-4">
            <button onClick={runDataCleanup} className="flex items-center justify-between p-6 bg-slate-50 rounded-3xl border border-slate-100 hover:border-blue-500 transition-all group">
              <div className="flex items-center gap-4">
                <div className="p-3 bg-white rounded-2xl text-blue-600 shadow-sm">
                  <RotateCcw size={20} />
                </div>
                <div className="text-left">
                  <p className="font-black text-slate-900">{language === 'zh' ? '数据完整性检查' : 'Data Integrity Check'}</p>
                  <p className="text-xs font-bold text-slate-400">{language === 'zh' ? '清理无关联的脏数据' : 'Clean up orphaned records'}</p>
                </div>
              </div>
              <ChevronRight className="text-slate-300 group-hover:text-blue-600 transition-colors" />
            </button>

            <button 
              onMouseDown={() => handleLongPressStart('init')}
              onMouseUp={handleLongPressEnd}
              onMouseLeave={handleLongPressEnd}
              onTouchStart={() => handleLongPressStart('init')}
              onTouchEnd={handleLongPressEnd}
              onClick={() => handleQuickClick('init')}
              className="flex items-center justify-between p-6 bg-slate-50 rounded-3xl border border-slate-100 hover:border-emerald-500 transition-all group relative overflow-hidden"
            >
              <div className="flex items-center gap-4">
                <div className="p-3 bg-white rounded-2xl text-emerald-600 shadow-sm">
                  <UserCheck size={20} />
                </div>
                <div className="text-left">
                  <p className="font-black text-slate-900">{language === 'zh' ? '强制初始化系统数据' : 'Force Initialize Data'}</p>
                  <p className="text-xs font-bold text-slate-400">{language === 'zh' ? '长按5秒或点击10次' : 'Long press 5s or click 10x'}</p>
                </div>
              </div>
              <div className="flex flex-col items-end">
                <ChevronRight className="text-slate-300 group-hover:text-emerald-600 transition-colors" />
                {initClickCount > 0 && <span className="text-[10px] font-black text-emerald-500 mt-1">{initClickCount}/10</span>}
              </div>
            </button>

            <button 
              onMouseDown={() => handleLongPressStart('log')}
              onMouseUp={handleLongPressEnd}
              onMouseLeave={handleLongPressEnd}
              onTouchStart={() => handleLongPressStart('log')}
              onTouchEnd={handleLongPressEnd}
              onClick={() => handleQuickClick('log')}
              className="flex items-center justify-between p-6 bg-slate-50 rounded-3xl border border-slate-100 hover:border-rose-500 transition-all group relative overflow-hidden"
            >
              <div className="flex items-center gap-4">
                <div className="p-3 bg-white rounded-2xl text-rose-600 shadow-sm">
                  <Trash2 size={20} />
                </div>
                <div className="text-left">
                  <p className="font-black text-slate-900">{language === 'zh' ? '清空操作日志' : 'Clear Operation Logs'}</p>
                  <p className="text-xs font-bold text-slate-400">{language === 'zh' ? '长按5秒或点击10次' : 'Long press 5s or click 10x'}</p>
                </div>
              </div>
              <div className="flex flex-col items-end">
                <ChevronRight className="text-slate-300 group-hover:text-rose-600 transition-colors" />
                {logClickCount > 0 && <span className="text-[10px] font-black text-rose-500 mt-1">{logClickCount}/10</span>}
              </div>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

export function AppProvider({ children }: { children: React.ReactNode }) {
  const [state, setState] = useState<AppState>(INITIAL_STATE);
  const [view, setView] = useState<AppView>('login');
  const [language, setLanguage] = useState<'zh' | 'en'>('zh');
  const [isConnected, setIsConnected] = useState(false);
  const [connectionType, setConnectionType] = useState<'none' | 'server' | 'file'>('none');
  const [apiBase, setApiBase] = useState<string>(() => localStorage.getItem('tms_api_base') || '');
  const serverVersionRef = useRef<number>(Number(localStorage.getItem('tms_server_version') || '1') || 1);
  const serverStateSnapshotRef = useRef<string>(localStorage.getItem('tms_server_state_snapshot') || '');
  const materialReqSnapshotRef = useRef<Record<string, any>>({});
  // Server save queue: avoid concurrent PUT /api/state causing self-conflicts (single admin multi-tab / double click)
  const serverSaveQueueRef = useRef<Promise<void>>(Promise.resolve());
  const [permissionError, setPermissionError] = useState(false);

  // --- Notifications (server mode only, MVP) ---
  const [showNotificationCenter, setShowNotificationCenter] = useState(false);
  const [notificationUnread, setNotificationUnread] = useState(0);
  const [notificationItems, setNotificationItems] = useState<any[]>([]);
  const [notificationLoading, setNotificationLoading] = useState(false);
  const [procurementJump, setProcurementJump] = useState<any | null>(null);

  const refreshNotificationUnread = async () => {
    try {
      if (connectionType !== 'server' || !apiBase || !state.currentUser) return;
      const base = normalizeApiBase(apiBase);
      const resp = await fetch(`${base}/api/notifications/unread-count?userId=${encodeURIComponent(state.currentUser.id)}`, {
        method: 'GET',
        headers: { 'x-user-id': state.currentUser.id },
      });
      const data = await resp.json();
      if (resp.ok) setNotificationUnread(Number(data?.unread || 0) || 0);
    } catch (e) {
      // ignore
    }
  };

  const loadNotificationList = async () => {
    try {
      if (connectionType !== 'server' || !apiBase || !state.currentUser) return;
      setNotificationLoading(true);
      const base = normalizeApiBase(apiBase);
      const resp = await fetch(`${base}/api/notifications?top=200&userId=${encodeURIComponent(state.currentUser.id)}`, {
        method: 'GET',
        headers: { 'x-user-id': state.currentUser.id },
      });
      const data = await resp.json();
      if (resp.ok) {
        setNotificationItems(Array.isArray(data?.items) ? data.items : []);
        // also refresh unread count
        const unread = (Array.isArray(data?.items) ? data.items : []).filter((x: any) => !x.readAt).length;
        setNotificationUnread(unread);
      }
    } finally {
      setNotificationLoading(false);
    }
  };

  const markNotificationsRead = async (ids: number[]) => {
    try {
      if (connectionType !== 'server' || !apiBase || !state.currentUser) return;
      const base = normalizeApiBase(apiBase);
      await fetch(`${base}/api/notifications/mark-read`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-user-id': state.currentUser.id },
        body: JSON.stringify({ userId: state.currentUser.id, ids }),
      });
    } catch {}
    await refreshNotificationUnread();
  };

  const markAllNotificationsRead = async () => {
    try {
      if (connectionType !== 'server' || !apiBase || !state.currentUser) return;
      const base = normalizeApiBase(apiBase);
      await fetch(`${base}/api/notifications/mark-all-read`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-user-id': state.currentUser.id },
        body: JSON.stringify({ userId: state.currentUser.id }),
      });
    } catch {}
    await loadNotificationList();
  };

  useEffect(() => {
    if (CHINESE_ONLY_MODE && language !== 'zh') {
      setLanguage('zh');
    }
  }, [language]);

  // Poll unread count (server mode)
  useEffect(() => {
    if (connectionType !== 'server' || !apiBase || !state.currentUser) return;
    void refreshNotificationUnread();
    const t = window.setInterval(() => void refreshNotificationUnread(), 30000);
    return () => window.clearInterval(t);
  }, [connectionType, apiBase, state.currentUser?.id]);

  // When open center, load list
  useEffect(() => {
    if (!showNotificationCenter) return;
    void loadNotificationList();
  }, [showNotificationCenter]);

  // Cross-tab sync: keep serverVersionRef in sync with localStorage updates from other tabs
  useEffect(() => {
    const onStorage = (e: StorageEvent) => {
      if (e.key === 'tms_server_version' && e.newValue) {
        const v = Number(e.newValue || '0');
        if (Number.isFinite(v) && v > 0) serverVersionRef.current = v;
      }
      if (e.key === 'tms_server_state_snapshot' && typeof e.newValue === 'string') {
        serverStateSnapshotRef.current = e.newValue || '';
      }
    };
    window.addEventListener('storage', onStorage);
    return () => window.removeEventListener('storage', onStorage);
  }, []);

  const makeTempCsOrder = (projectNo: string, existing: Array<{ csOrder?: string }>) => {
    const raw = (projectNo || '').trim();
    const prefix = (raw.split(/\s+/)[0] || 'P').replace(/[^\w\-]/g, '');
    let n = 1;
    const mk = (i: number) => `${prefix}${String(i).padStart(3, '0')}`; // 项目号 + 00x（001/002…）
    const used = new Set((existing || []).map(p => String(p.csOrder || '').trim()).filter(Boolean));
    while (used.has(mk(n)) && n < 999) n++;
    return mk(n);
  };

  const loadDatabaseFromHandle = async (handle: FileSystemFileHandle) => {
    try {
      const file = await handle.getFile();
      const buffer = await file.arrayBuffer();
      const content = deobfuscate(new Uint8Array(buffer));
      const data = JSON.parse(content);
      
      setState(prev => ({
        ...prev,
        ...data,
        projects: Array.isArray(data.projects)
          ? data.projects.map((p: any) => {
              const cs = String(p?.csOrder || '').trim();
              return cs ? p : { ...p, csOrder: makeTempCsOrder(p?.id || p?.name || 'P', data.projects || []) };
            })
          : prev.projects,
        materialBoms: Array.isArray(data.materialBoms) ? data.materialBoms : [],
        deviceModels: Array.isArray(data.deviceModels) ? data.deviceModels : [],
        deviceBomTemplates: Array.isArray(data.deviceBomTemplates) ? data.deviceBomTemplates : [],
        deviceBomLines: Array.isArray(data.deviceBomLines) ? data.deviceBomLines : [],
        changeRecords: Array.isArray((data as any).changeRecords) ? (data as any).changeRecords : [],
        materialTrackingTemplate: Array.isArray(data.materialTrackingTemplate) && data.materialTrackingTemplate.length > 0
          ? data.materialTrackingTemplate
          : DEFAULT_MATERIAL_TRACKING_TEMPLATE,
        fileHandle: handle,
        currentUser: prev.currentUser
      }));
      setIsConnected(true);
      setConnectionType('file');
      setPermissionError(false);
    } catch (err) {
      console.error('Failed to load database:', err);
      if ((err as any).name === 'NotAllowedError') {
        setPermissionError(true);
      }
    }
  };

  const stableStringify = (obj: any) => {
    try { return JSON.stringify(obj); } catch { return ''; }
  };
  const deepClone = <T,>(obj: T): T => {
    try { return JSON.parse(JSON.stringify(obj)); } catch { return obj; }
  };

  const loadMaterialRequirementsFromServer = async (base: string) => {
    const b = normalizeApiBase(base);
    const resp = await fetch(`${b}/api/material-requirements`, { method: 'GET' });
    if (!resp.ok) throw new Error(`Load material requirements failed: ${resp.status}`);
    const payload = await resp.json();
    const items = Array.isArray(payload?.items) ? payload.items : [];
    // Snapshot map for diff
    const map: Record<string, any> = {};
    items.forEach((r: any) => {
      if (r?.id) map[String(r.id)] = deepClone(r);
    });
    materialReqSnapshotRef.current = map;
    setState(prev => ({ ...prev, materialRequirements: items }));
  };

  const loadStateFromServer = async (base: string) => {
    const b = normalizeApiBase(base);
    const health = await fetch(`${b}/api/health`, { method: 'GET' });
    if (!health.ok) throw new Error(`Health check failed: ${health.status}`);
    const resp = await fetch(`${b}/api/state`, { method: 'GET' });
    if (!resp.ok) throw new Error(`Load state failed: ${resp.status}`);
    const payload = await resp.json();
    const data = payload?.state || {};
    const metaVersion = Number(payload?.meta?.version || 1);
    if (Number.isFinite(metaVersion) && metaVersion > 0) {
      serverVersionRef.current = metaVersion;
      localStorage.setItem('tms_server_version', String(metaVersion));
    }
    // We store non-procurement data in tms_state; procurement table lives in its own table
    const { materialRequirements: _ignoreMr, ...rest } = data || {};
    const snapshot = stableStringify(rest);
    serverStateSnapshotRef.current = snapshot;
    localStorage.setItem('tms_server_state_snapshot', snapshot);
    setState(prev => ({
      ...prev,
      ...rest,
      projects: Array.isArray((rest as any).projects)
        ? (rest as any).projects.map((p: any) => {
            const cs = String(p?.csOrder || '').trim();
            return cs ? p : { ...p, csOrder: makeTempCsOrder(p?.id || p?.name || 'P', (rest as any).projects || []) };
          })
        : prev.projects,
      materialBoms: Array.isArray(data.materialBoms) ? data.materialBoms : [],
      deviceModels: Array.isArray(data.deviceModels) ? data.deviceModels : [],
      deviceBomTemplates: Array.isArray(data.deviceBomTemplates) ? data.deviceBomTemplates : [],
      deviceBomLines: Array.isArray(data.deviceBomLines) ? data.deviceBomLines : [],
      changeRecords: Array.isArray((data as any).changeRecords) ? (data as any).changeRecords : [],
      materialTrackingTemplate: Array.isArray(data.materialTrackingTemplate) && data.materialTrackingTemplate.length > 0
        ? data.materialTrackingTemplate
        : DEFAULT_MATERIAL_TRACKING_TEMPLATE,
      // 保持当前登录态，不被服务器数据覆盖
      currentUser: prev.currentUser
    }));
    // Load procurement table rows separately (row-level store)
    await loadMaterialRequirementsFromServer(b);
    setIsConnected(true);
    setConnectionType('server');
    setPermissionError(false);
  };

  const handleConnectDatabase = async () => {
    try {
      const guess = apiBase || (window.location.origin.startsWith('http') ? window.location.origin : 'http://localhost:3000');
      const input = (window.prompt(language === 'zh' ? '请输入服务器地址（例如 http://192.168.1.10:3000）' : 'Enter server URL (e.g. http://192.168.1.10:3000)', guess) || '').trim();
      if (!input) return;
      const base = normalizeApiBase(input);
      await loadStateFromServer(base);
      setApiBase(base);
      localStorage.setItem('tms_api_base', base);
    } catch (err) {
      if ((err as any).name === 'AbortError') return;
      console.error('Failed to connect database:', err);
      setPermissionError(true);
      alert(language === 'zh'
        ? `连接服务器失败：${(err as any)?.message || err}`
        : `Failed to connect: ${(err as any)?.message || err}`);
    }
  };

  const handleSaveToDatabase = async (newState: AppState) => {
    // --- Server mode ---
    if (connectionType === 'server' && apiBase) {
      // serialize server saves to avoid version self-conflict
      const run = async () => {
        const { fileHandle, backupHandle, currentUser, ...dataToSave } = newState;
        const operator = currentUser?.id || currentUser?.nameEn || currentUser?.nameZh || null;

        // 1) Save procurement rows with field-level merge
        const saveMaterialRequirements = async () => {
          const oldMap = materialReqSnapshotRef.current || {};
          const nextList = Array.isArray(newState.materialRequirements) ? newState.materialRequirements : [];
          const patches: any[] = [];
          const deleteIds: string[] = [];

          const eq = (a: any, b: any) => {
            if (a === b) return true;
            if ((a === null || a === undefined) && (b === null || b === undefined)) return true;
            return false;
          };

          const addChange = (changes: any, field: string, oldVal: any, newVal: any) => {
            changes[field] = { old: oldVal, new: newVal };
          };

          const sanitizeRowForUpsert = (r: any) => {
            const { _meta, ...rest } = r || {};
            return rest;
          };

          nextList.forEach((r: any) => {
            const id = String(r?.id || '').trim();
            if (!id) return;
            const old = oldMap[id];
            if (!old) {
              patches.push({ id, upsertRow: sanitizeRowForUpsert(r), changes: {} });
              return;
            }

            const changes: any = {};
            const ignoreKeys = new Set(['_meta', 'createdAt', 'updatedAt']);
            const keys = new Set([...Object.keys(old || {}), ...Object.keys(r || {})]);
            keys.forEach((k) => {
              if (ignoreKeys.has(k)) return;
              if (k === 'customFields') return;
              const ov = (old as any)[k];
              const nv = (r as any)[k];
              // only support primitives here
              const isPrimitive = (v: any) => v === null || v === undefined || ['string', 'number', 'boolean'].includes(typeof v);
              if (!isPrimitive(ov) || !isPrimitive(nv)) return;
              if (!eq(ov, nv)) addChange(changes, k, ov, nv);
            });

            const oldCF = (old as any).customFields || {};
            const newCF = (r as any).customFields || {};
            const cfKeys = new Set([...Object.keys(oldCF), ...Object.keys(newCF)]);
            cfKeys.forEach((k) => {
              const ov = oldCF[k];
              const nv = newCF[k];
              const isPrimitive = (v: any) => v === null || v === undefined || ['string', 'number', 'boolean'].includes(typeof v);
              if (!isPrimitive(ov) || !isPrimitive(nv)) return;
              if (!eq(ov, nv)) addChange(changes, `customFields.${k}`, ov, nv);
            });

            if (Object.keys(changes).length) patches.push({ id, changes });
          });

          // Detect deletions (client removed rows)
          const nextIdSet = new Set(nextList.map((r: any) => String(r?.id || '').trim()).filter(Boolean));
          Object.keys(oldMap || {}).forEach((id) => {
            const sid = String(id || '').trim();
            if (sid && !nextIdSet.has(sid)) deleteIds.push(sid);
          });

          if (!patches.length && !deleteIds.length) return { ok: true, applied: 0, deleted: 0, conflicts: [] };

          const post = async (forceFieldsById?: Record<string, string[]>) => {
            const bodyPatches = patches.map(p => {
              if (!forceFieldsById) return p;
              const ff = forceFieldsById[p.id];
              return ff ? { ...p, forceFields: ff } : p;
            });
            const resp = await fetch(`${normalizeApiBase(apiBase)}/api/material-requirements/patch`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ patches: bodyPatches, deleteIds, operator })
            });
            const payload = await resp.json().catch(() => ({}));
            // 兼容：如果后端还没更新（不支持 deleteIds），这里会导致“覆盖导入后刷新又变多”
            if (deleteIds.length > 0 && resp.ok && payload && (payload.deleted === undefined || payload.deleted === null)) {
              console.warn('[material-requirements] Server does not support deleteIds yet. Please update backend.');
              alert(language === 'zh'
                ? '检测到服务器后端未更新：覆盖导入/清空BOM 的“删除旧行”不会生效（刷新后行数会反弹）。请用最新部署包更新并重启后端。'
                : 'Server backend is outdated: deletions are not supported. Please update & restart backend.');
            }
            return { status: resp.status, payload };
          };

          const first = await post();
          if (first.status === 409) {
            const conflicts = Array.isArray(first.payload?.conflicts) ? first.payload.conflicts : [];
            const msg = language === 'zh'
              ? `检测到 ${conflicts.length} 处冲突（同一行同一列被他人修改）。\n\n是否用“我的值”覆盖这些冲突字段？\n- 确定：覆盖冲突字段\n- 取消：重新加载服务器数据`
              : `Detected ${conflicts.length} field conflicts.\nOverwrite with your values? (Cancel to reload)`;
            const overwrite = window.confirm(msg);
            if (!overwrite) {
              await loadMaterialRequirementsFromServer(apiBase);
              return { ok: false, conflict: true };
            }
            const byId: Record<string, string[]> = {};
            conflicts.forEach((c: any) => {
              const id = String(c.id || '');
              const field = String(c.field || '');
              if (!id || !field) return;
              if (!byId[id]) byId[id] = [];
              byId[id].push(field);
            });
            const second = await post(byId);
            if (second.status === 409) {
              alert(language === 'zh' ? '仍存在冲突，请重新加载后再操作。' : 'Still conflicting. Please reload.');
              await loadMaterialRequirementsFromServer(apiBase);
              return { ok: false, conflict: true };
            }
          } else if (first.status >= 400) {
            throw new Error(first.payload?.error || `patch failed: ${first.status}`);
          }

          // Update snapshot to latest local state after successful patch
          const map: Record<string, any> = {};
          nextList.forEach((r: any) => { if (r?.id) map[String(r.id)] = deepClone(r); });
          materialReqSnapshotRef.current = map;
          return { ok: true };
        };

        await saveMaterialRequirements();

        // 2) Save non-procurement state blob with optimistic locking (avoid writing materialRequirements)
        const { materialRequirements: _mr, ...stateWithoutMR } = dataToSave as any;
        const nextSnapshot = stableStringify(stateWithoutMR);
        if (nextSnapshot !== serverStateSnapshotRef.current) {
          const putState = async (expectedVersion: number) => {
            const resp = await fetch(`${normalizeApiBase(apiBase)}/api/state`, {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                state: stateWithoutMR,
                operator,
                expectedVersion,
                audit: { action: 'save_state', details: 'client save (non-procurement)' }
              })
            });
            const payload = await resp.json().catch(() => ({}));
            return { resp, payload };
          };

          let { resp, payload } = await putState(serverVersionRef.current);

          if (resp.status === 409) {
            const latest = Number(payload?.meta?.version || 0);
            const by = payload?.meta?.updatedBy || '';
            const when = payload?.meta?.updatedAt || '';

            // Fast-path: if server already has the exact same snapshot we are trying to save
            // (e.g. double click, request retry, multi-tab race), treat as success and just sync version.
            try {
              const sResp = await fetch(`${normalizeApiBase(apiBase)}/api/state`, { method: 'GET' });
              const sPayload = await sResp.json().catch(() => ({}));
              const sState = sPayload?.state;
              const sMetaVer = Number(sPayload?.meta?.version || 0);
              const sSnapshot = stableStringify(sState || {});
              if (sSnapshot === nextSnapshot && Number.isFinite(sMetaVer) && sMetaVer > 0) {
                serverVersionRef.current = sMetaVer;
                localStorage.setItem('tms_server_version', String(sMetaVer));
                serverStateSnapshotRef.current = nextSnapshot;
                localStorage.setItem('tms_server_state_snapshot', nextSnapshot);
                return;
              }
            } catch (e) {
              // ignore and continue conflict handling
            }

            // If the updater is current operator (same admin in another tab / earlier save),
            // auto-sync version and retry once to avoid unnecessary prompt.
            if (operator && by && String(by) === String(operator) && Number.isFinite(latest) && latest > 0) {
              serverVersionRef.current = latest;
              localStorage.setItem('tms_server_version', String(latest));
              ({ resp, payload } = await putState(latest));
            }
          }

          if (resp.status === 409) {
            const latest = Number(payload?.meta?.version || 0);
            const by = payload?.meta?.updatedBy || '';
            const when = payload?.meta?.updatedAt || '';
            const msg = language === 'zh'
              ? `保存冲突：服务器配置数据已被他人更新。\n最新版本：${latest}${by ? `\n更新人：${by}` : ''}${when ? `\n更新时间：${when}` : ''}\n\n是否重新加载服务器数据？`
              : `Save conflict: server state was updated.\nLatest version: ${latest}\nReload from server now?`;
            const reload = window.confirm(msg);
            if (reload) await loadStateFromServer(apiBase);
            else alert(language === 'zh' ? '本次保存已取消（未覆盖服务器）。' : 'Save cancelled (not overwritten).');
            return;
          }
          if (!resp.ok) {
            throw new Error(payload?.error || `Save failed: ${resp.status}`);
          }
          const newVer = Number(payload?.meta?.version || 0);
          if (Number.isFinite(newVer) && newVer > 0) {
            serverVersionRef.current = newVer;
            localStorage.setItem('tms_server_version', String(newVer));
          }
          serverStateSnapshotRef.current = nextSnapshot;
          localStorage.setItem('tms_server_state_snapshot', nextSnapshot);
        }
      };

      serverSaveQueueRef.current = serverSaveQueueRef.current.then(run, run);
      try {
        await serverSaveQueueRef.current;
      } catch (err) {
        console.error('Failed to save to server:', err);
        alert(language === 'zh'
          ? `保存到服务器失败：${(err as any)?.message || err}`
          : `Save to server failed: ${(err as any)?.message || err}`);
      }
      return;
    }

    // --- Local file mode (legacy fallback) ---
    if (!newState.fileHandle) return;
    try {
      const { fileHandle, backupHandle, currentUser, ...dataToSave } = newState;
      const content = JSON.stringify(dataToSave);
      const obfuscated = obfuscate(content);
      
      const writable = await (fileHandle as any).createWritable();
      await writable.write(obfuscated);
      await writable.close();

      // Backup
      if (backupHandle) {
        try {
          const fileName = `backup_${format(new Date(), 'yyyyMMdd_HHmmss')}.json`;
          const backupFile = await backupHandle.getFileHandle(fileName, { create: true });
          const backupWritable = await (backupFile as any).createWritable();
          await backupWritable.write(obfuscated);
          await backupWritable.close();
        } catch (bErr) {
          console.error('Backup failed:', bErr);
        }
      }
    } catch (err) {
      console.error('Failed to save to database:', err);
    }
  };

  const handleReauthorize = async () => {
    if (connectionType === 'server' && apiBase) {
      try {
        await loadStateFromServer(apiBase);
      } catch (err) {
        console.error('Server reconnect failed:', err);
        alert(language === 'zh' ? '服务器重连失败，请检查地址/网络。' : 'Reconnect failed.');
      }
      return;
    }
    if (!state.fileHandle) return;
    try {
      const permission = await (state.fileHandle as any).requestPermission({ mode: 'readwrite' });
      if (permission === 'granted') {
        await loadDatabaseFromHandle(state.fileHandle);
      }
    } catch (err) {
      console.error('Reauthorization failed:', err);
    }
  };

  const [initialLoading, setInitialLoading] = useState(true);

  useEffect(() => {
    document.title = "MBC工时与物料管理工具";
    const init = async () => {
      try {
        const savedHandle = await get('tms_file_handle');
        const savedBackupHandle = await get('tms_backup_handle');
        
        // F5 Refresh: Hydrate Auth State
        const savedUser = localStorage.getItem('timesheet_currentUser');
        if (savedUser) {
          try {
            const user = JSON.parse(savedUser);
            setState(prev => ({ ...prev, currentUser: user }));
            // Set view immediately to prevent login flash
            setView('entry');
          } catch (e) {
            console.error('Failed to parse saved user', e);
          }
        }

        // Prefer server mode if apiBase exists
        const base = localStorage.getItem('tms_api_base') || apiBase;
        if (base) {
          try {
            await loadStateFromServer(base);
            setApiBase(normalizeApiBase(base));
            localStorage.setItem('tms_api_base', normalizeApiBase(base));
            return;
          } catch (e) {
            // fallback to local file
            console.warn('Server auto-connect failed, fallback to local file:', e);
          }
        }

        if (savedHandle) {
          setState(prev => ({ ...prev, fileHandle: savedHandle, backupHandle: savedBackupHandle }));
          const permission = await (savedHandle as any).queryPermission({ mode: 'readwrite' });
          if (permission === 'granted') {
            await loadDatabaseFromHandle(savedHandle);
          } else {
            setPermissionError(true);
          }
        }
      } finally {
        setInitialLoading(false);
      }
    };
    init();
  }, []);

  useEffect(() => {
    if (initialLoading) return; // Don't wipe storage while we are hydration auth state
    if (state.currentUser) {
      localStorage.setItem('timesheet_currentUser', JSON.stringify(state.currentUser));
    } else {
      localStorage.removeItem('timesheet_currentUser');
    }
  }, [state.currentUser, initialLoading]);

  // 自动写入“员工多角色 -> 项目分工映射”（仅管理员、且当前没有任何映射时触发一次）
  useEffect(() => {
    if (initialLoading) return;
    if (!state.currentUser || state.currentUser.role !== 'admin') return;
    if ((state.userMappings || []).length > 0) return;
    const flagKey = 'tms_role_mapping_v3_applied';
    if (localStorage.getItem(flagKey) === '1') return;

    const { mappings, warnings } = buildUserMappingsFromRoleRows(EMBEDDED_ROLE_MAPPING_ROWS, state);
    if (!mappings.length) return;
    const newState = { ...state, userMappings: mappings };
    setState(newState);
    handleSaveToDatabase(newState);
    localStorage.setItem(flagKey, '1');
    if (warnings.length > 0) {
      console.warn('Role mapping import warnings:', warnings.slice(0, 50));
    }
  }, [initialLoading, state.currentUser, state.userMappings?.length]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleLogout = () => {
    setState(prev => ({ ...prev, currentUser: null }));
    setView('login');
  };

  const handleSetBackupFolder = async () => {
    try {
      const handle = await (window as any).showDirectoryPicker();
      setState(prev => ({ ...prev, backupHandle: handle }));
      await set('tms_backup_handle', handle);
      alert(language === 'zh' ? '备份文件夹设置成功！' : 'Backup folder set successfully!');
    } catch (err) {
      if ((err as any).name === 'AbortError') return;
      console.error('Failed to set backup folder:', err);
    }
  };

  useEffect(() => {
    if (state.currentUser && view === 'login') {
      setView('entry');
    }
  }, [state.currentUser]);

  const handleExportExcel = () => {
    const sortedEntries = [...state.entries].sort((a, b) => {
      if (a.employeeId !== b.employeeId) return a.employeeId.localeCompare(b.employeeId);
      return a.date.localeCompare(b.date);
    });

    const headers = [
      language === 'zh' ? '日期' : 'Date',
      language === 'zh' ? '员工' : 'Employee',
      language === 'zh' ? '分工' : 'Allocation',
      language === 'zh' ? '项目' : 'Project',
      'Category 2',
      'Category 3',
      language === 'zh' ? '工时' : 'Hours',
      language === 'zh' ? '备注' : 'Comment'
    ];

    const uniqueEmpIds = Array.from(new Set(sortedEntries.map(e => e.employeeId)));
    const pastelColors = [
      'E3F2FD', 'F1F8E9', 'FFFDE7', 'F3E5F5', 'FBE9E7', 'E0F2F1', 'FFF3E0', 'FCE4EC',
      'E8EAF6', 'EFEBE9', 'FAFAFA', 'ECEFF1'
    ];

    const wb = XLSX.utils.book_new();
    const ws: any = {};
    
    // Style headers
    const headerStyle = {
      fill: { fgColor: { rgb: "334155" } },
      font: { color: { rgb: "FFFFFF" }, bold: true },
      alignment: { horizontal: "center", vertical: "center" },
      border: {
        top: { style: "thin", color: { rgb: "000000" } },
        bottom: { style: "thin", color: { rgb: "000000" } },
        left: { style: "thin", color: { rgb: "000000" } },
        right: { style: "thin", color: { rgb: "000000" } }
      }
    };

    headers.forEach((h, i) => {
      const cellRef = XLSX.utils.encode_cell({ r: 0, c: i });
      ws[cellRef] = { v: h, t: 's', s: headerStyle };
    });

    let lastEmpId = '';
    let empBlockStart = true;

    sortedEntries.forEach((e, rIdx) => {
      const emp = state.employees.find(emp => emp.id === e.employeeId);
      const proj = state.projects.find(p => p.id === e.projectId);
      const alloc = state.allocations.find(a => a.id === e.allocationId);
      const cat2 = state.categories2.find(c => c.id === e.category2Id);
      const cat3 = state.categories3.find(c => c.id === e.category3Id);
      const allocName = alloc ? (language === 'zh' ? alloc.nameZh : alloc.nameEn) : '';
      
      const empIdx = uniqueEmpIds.indexOf(e.employeeId);
      const baseColor = pastelColors[empIdx % pastelColors.length];
      
      // Alternating row color within block for readability
      const isEven = rIdx % 2 === 0;
      const bgColor = isEven ? baseColor : baseColor; // Could adjust slightly if needed

      const row = [
        e.date,
        getEmpName(emp, language),
        allocName,
        proj?.name || '',
        getCatName(cat2, language),
        getCatName(cat3, language),
        e.hours,
        e.comment
      ];

      const rowIdx = rIdx + 1;
      const isNewEmp = e.employeeId !== lastEmpId;
      lastEmpId = e.employeeId;

      row.forEach((val, cIdx) => {
        const cellRef = XLSX.utils.encode_cell({ r: rowIdx, c: cIdx });
        ws[cellRef] = {
          v: val,
          t: typeof val === 'number' ? 'n' : 's',
          s: {
            fill: { fgColor: { rgb: bgColor } },
            alignment: { vertical: "center" },
            border: {
              top: { style: isNewEmp ? "medium" : "thin", color: { rgb: isNewEmp ? "334155" : "CBD5E1" } },
              bottom: { style: "thin", color: { rgb: "CBD5E1" } },
              left: { style: "thin", color: { rgb: "CBD5E1" } },
              right: { style: "thin", color: { rgb: "CBD5E1" } }
            }
          }
        };

        // Add comment to hours cell if exists
        if (cIdx === 6 && e.comment) {
          ws[cellRef].c = [{ t: e.comment, a: "System" }];
        }
      });
    });

    // Set range
    const range = { s: { r: 0, c: 0 }, e: { r: sortedEntries.length, c: headers.length - 1 } };
    ws['!ref'] = XLSX.utils.encode_range(range);

    // Set column widths
    ws['!cols'] = [
      { wch: 12 }, { wch: 15 }, { wch: 20 }, { wch: 30 }, { wch: 20 }, { wch: 25 }, { wch: 10 }, { wch: 40 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Timesheet");
    XLSX.writeFile(wb, `TMS_Export_${format(new Date(), 'yyyyMMdd')}.xlsx`);
  };

  const handleBatchImportEntries = (files: FileList) => {
    Array.from(files).forEach(file => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const result = e.target?.result;
          if (!result) return;
          
          let imported: any;
          if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            const data = new Uint8Array(result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rawData = XLSX.utils.sheet_to_json(sheet);
            
            // Map common headers to internal keys
            const mappedEntries = rawData.map((row: any) => {
              const entry: any = {};
              const findVal = (keys: string[]) => {
                const key = Object.keys(row).find(k => keys.includes(k.trim()));
                return key ? row[key] : undefined;
              };

              entry.date = findVal(['日期', 'Date', 'date']);
              entry.hours = parseFloat(findVal(['工时', 'Hours', 'hours', 'Time', 'time']) || '0');
              entry.comment = findVal(['备注', 'Comment', 'comment', 'Details', 'details', '工作内容']) || '';
              
              const projVal = findVal(['项目', 'Project', 'project', 'projectId']);
              const proj = state.projects.find(p => p.name === projVal || p.id === projVal);
              entry.projectId = proj ? proj.id : (projVal || '');

              const cat2Val = findVal(['分类2', 'Category 2', 'category2', 'category2Id']);
              const cat2 = state.categories2.find(c => c.nameZh === cat2Val || c.nameEn === cat2Val || c.id === cat2Val);
              entry.category2Id = cat2 ? cat2.id : (cat2Val || '');

              const cat3Val = findVal(['分类3', 'Category 3', 'category3', 'category3Id']);
              const cat3 = state.categories3.find(c => c.nameZh === cat3Val || c.nameEn === cat3Val || c.id === cat3Val);
              entry.category3Id = cat3 ? cat3.id : (cat3Val || '');

              const empVal = findVal(['员工', 'Employee', 'employee', 'employeeId', '姓名', 'Name']);
              const emp = state.employees.find(e => e.nameZh === empVal || e.nameEn === empVal || e.id === empVal);
              entry.employeeId = emp ? emp.id : (state.currentUser?.id || '');

              const allocVal = findVal(['分工', 'Allocation', 'allocation', 'allocationId', '身份']);
              const alloc = state.allocations.find(a => a.nameZh === allocVal || a.nameEn === allocVal || a.id === allocVal);
              entry.allocationId = alloc ? alloc.id : '';

              return entry;
            }).filter(e => e.date && e.hours > 0);

            imported = { entries: mappedEntries };
          } else {
            imported = JSON.parse(result as string);
          }

          const newEntries = imported.entries || (Array.isArray(imported) ? imported : []);
          if (!Array.isArray(newEntries)) return;

          setState(prev => {
            const existingIds = new Set(prev.entries.map(ent => ent.id));
            const processedNew = newEntries.map((ent: any) => {
              if (!ent) return null;
              return {
                ...ent,
                id: ent.id || `imp-${Math.random().toString(36).substr(2, 9)}-${Date.now()}`
              };
            }).filter(ent => ent && !existingIds.has(ent.id));

            const newState = { ...prev, entries: [...processedNew, ...prev.entries] };
            handleSaveToDatabase(newState);
            return newState;
          });
        } catch (err) {
          console.error('Import failed', err);
        }
      };
      if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        reader.readAsArrayBuffer(file);
      } else {
        reader.readAsText(file);
      }
    });
    alert(language === 'zh' ? '数据导入处理中' : 'Import processing');
  };
  const handleExportMyData = () => {
    const blob = new Blob([JSON.stringify({ entries: state.entries }, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `My_Timesheet_${format(new Date(), 'yyyyMMdd')}.json`;
    a.click();
  };

  const handleExportFullSystem = () => {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `TMS_Full_System_${format(new Date(), 'yyyyMMdd')}.json`;
    a.click();
  };

  const addLog = (action: string, details: string) => {
    const newLog: LogEntry = {
      id: 'log-' + Date.now(),
      userId: state.currentUser?.id || 'system',
      userName: getEmpName(state.currentUser, language),
      action,
      details,
      timestamp: new Date().toISOString(),
    };
    setState(prev => {
      const newState = { ...prev, logs: [newLog, ...(prev.logs || [])].slice(0, 1000) };
      handleSaveToDatabase(newState);
      return newState;
    });
  };

  const updateUserSetting = (key: string, value: any) => {
    if (!state.currentUser) return;
    const userId = state.currentUser.id;
    setState(prev => {
      const currentSettings = prev.userSettings?.[userId] || {};
      const newState = {
        ...prev,
        userSettings: {
          ...prev.userSettings,
          [userId]: {
            ...currentSettings,
            [key]: value
          }
        }
      };
      handleSaveToDatabase(newState);
      return newState;
    });
  };

  if (initialLoading) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6">
        <div className="w-16 h-16 border-4 border-blue-600 border-t-transparent rounded-full animate-spin mb-4" />
        <p className="text-slate-400 font-black uppercase tracking-widest text-[10px] animate-pulse">
          {language === 'zh' ? '正在加载系统...' : 'Loading System...'}
        </p>
      </div>
    );
  }

  return (
    <AppContext.Provider value={{
      state,
      setState,
      view,
      language,
      setLanguage,
      handleExportFullSystem,
      handleBatchImportEntries,
      handleSaveToDatabase,
      handleLogout,
      updateUserSetting,
      setView,
      addLog,
      isConnected,
      connectionType,
      apiBase,
      handleConnectDatabase,
      handleSetBackupFolder,
      loadDatabaseFromHandle,
      permissionError,
      handleReauthorize,
    }}>
      {children}
    </AppContext.Provider>
  );
}

export function LegacyShell() {
  const context = useContext(AppContext);
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [newPassword, setNewPassword] = useState('');
  const lastDeniedRef = useRef<string>('');

  if (!context) return null;

  const {
    state,
    setState,
    view,
    setView,
    language,
    setLanguage,
    handleSaveToDatabase,
    handleLogout,
    isConnected,
    connectionType,
    apiBase,
    handleConnectDatabase,
    permissionError,
    handleReauthorize,
  } = context;

  // 页面可见性：不仅隐藏左侧菜单，也要阻止“已打开/手动切换”的页面继续访问
  useEffect(() => {
    const user = state.currentUser;
    if (!user) return;
    if (view === 'login') return;
    if (user.role === 'admin') return;

    const adminOnlyViews: AppView[] = ['pm_view', 'setup', 'page_access', 'admin_db'];
    const isAdminOnly = adminOnlyViews.includes(view);
    if (isAdminOnly) {
      if (lastDeniedRef.current !== view) {
        lastDeniedRef.current = view;
        alert(language === 'zh' ? '无权限访问该页面' : 'No permission');
      }
      setView('entry');
      return;
    }

    const cfg = state.pageAccess?.[view as any];
    const allowed = cfg?.mode === 'restricted' ? (cfg.userIds || []).includes(user.id) : true;
    if (!allowed) {
      if (lastDeniedRef.current !== view) {
        lastDeniedRef.current = view;
        alert(language === 'zh' ? '该页面已被管理员限制访问' : 'This page is restricted');
      }
      setView('entry');
    }
  }, [view, state.pageAccess, state.currentUser?.id, state.currentUser?.role, language, setView]);

  const handleChangePassword = async () => {
    if (!newPassword || !state.currentUser) return;
    const newState = {
      ...state,
      employees: state.employees.map(e => e.id === state.currentUser?.id ? { ...e, password: newPassword } : e)
    };
    setState(newState);
    await handleSaveToDatabase(newState);
    alert(language === 'zh' ? '密码修改成功' : 'Password changed');
    setShowPasswordModal(false);
    setNewPassword('');
  };

  return (
    <>
      <div className="min-h-screen bg-[#F0F2F5] font-sans text-slate-900 flex">
        {view !== 'login' && (
          <aside className="w-72 bg-[#1E293B] text-white flex flex-col sticky top-0 h-screen shadow-2xl">
            <div className="p-8">
              <div className="flex flex-col gap-4 mb-12">
                <div className="flex items-center gap-3">
                  <div className="w-12 h-12 bg-white rounded-2xl flex items-center justify-center shadow-xl shadow-blue-900/40 border border-white/20">
                    <span className="text-blue-600 font-black text-xl tracking-tighter select-none">SP</span>
                  </div>
                  <div>
                    <h1 className="font-black text-2xl tracking-tighter text-white leading-none">TMS</h1>
                    <span className="text-[10px] text-blue-400 font-black uppercase tracking-[0.2em] mt-1 block">SCHOTT PHARMA</span>
                  </div>
                </div>
              </div>

              <nav className="space-y-2">
                {(() => {
                  const user = state.currentUser;
                  const isAdmin = user?.role === 'admin';
                  const canSee = (pageId: string) => {
                    if (!user) return false;
                    if (isAdmin) return true;
                    const cfg = state.pageAccess?.[pageId];
                    if (cfg?.mode === 'restricted') return cfg.userIds.includes(user.id);
                    return true; // default: visible to all
                  };
                  return (
                    <>
                      <SideNavItem active={view === 'entry'} onClick={() => setView('entry')} icon={<Clock />} label={language === 'zh' ? '工时填报' : 'Time Entry'} />
                      {canSee('procurement') && (
                        <SideNavItem active={view === 'procurement'} onClick={() => setView('procurement')} icon={<ShoppingCart size={18} />} label={language === 'zh' ? '采购管理' : 'Procurement'} />
                      )}
                      {canSee('device_bom') && (
                        <SideNavItem active={view === 'device_bom'} onClick={() => setView('device_bom')} icon={<Layers size={18} />} label={language === 'zh' ? '设备BOM中心' : 'Device BOM Center'} />
                      )}
                      {canSee('production_forecast') && (
                        <SideNavItem active={view === 'production_forecast'} onClick={() => setView('production_forecast')} icon={<Calendar size={18} />} label={language === 'zh' ? '生产预测排班' : 'Production Forecast'} />
                      )}
                      {canSee('kit_readiness') && (
                        <SideNavItem active={view === 'kit_readiness'} onClick={() => setView('kit_readiness')} icon={<ClipboardCheck size={18} />} label={language === 'zh' ? '齐套检查' : 'Kit Readiness'} />
                      )}
                      {isAdmin && (
                        <>
                          <SideNavItem active={view === 'pm_view'} onClick={() => setView('pm_view')} icon={<BarChart3 />} label={language === 'zh' ? '项目看板' : 'Project Dashboard'} />
                          <SideNavItem active={view === 'setup'} onClick={() => setView('setup')} icon={<SettingsIcon />} label={language === 'zh' ? '项目配置' : 'Project Configuration'} />
                          <SideNavItem active={view === 'page_access'} onClick={() => setView('page_access')} icon={<SettingsIcon />} label={language === 'zh' ? '页面配置' : 'Page Access'} />
                          <SideNavItem active={view === 'admin_db'} onClick={() => setView('admin_db')} icon={<Database />} label={language === 'zh' ? '数据库管理' : 'DB Admin'} />
                        </>
                      )}
                    </>
                  );
                })()}
              </nav>
            </div>

            <div className="mt-auto p-8 space-y-4">
              <div className="p-4 bg-slate-800/50 rounded-2xl border border-slate-700/50">
                <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest mb-2">{language === 'zh' ? '当前用户' : 'Current User'}</p>
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 bg-blue-500/20 rounded-lg flex items-center justify-center text-blue-400 font-bold text-xs">
                    {getEmpName(state.currentUser, language)[0]}
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="text-xs font-bold truncate">{getEmpName(state.currentUser, language)}</p>
                    <p className="text-[10px] text-slate-500 font-medium uppercase">{state.currentUser?.role}</p>
                  </div>
                </div>
                <button onClick={handleLogout} className="w-full mt-4 text-[10px] font-black text-red-400 hover:text-red-300 uppercase tracking-widest text-left transition-colors">
                  {language === 'zh' ? '退出登录' : 'Logout'}
                </button>
              </div>

              <div className="space-y-2">
                <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">{language === 'zh' ? '数据同步' : 'Data Sync'}</p>
                {isConnected ? (
                  <div className="px-4 py-3 bg-emerald-500/10 border border-emerald-500/20 rounded-xl flex items-center gap-3 text-[10px] font-black text-emerald-400 shadow-sm shadow-emerald-500/5">
                    <div className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse"></div>
                    {language === 'zh'
                      ? (connectionType === 'server' ? `已连接服务器 ${apiBase ? `(${apiBase})` : ''}` : '已连接本地文件库')
                      : (connectionType === 'server' ? `CONNECTED TO SERVER ${apiBase ? `(${apiBase})` : ''}` : 'CONNECTED TO LOCAL FILE')}
                  </div>
                ) : (
                  <button
                    onClick={handleConnectDatabase}
                    className="w-full px-4 py-3 bg-amber-500/10 border border-amber-500/20 rounded-xl flex items-center gap-3 text-[10px] font-black text-amber-500 hover:bg-amber-500/20 transition-all"
                  >
                    <div className="w-2 h-2 bg-amber-500 rounded-full"></div>
                    {language === 'zh' ? '点击连接服务器' : 'CONNECT SERVER'}
                  </button>
                )}
                {permissionError && (
                  <button
                    onClick={handleReauthorize}
                    className="w-full px-4 py-3 bg-rose-500/10 border border-rose-500/20 rounded-xl flex items-center gap-3 text-[10px] font-black text-rose-500 hover:bg-rose-500/20 transition-all"
                  >
                    <AlertCircle size={12} />
                    {language === 'zh' ? '重新连接' : 'RECONNECT'}
                  </button>
                )}
              </div>

              {!CHINESE_ONLY_MODE && (
                <button onClick={() => setLanguage(language === 'zh' ? 'en' : 'zh')} className="w-full bg-slate-800 hover:bg-slate-700 py-3 rounded-xl text-xs font-bold text-slate-400 transition-all">
                  {language === 'zh' ? 'Switch to English' : '切换至中文'}
                </button>
              )}

              <button
                onClick={() => setShowPasswordModal(true)}
                className="w-full bg-slate-800 hover:bg-slate-700 py-3 rounded-xl text-xs font-bold text-slate-400 transition-all flex items-center justify-center gap-2"
              >
                <Lock size={14} />
                {language === 'zh' ? '修改密码' : 'Change Password'}
              </button>
            </div>
          </aside>
        )}

        <main className={cn("flex-1 py-4 pr-12 pl-0", (view === 'login') && "flex items-center justify-center p-0")}>
          <AnimatePresence mode="wait">
            {view === 'login' && <LoginView />}
            {view === 'register' && <RegisterView />}
            {view === 'entry' && <EntryView />}
            {view === 'procurement' && <ProcurementView />}
            {view === 'device_bom' && <DeviceBomCenterView />}
            {view === 'production_forecast' && <ProductionForecastBoard />}
            {view === 'kit_readiness' && <KitReadinessView />}
            {view === 'pm_view' && <ProjectDashboard />}
            {view === 'setup' && <SetupView />}
            {view === 'page_access' && <PageAccessView />}
            {view === 'admin_db' && <DatabaseAdminView />}
          </AnimatePresence>
        </main>
      </div>

      <AnimatePresence>
        {showPasswordModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-6">
            <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setShowPasswordModal(false)} className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm" />
            <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.9, opacity: 0 }} className="relative bg-white rounded-[40px] p-10 w-full max-w-md shadow-2xl border border-slate-100">
              <h3 className="text-2xl font-black text-slate-900 mb-6">{language === 'zh' ? '修改登录密码' : 'Change Password'}</h3>
              <div className="space-y-6">
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '新密码' : 'New Password'}</label>
                  <input
                    type="password"
                    value={newPassword}
                    onChange={e => setNewPassword(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-5 py-4 font-bold text-sm outline-none focus:ring-4 focus:ring-blue-500/10"
                    placeholder="••••••••"
                  />
                </div>
                <div className="flex gap-4">
                  <button onClick={() => setShowPasswordModal(false)} className="flex-1 py-4 rounded-2xl font-black text-sm text-slate-400 hover:bg-slate-50 transition-all">
                    {language === 'zh' ? '取消' : 'Cancel'}
                  </button>
                  <button onClick={handleChangePassword} className="flex-1 bg-blue-600 text-white py-4 rounded-2xl font-black text-sm shadow-xl shadow-blue-200 hover:bg-blue-700 transition-all">
                    {language === 'zh' ? '确认修改' : 'Confirm'}
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <footer className="py-10 border-t border-slate-100 flex justify-between items-center text-[10px] font-black text-slate-300 uppercase tracking-[0.2em] px-12">
        <p>© 2026 MBC TIMESHEET & PROCUREMENT SYS</p>
        <p className="flex items-center gap-2">
          CRAFTED WITH PRECISION BY <span className="text-slate-400">DARRAN</span>
        </p>
      </footer>
    </>
  );
}

export default function LegacyApp() {
  return (
    <AppProvider>
      <AppErrorBoundary>
        <LegacyShell />
      </AppErrorBoundary>
    </AppProvider>
  );
}

class AppErrorBoundary extends React.Component<{ children: React.ReactNode }, { error: any }> {
  constructor(props: any) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error: any) {
    return { error };
  }
  componentDidCatch(error: any, info: any) {
    // 避免白屏：把错误输出到控制台，便于用户截图反馈
    // eslint-disable-next-line no-console
    console.error('[AppErrorBoundary]', error, info);
  }
  render() {
    if (this.state.error) {
      const msg = String(this.state.error?.message || this.state.error || 'Unknown error');
      return (
        <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6">
          <div className="max-w-2xl w-full bg-white rounded-[32px] border border-slate-200 p-6 space-y-4 shadow-sm">
            <div className="text-xl font-black text-slate-900">页面渲染出错（避免白屏）</div>
            <div className="text-xs font-bold text-slate-500">
              请把下面错误信息截图发给我，我可以定位并修复。
            </div>
            <pre className="text-[11px] bg-slate-50 border border-slate-200 rounded-xl p-3 overflow-auto whitespace-pre-wrap break-words">{msg}</pre>
            <div className="flex gap-2">
              <button
                className="bg-blue-600 text-white px-4 py-2 rounded-xl text-xs font-black"
                onClick={() => window.location.reload()}
              >
                刷新重试
              </button>
              <button
                className="btn-secondary px-4 py-2 text-xs font-black"
                onClick={() => this.setState({ error: null })}
              >
                忽略并继续
              </button>
            </div>
          </div>
        </div>
      );
    }
    return this.props.children as any;
  }
}

function SideNavItem({ active, onClick, icon, label }: { active: boolean; onClick: () => void; icon: React.ReactElement; label: string }) {
  return (
    <button 
      onClick={onClick}
      className={cn(
        "flex items-center gap-4 w-full px-6 py-4 rounded-2xl text-sm font-bold transition-all",
        active ? "bg-blue-600 text-white shadow-lg shadow-blue-600/20" : "text-slate-400 hover:text-white hover:bg-slate-800"
      )}
    >
      {React.cloneElement(icon, { size: 20 } as any)}
      {label}
    </button>
  );
}

function RegisterView() {
  const context = useContext(AppContext);
  const [formData, setFormData] = useState({
    nameZh: '',
    nameEn: '',
    type: 'WC' as 'WC' | 'BC',
    password: ''
  });
  const [error, setError] = useState('');

  if (!context) return null;
  const { state, setState, language, handleSaveToDatabase, setView } = context;

  const handleRegister = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.nameZh || !formData.nameEn || !formData.password) {
      setError(language === 'zh' ? '请填写所有必填项' : 'Please fill all required fields');
      return;
    }

    const id = 'e-' + Date.now();
    const newEmp: Employee = {
      id,
      ...formData,
      role: 'member'
    };

    const newState = {
      ...state,
      employees: [...state.employees, newEmp]
    };
    setState(newState);
    handleSaveToDatabase(newState);
    alert(language === 'zh' ? '注册成功，请登录' : 'Registration successful, please login');
    setView('login');
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-6 bg-[#F0F2F5]">
      <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="w-full max-w-md p-10 bg-white rounded-[3rem] shadow-2xl border border-slate-100">
        <div className="flex flex-col items-center mb-8">
          <div className="w-16 h-16 bg-indigo-600 rounded-2xl flex items-center justify-center shadow-lg shadow-indigo-600/20 mb-4">
            <UserCheck className="text-white w-8 h-8" />
          </div>
          <h2 className="text-2xl font-black text-slate-900 tracking-tight">{language === 'zh' ? '新员工注册' : 'Registration'}</h2>
        </div>

        <form onSubmit={handleRegister} className="space-y-4">
          <div className="space-y-1">
            <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">{language === 'zh' ? '姓名 (中文/英文)' : 'Name (ZH/EN)'}</label>
            <div className="relative">
              <input 
                type="text" 
                list="employee-names"
                value={formData.nameZh} 
                onChange={e => {
                  const val = e.target.value;
                  const existing = state.employees.find(emp => emp.nameZh === val || emp.nameEn === val);
                  if (existing) {
                    setFormData({
                      ...formData,
                      nameZh: existing.nameZh,
                      nameEn: existing.nameEn,
                      type: existing.type
                    });
                  } else {
                    setFormData({...formData, nameZh: val});
                  }
                }}
                className="w-full px-4 py-3 bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:bg-white rounded-xl font-bold transition-all outline-none text-sm"
                placeholder={language === 'zh' ? '输入或选择姓名' : 'Type or select name'}
              />
              <datalist id="employee-names">
                {state.employees.map(e => (
                  <option key={e.id} value={e.nameZh}>{e.nameEn}</option>
                ))}
              </datalist>
            </div>
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">{language === 'zh' ? '蓝白领' : 'Type'}</label>
            <select 
              value={formData.type}
              onChange={e => setFormData({...formData, type: e.target.value as 'WC' | 'BC'})}
              className="w-full px-4 py-3 bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:bg-white rounded-xl font-bold transition-all outline-none text-sm"
            >
              <option value="WC">White-collar (白领)</option>
              <option value="BC">Blue-collar (蓝领)</option>
            </select>
          </div>

          <div className="space-y-1">
            <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">{language === 'zh' ? '密码' : 'Password'}</label>
            <input 
              type="password" 
              value={formData.password} 
              onChange={e => setFormData({...formData, password: e.target.value})}
              className="w-full px-4 py-3 bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:bg-white rounded-xl font-bold transition-all outline-none text-sm"
            />
          </div>

          {error && (
            <div className="p-3 bg-red-50 text-red-500 rounded-xl flex items-center gap-2 text-xs font-bold">
              <AlertCircle size={14} />
              {error}
            </div>
          )}

          <div className="pt-4 space-y-3">
            <button type="submit" className="w-full bg-indigo-600 hover:bg-indigo-500 text-white py-4 rounded-xl font-black shadow-lg shadow-indigo-600/20 transition-all">
              {language === 'zh' ? '立即注册' : 'Register Now'}
            </button>
            <button type="button" onClick={() => setView('login')} className="w-full text-slate-400 font-bold text-xs hover:text-slate-600 transition-colors">
              {language === 'zh' ? '已有账号？返回登录' : 'Already have an account? Login'}
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
}

// --- Login View ---
function LoginView() {
  const context = useContext(AppContext);
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');

  if (!context) return null;
  const { state, setState, language, setLanguage, handleLogout, isConnected, loadDatabaseFromHandle, handleConnectDatabase } = context;

  const [hasSavedHandle, setHasSavedHandle] = useState(false);

  useEffect(() => {
    get('tms_db_handle').then(handle => {
      if (handle) setHasSavedHandle(true);
    });
  }, []);

  const handleQuickConnect = async () => {
    try {
      const handle = await get('tms_db_handle');
      if (handle) {
        const permission = await handle.requestPermission({ mode: 'readwrite' });
        if (permission === 'granted') {
          await loadDatabaseFromHandle(handle);
        }
      }
    } catch (err) {
      console.error('Quick connect failed:', err);
    }
  };

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    const inputVal = username.trim();
    const emp = state.employees.find(emp => 
      emp.nameZh === inputVal || emp.nameEn === inputVal || emp.id === inputVal
    );
    
    if (!emp) {
      setError(language === 'zh' ? '用户不存在' : 'User not found');
      return;
    }

    const pwd = password || '1';
    if (emp.password && emp.password !== pwd) {
      setError(language === 'zh' ? '密码错误' : 'Invalid password');
      return;
    }

    setState(prev => ({ ...prev, currentUser: emp }));
    context.setView('entry');
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 relative overflow-hidden">
      {/* Background Accents */}
      <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-blue-600 via-indigo-600 to-emerald-600" />
      <div className="absolute -top-24 -left-24 w-96 h-96 bg-blue-600/5 rounded-full blur-3xl" />
      <div className="absolute -bottom-24 -right-24 w-96 h-96 bg-indigo-600/5 rounded-full blur-3xl" />

      <motion.div 
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="w-full max-w-md bg-white rounded-[40px] shadow-2xl shadow-slate-200/50 border border-slate-100 p-10 relative z-10"
      >
        <div className="flex justify-between items-center mb-10">
          <div className="flex items-center gap-3">
            <div className="w-12 h-12 bg-gradient-to-br from-blue-600 to-indigo-700 rounded-2xl flex items-center justify-center shadow-lg shadow-blue-200 border border-blue-100">
              <span className="text-white font-black text-xl tracking-tighter select-none">SP</span>
            </div>
            <div>
              <h1 className="text-2xl font-black text-slate-900 tracking-tight leading-none">MBC</h1>
              <span className="text-[10px] text-blue-600 font-black uppercase tracking-widest block mt-1">SCHOTT PHARMA</span>
            </div>
          </div>
          <button 
            onClick={() => setLanguage(language === 'zh' ? 'en' : 'zh')}
            className="text-[10px] font-black text-slate-400 uppercase tracking-widest hover:text-blue-600 transition-colors"
          >
            {language === 'zh' ? 'English' : '中文'}
          </button>
        </div>

        <form onSubmit={handleLogin} className="space-y-6">
          <div className="space-y-2">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">{language === 'zh' ? '用户名' : 'Username'}</label>
            <div className="relative group">
              <User className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-300 group-focus-within:text-blue-600 transition-colors" size={18} />
              <input 
                list="employee-list"
                value={username}
                onChange={e => setUsername(e.target.value)}
                className="w-full bg-slate-50 border border-slate-200 rounded-2xl pl-14 pr-6 py-4 font-bold text-slate-900 outline-none focus:border-blue-600 focus:ring-4 focus:ring-blue-600/5 transition-all"
                placeholder={language === 'zh' ? '请输入或选择用户' : 'Enter or select user'}
              />
              <datalist id="employee-list">
                {state.employees.map(emp => (
                  <option key={emp.id} value={getEmpName(emp, language)} />
                ))}
              </datalist>
            </div>
          </div>

          <div className="space-y-2">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">{language === 'zh' ? '密码' : 'Password'}</label>
            <div className="relative group">
              <Lock className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-300 group-focus-within:text-blue-600 transition-colors" size={18} />
              <input 
                type="password" 
                value={password} 
                onChange={e => setPassword(e.target.value)}
                placeholder="••••••••"
                className="w-full bg-slate-50 border border-slate-200 rounded-2xl pl-14 pr-6 py-4 font-bold text-slate-900 outline-none focus:border-blue-600 focus:ring-4 focus:ring-blue-600/5 transition-all"
              />
            </div>
          </div>

          {error && <p className="text-rose-500 text-xs font-bold text-center">{error}</p>}

          <div className="pt-4 space-y-3">
            {!isConnected ? (
              <button 
                type="button"
                onClick={handleConnectDatabase}
                className="w-full bg-amber-500 hover:bg-amber-400 text-white py-5 rounded-2xl font-black text-sm shadow-xl shadow-amber-500/20 transition-all flex items-center justify-center gap-2"
              >
                <Database size={18} />
                {language === 'zh' ? '请先连接数据库文件 (.tmsdb)' : 'Connect Database (.tmsdb) First'}
              </button>
            ) : (
              <button 
                type="submit"
                className="w-full bg-blue-600 text-white py-5 rounded-2xl font-black text-sm shadow-xl shadow-blue-600/20 hover:bg-blue-500 hover:-translate-y-0.5 active:translate-y-0 transition-all"
              >
                {language === 'zh' ? '登录系统' : 'Login System'}
              </button>
            )}
            {hasSavedHandle && !isConnected && (
              <button 
                type="button" 
                onClick={handleQuickConnect}
                className="w-full text-blue-600 font-bold text-xs hover:text-blue-700 transition-colors flex items-center justify-center gap-1"
              >
                <RotateCcw size={14} />
                {language === 'zh' ? '尝试自动重新连接' : 'Try Quick Reconnect'}
              </button>
            )}
          </div>
        </form>
      </motion.div>
      
      <p className="mt-8 text-[10px] font-black text-slate-300 uppercase tracking-[0.2em] relative z-10">
        Time Management System v2.5
      </p>
    </div>
  );
}

// --- Database Init View ---
function DatabaseInitView() {
  const context = useContext(AppContext);
  if (!context) return null;
  const { language, setState, handleSaveToDatabase } = context;

  const handleCreateNew = async () => {
    try {
      const handle = await (window as any).showSaveFilePicker({
        suggestedName: 'TMS_Data.tmsdb',
        types: [{ description: 'TMS Database', accept: { 'application/octet-stream': ['.tmsdb'] } }]
      });
      const newState = { ...INITIAL_STATE, fileHandle: handle };
      await handleSaveToDatabase(newState);
      setState(newState);
      alert(language === 'zh' ? '新数据库已创建！' : 'New database created!');
    } catch (err) {
      console.error('Creation failed:', err);
    }
  };

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="max-w-md text-center space-y-8 p-12">
      <div className="w-24 h-24 bg-blue-100 rounded-[2.5rem] flex items-center justify-center mx-auto mb-8">
        <FileSpreadsheet className="text-blue-600 w-12 h-12" />
      </div>
      <h2 className="text-3xl font-black text-slate-900">{language === 'zh' ? '初始化数据库' : 'Initialize Database'}</h2>
      <p className="text-slate-500 font-medium leading-relaxed">
        {language === 'zh' 
          ? '请在 Q 盘创建一个新的数据库文件，或选择已有的文件进行连接。所有数据将以加密 JSON 格式存储。' 
          : 'Please create a new database file on Q Drive or select an existing one. All data is stored in JSON format.'}
      </p>
      <div className="space-y-4">
        <button onClick={handleCreateNew} className="w-full bg-blue-600 hover:bg-blue-500 text-white py-5 rounded-2xl font-black text-lg shadow-xl shadow-blue-600/20 transition-all">
          {language === 'zh' ? '创建新数据库' : 'Create New Database'}
        </button>
        <button onClick={() => window.location.reload()} className="w-full bg-slate-100 hover:bg-slate-200 text-slate-600 py-5 rounded-2xl font-black text-lg transition-all">
          {language === 'zh' ? '返回登录' : 'Back to Login'}
        </button>
      </div>
    </motion.div>
  );
}

// --- Weekly Grid View (Excel-like) ---
// --- Weekly Grid Component (Excel-like) ---
function WeeklyGrid({ language, state, setState, handleSaveToDatabase, initialRange }: { language: 'zh' | 'en', state: AppState, setState: any, handleSaveToDatabase: any, initialRange?: { start: string, end: string } }) {
  const { isConnected, loadDatabaseFromHandle, updateUserSetting } = useContext(AppContext)!;
  const [startDate, setStartDate] = useState(initialRange?.start || format(startOfWeek(new Date(), { weekStartsOn: 1 }), 'yyyy-MM-dd'));
  const [endDate, setEndDate] = useState(initialRange?.end || format(addDays(parseISO(startDate), 6), 'yyyy-MM-dd'));
  const hotRef = useRef<any>(null);

  // 自动保存草稿功能 (带防抖)
  const saveTimerRef = useRef<any>(null);
  const isResettingRef = useRef<boolean>(false);
  const hasUserEditedRef = useRef<boolean>(false);

  const saveDraft = (data: any[]) => {
    if (isResettingRef.current) return;
    if (!hasUserEditedRef.current) return;
    
    // 检查数据是否真的有效（即是否包含除了员工姓名以外的任何工时）
    const hasAnyHours = data.some(row => {
      if (!Array.isArray(row)) return false;
      for (let i = 5; i < row.length; i++) {
        const val = parseFloat(row[i]);
        if (!isNaN(val) && val > 0) return true;
      }
      return false;
    });

    if (!hasAnyHours) return; // 不保存全空的草稿

    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(async () => {
      if (state.currentUser) {
        const draftKey = `tms_draft_${state.currentUser.id}`;
        await set(draftKey, { data, timestamp: Date.now(), startDate, endDate });
      }
    }, 1000);
  };

  // 加载草稿
  const [draftRestored, setDraftRestored] = useState(false);
  const loadDraft = async () => {
    if (state.currentUser && isConnected) {
      const draftKey = `tms_draft_${state.currentUser.id}`;
      const draft = await get(draftKey);
      if (draft && draft.startDate === startDate && draft.endDate === endDate) {
        // 只有当日期范围一致时才建议恢复
        if (window.confirm(language === 'zh' ? '检测到未提交的草稿，是否恢复？' : 'Unsubmitted draft detected, restore it?')) {
          setHotData(draft.data);
          setDraftRestored(true);
        } else {
          // 如果用户选择不恢复，则清理草稿，避免重复提示
          await del(draftKey);
        }
      }
    }
  };

  const days = useMemo(() => {
    try {
      return eachDayOfInterval({
        start: parseISO(startDate),
        end: parseISO(endDate)
      });
    } catch (e) {
      return [];
    }
  }, [startDate, endDate]);

  // 月视图（按你要求“按月”）：选择月份后自动切换为当月 1 日 ~ 月末
  const monthValue = useMemo(() => {
    try {
      return format(parseISO(startDate), 'yyyy-MM');
    } catch (e) {
      return '';
    }
  }, [startDate]);

  const setMonthRange = (ym: string) => {
    if (!ym) return;
    const d = parseISO(`${ym}-01`);
    setStartDate(format(startOfMonth(d), 'yyyy-MM-dd'));
    setEndDate(format(endOfMonth(d), 'yyyy-MM-dd'));
  };

  const visibleEmployees = useMemo(
    () => getVisibleEmployees(state.currentUser, state.employees),
    [state.currentUser, state.employees]
  );
  const editableEmployees = useMemo(
    () => getEditableEmployees(state.currentUser, state.employees),
    [state.currentUser, state.employees]
  );
  const editableIdSet = useMemo(() => new Set(editableEmployees.map(e => e.id)), [editableEmployees]);
  const defaultSelectableIds = useMemo(() => visibleEmployees.filter(e => e.role !== 'admin').map(e => e.id), [visibleEmployees]);
  const savedSelectedEmployeeIds = useMemo(() => {
    const key = state.currentUser ? `${state.currentUser.id}_selectedEmployeeIds` : '';
    const raw = key ? state.userSettings?.[key] : null;
    if (Array.isArray(raw)) {
      const allowed = new Set(defaultSelectableIds);
      const cleaned = raw.filter((id: string) => allowed.has(id));
      if (cleaned.length) return cleaned;
    }
    return defaultSelectableIds;
  }, [state.currentUser, state.userSettings, defaultSelectableIds]);
  const savedLastSelection = useMemo(() => {
    const key = state.currentUser ? `${state.currentUser.id}_lastSelection` : '';
    return (key ? state.userSettings?.[key] : null) || null;
  }, [state.currentUser, state.userSettings]);
  const [selectedEmployeeIds, setSelectedEmployeeIds] = useState<string[]>(savedSelectedEmployeeIds);
  const selectedEmployees = visibleEmployees.filter(e => selectedEmployeeIds.includes(e.id));

  useEffect(() => {
    setSelectedEmployeeIds(prev => {
      if (prev.length === savedSelectedEmployeeIds.length && prev.every((id, idx) => id === savedSelectedEmployeeIds[idx])) {
        return prev;
      }
      return savedSelectedEmployeeIds;
    });
  }, [savedSelectedEmployeeIds]);

  const [quickAdj, setQuickAdj] = useState({
    employeeId: state.currentUser?.id || '',
    date: format(new Date(), 'yyyy-MM-dd'),
    overtime: '',
    leave: ''
  });

  const handleSaveAdj = () => {
    if (!quickAdj.employeeId || !quickAdj.date) return;
    if (!editableIdSet.has(quickAdj.employeeId)) {
      alert(language === 'zh' ? '你只有查看权限，不能保存他人的加班/请假调整。' : 'Read-only: cannot save other people adjustments.');
      return;
    }
    const ot = parseFloat(quickAdj.overtime) || 0;
    const lv = parseFloat(quickAdj.leave) || 0;
    if (ot <= 0 && lv <= 0) {
      alert(language === 'zh' ? '请至少填写加班或请假小时数（>0）' : 'Please enter overtime or leave hours (>0).');
      return;
    }
    
    const existing = state.dailyAdjustments.find(a => a.employeeId === quickAdj.employeeId && a.date === quickAdj.date);
    let newAdjs;
    if (existing) {
      newAdjs = state.dailyAdjustments.map(a => a.id === existing.id ? { ...a, overtime: ot, leave: lv } : a);
    } else {
      newAdjs = [...state.dailyAdjustments, {
        id: 'adj-' + Date.now(),
        employeeId: quickAdj.employeeId,
        date: quickAdj.date,
        overtime: ot,
        leave: lv
      }];
    }
    const newState = { ...state, dailyAdjustments: newAdjs };
    setState(newState);
    handleSaveToDatabase(newState);
    alert(language === 'zh' ? '调整已保存' : 'Adjustment saved');
    setQuickAdj(prev => ({ ...prev, overtime: '', leave: '' }));
  };

  // 行内 Employee 下拉：只允许选择“可代填/可修改”的人员，避免创建/改写无权限人员的记录
  const employeeOptions = useMemo(() => editableEmployees.map(e => getEmpName(e, language)), [editableEmployees, language]);
  const norm = (v: any) => String(v ?? '').trim();
  const getEmployeeByName = (name: string) =>
    visibleEmployees.find(e => norm(getEmpName(e, language)) === norm(name));

  useEffect(() => {
    if (quickAdj.employeeId && quickAdj.date) {
      const existing = state.dailyAdjustments.find(a => a.employeeId === quickAdj.employeeId && a.date === quickAdj.date);
      if (existing) {
        setQuickAdj(prev => ({
          ...prev,
          overtime: existing.overtime > 0 ? existing.overtime.toString() : '',
          leave: existing.leave > 0 ? existing.leave.toString() : ''
        }));
      } else {
        setQuickAdj(prev => ({ ...prev, overtime: '', leave: '' }));
      }
    }
  }, [quickAdj.employeeId, quickAdj.date]);

  const allocationOptions = useMemo(() => state.allocations.map(a => language === 'en' ? (a.nameEn || a.nameZh) : (a.nameZh || a.nameEn)), [state.allocations, language]);
  const getAllocationByName = (name: string) => state.allocations.find(a => norm(language === 'en' ? (a.nameEn || a.nameZh) : (a.nameZh || a.nameEn)) === norm(name));

  // Initial data for Handsontable: pre-fill with visible employees
  const [hotData, setHotData] = useState<any[]>([]);
  const [cellComments, setCellComments] = useState<any[]>([]);
  const [modifiedCells, setModifiedCells] = useState<Set<string>>(new Set());

  // Pre-fill logic: fetch existing entries for the range
  useEffect(() => {
    if (draftRestored) return; // Important: Skip sync if draft was restored to prevent overwrite
    
    const initialRows: any[] = [];
    const initialComments: any[] = [];
    const dateList = days.map(d => format(d, 'yyyy-MM-dd'));

    selectedEmployees.forEach((emp, empIdx) => {
      const empEntries = state.entries.filter(e => e.employeeId === emp.id && e.date >= startDate && e.date <= endDate);
      
      if (empEntries.length > 0) {
        // Group entries by project/cat2/cat3/alloc
        const groups: { [key: string]: TimesheetEntry[] } = {};
        empEntries.forEach(e => {
          const key = `${e.projectId}|${e.category2Id}|${e.category3Id}|${e.allocationId}`;
          if (!groups[key]) groups[key] = [];
          groups[key].push(e);
        });

        Object.values(groups).forEach((group, groupIdx) => {
          const first = group[0];
          const proj = state.projects.find(p => p.id === first.projectId);
          const cat2 = state.categories2.find(c => c.id === first.category2Id);
          const cat3 = state.categories3.find(c => c.id === first.category3Id);
          const alloc = state.allocations.find(a => a.id === first.allocationId);

          // 支持“多条信息完全一致”的场景：
          // 同一天可能存在多条完全相同的记录（同员工/项目/分类/分工/日期）。
          // Handsontable 单元格只能显示一条记录，所以这里按“车道(lane)”拆成多行：
          // 每行最多承载每天一条记录，从而做到“一一展示”。
          const byDate: { [date: string]: TimesheetEntry[] } = {};
          group.forEach(e => {
            if (!byDate[e.date]) byDate[e.date] = [];
            byDate[e.date].push(e);
          });
          // 保持显示稳定：按提交时间/ID 排序
          Object.values(byDate).forEach(list => {
            list.sort((a: any, b: any) => {
              const at = (a.submittedAt || a.updatedAt || a.createdAt || '') as string;
              const bt = (b.submittedAt || b.updatedAt || b.createdAt || '') as string;
              if (at && bt && at !== bt) return at.localeCompare(bt);
              return (a.id || '').localeCompare(b.id || '');
            });
          });

          const laneCount = Math.max(1, ...dateList.map(d => (byDate[d]?.length || 0)));
          for (let lane = 0; lane < laneCount; lane++) {
            const rowIndex = initialRows.length;
            const row = [
              getEmpName(emp, language),
              proj?.name || '',
              getCatName(cat2, language),
              getCatName(cat3, language),
              alloc ? (language === 'en' ? (alloc.nameEn || alloc.nameZh) : (alloc.nameZh || alloc.nameEn)) : '',
              ...dateList.map((d, dIdx) => {
                const dayEntry = byDate[d]?.[lane];
                if (dayEntry && dayEntry.comment) {
                  initialComments.push({ row: rowIndex, col: 5 + dIdx, comment: { value: dayEntry.comment } });
                }
                return dayEntry ? dayEntry.hours : '';
              })
            ];
            initialRows.push(row);
          }
        });
      } else {
        // Just one empty row for this employee if no entries
        initialRows.push([
          getEmpName(emp, language),
          savedLastSelection?.projectName || '',
          savedLastSelection?.cat2Name || '',
          savedLastSelection?.cat3Name || '',
          savedLastSelection?.allocName || '',
          ...Array(days.length).fill('')
        ]);
      }
    });

    // Add some empty rows for flexibility
    const emptyRowsCount = Math.max(5, 15 - initialRows.length);
    const emptyRows = Array(emptyRowsCount).fill(null).map(() => ['', '', '', '', '', ...Array(days.length).fill('')]);
    setHotData([...initialRows, ...emptyRows]);
    setCellComments(initialComments);
  }, [startDate, endDate, state.entries, state.projects, state.categories2, state.categories3, state.allocations, language, selectedEmployeeIds, savedLastSelection]);

  // Update data structure if days change or visible employees change
  useEffect(() => {
    if (isConnected) {
      loadDraft();
    }
  }, [startDate, endDate, isConnected]);

  useEffect(() => {
  }, [hotData]);

  const projectOptions = useMemo(() => state.projects.filter(p => p.id !== 'p-leave').map(p => p.name), [state.projects]);
  const getProjectByName = (name: string) => state.projects.find(p => norm(p.name) === norm(name));
  const getProjectByNameLoose = (name: string) => {
    const n = norm(name).toLowerCase();
    if (!n) return null;
    const hits = state.projects.filter(p => norm(p.name).toLowerCase().includes(n));
    return hits.length === 1 ? hits[0] : null;
  };
  
  const cat2Options = useMemo(() => state.categories2.map(c => getCatName(c, language)), [state.categories2, language]);
  const getCat2ByName = (name: string) => state.categories2.find(c => norm(getCatName(c, language)) === norm(name));

  const getCat3Options = (cat2Name: string) => {
    const cat2 = getCat2ByName(cat2Name);
    if (!cat2) return [];
    return state.categories3.filter(c => c.parentId === cat2.id).map(c => getCatName(c, language));
  };
  const getCat3ByName = (name: string, cat2Id: string) => state.categories3.find(c => norm(getCatName(c, language)) === norm(name) && c.parentId === cat2Id);

  // Calculate total table width for alignment
  const colWidths = {
    emp: 120,
    proj: 150,
    cat2: 120,
    cat3: 150,
    alloc: 120,
    day: 80,
    rowHeader: 50
  };
  const totalTableWidth = useMemo(() => colWidths.emp + colWidths.proj + colWidths.cat2 + colWidths.cat3 + colWidths.alloc + (days.length * colWidths.day) + colWidths.rowHeader, [colWidths, days.length]);

  const tableColumns = useMemo(() => [
    { type: 'dropdown', source: employeeOptions, width: colWidths.emp },
    { type: 'dropdown', source: projectOptions, width: colWidths.proj },
    { type: 'dropdown', source: cat2Options, width: colWidths.cat2 },
    { 
      type: 'dropdown', 
      source: function(this: any, query: any, process: any) {
        try {
          const row = this.row;
          const hotInstance = hotRef.current?.hotInstance;
          if (typeof row === 'number' && hotInstance) {
            const cat2Name = hotInstance.getDataAtCell(row, 2);
            process(getCat3Options(cat2Name));
          } else {
            process([]);
          }
        } catch (e) {
          process([]);
        }
      },
      width: colWidths.cat3 
    },
    { type: 'dropdown', source: allocationOptions, width: colWidths.alloc },
    ...days.map(() => ({ type: 'numeric', format: '0.0', width: colWidths.day }))
  ], [employeeOptions, projectOptions, cat2Options, allocationOptions, colWidths, days.length, language]);

  // Calculate daily totals per employee for monitoring
  const employeeDailyTotals = useMemo(() => {
    const totals: { [empId: string]: { [date: string]: number } } = {};
    
    // ONLY use current grid data (hotData) for the days in the current view
    // because hotData is pre-filled with existing entries from the database.
    hotData.forEach(row => {
      if (!row || !Array.isArray(row)) return;
      const empName = row[0];
      if (!empName) return;
      const emp = getEmployeeByName(empName);
      if (!emp) return;

      for (let i = 0; i < days.length; i++) {
        const val = parseFloat(row[5 + i]);
        if (!isNaN(val) && val > 0) {
          const date = format(days[i], 'yyyy-MM-dd');
          if (!totals[emp.id]) totals[emp.id] = {};
          totals[emp.id][date] = (totals[emp.id][date] || 0) + val;
        }
      }
    });

    return totals;
  }, [hotData, state.entries, days]);

  const employeeDailyExpectedHours = useMemo(() => {
    const expected: { [empId: string]: { [date: string]: number } } = {};
    selectedEmployeeIds.forEach(empId => {
      expected[empId] = {};
      days.forEach(day => {
        const date = format(day, 'yyyy-MM-dd');
        const adj = state.dailyAdjustments.find(a => a.employeeId === empId && a.date === date);
        expected[empId][date] = Math.max(0, 8 + (adj?.overtime || 0) - (adj?.leave || 0));
      });
    });
    return expected;
  }, [selectedEmployeeIds, days, state.dailyAdjustments]);

  const handleSubmit = async () => {
    const hotInstance = hotRef.current?.hotInstance;
    if (!hotInstance) return;
    
    const data = hotInstance.getData();
    const newEntries: TimesheetEntry[] = [];

    try {
      for (let rowIdx = 0; rowIdx < data.length; rowIdx++) {
        const row = data[rowIdx];
        const empName = norm(row[0]);
        const projName = norm(row[1]);
        const cat2Name = norm(row[2]);
        const cat3Name = norm(row[3]);
        const allocName = norm(row[4]);
        
        // Check if this row has any hours filled
        let hasHours = false;
        for (let i = 0; i < days.length; i++) {
          const val = parseFloat(row[5 + i]);
          if (!isNaN(val) && val > 0) {
            hasHours = true;
            break;
          }
        }

        if (!hasHours) continue; // Skip empty rows

        // Validation logic
        const emp = getEmployeeByName(empName);
        const proj = getProjectByName(projName) || getProjectByNameLoose(projName);
        const cat2 = getCat2ByName(cat2Name);
        const cat3 = cat2 ? getCat3ByName(cat3Name, cat2.id) : null;
        const alloc = getAllocationByName(allocName);

        if (!emp) {
          alert(language === 'zh' ? `第 ${rowIdx + 1} 行：姓名不能为空` : `Row ${rowIdx + 1}: Name cannot be empty`);
          setTimeout(() => {
            hotInstance.selectCell(rowIdx, 0, rowIdx, 0, true);
          }, 50);
          return;
        }

        // 权限：仅当在“可代填/可修改范围”内才允许提交该员工的工时
        if (!editableIdSet.has(emp.id)) {
          alert(language === 'zh'
            ? `第 ${rowIdx + 1} 行：你只有查看权限，不能代填/修改【${getEmpName(emp, language)}】的工时。`
            : `Row ${rowIdx + 1}: Read-only. You cannot edit entries for ${getEmpName(emp, language)}.`);
          return;
        }

        if (!proj) {
          alert(language === 'zh'
            ? `第 ${rowIdx + 1} 行：项目不能为空或不在列表中（当前值：${projName || '-'}）`
            : `Row ${rowIdx + 1}: Project is empty or invalid (current: ${projName || '-'})`);
          setTimeout(() => {
            hotInstance.selectCell(rowIdx, 1, rowIdx, 1, true);
          }, 50);
          return;
        }

        if (proj.id === 'p-general') {
          if (allocName) {
            alert(language === 'zh' ? `第 ${rowIdx + 1} 行：选择 General 时分工必须为空` : `Row ${rowIdx + 1}: Allocation must be empty for General time`);
            setTimeout(() => {
              hotInstance.selectCell(rowIdx, 4, rowIdx, 4, true);
            }, 50);
            return;
          }
        } else {
          // Project mode
          if (!cat2Name) {
            alert(language === 'zh' ? `第 ${rowIdx + 1} 行：分类二不能为空` : `Row ${rowIdx + 1}: Category 2 cannot be empty`);
            setTimeout(() => {
              hotInstance.selectCell(rowIdx, 2, rowIdx, 2, true);
            }, 50);
            return;
          }
          if (!cat3Name) {
            alert(language === 'zh' ? `第 ${rowIdx + 1} 行：分类三不能为空` : `Row ${rowIdx + 1}: Category 3 cannot be empty`);
            setTimeout(() => {
              hotInstance.selectCell(rowIdx, 3, rowIdx, 3, true);
            }, 50);
            return;
          }
          if (!allocName) {
            alert(language === 'zh' ? `第 ${rowIdx + 1} 行：分工不能为空` : `Row ${rowIdx + 1}: Allocation cannot be empty`);
            setTimeout(() => {
              hotInstance.selectCell(rowIdx, 4, rowIdx, 4, true);
            }, 50);
            return;
          }
          if (!cat2 || !cat3) {
            alert(language === 'zh' ? `第 ${rowIdx + 1} 行：无效的分类选择` : `Row ${rowIdx + 1}: Invalid category selection`);
            setTimeout(() => {
              hotInstance.selectCell(rowIdx, 2, rowIdx, 2, true);
            }, 50);
            return;
          }
          if (!alloc) {
            alert(language === 'zh' ? `第 ${rowIdx + 1} 行：无效的分工选择` : `Row ${rowIdx + 1}: Invalid allocation selection`);
            setTimeout(() => {
              hotInstance.selectCell(rowIdx, 4, rowIdx, 4, true);
            }, 50);
            return;
          }
        }

        for (let i = 0; i < days.length; i++) {
          const val = row[5 + i];
          const hours = parseFloat(val);
          const date = format(days[i], 'yyyy-MM-dd');

          // Get comment from Handsontable comments plugin
          const comment = hotInstance.getPlugin('comments').getCommentAtCell(rowIdx, 5 + i) || '';

          if (isNaN(hours) || hours <= 0) continue;
          
          let allocId = alloc?.id || '';
          if (!allocId) {
            const mapping = state.userMappings.find(m => 
              (m.userId === emp.id || !m.userId || m.userId === '') && 
              m.category2Ids?.includes(cat2.id) && 
              m.category3Ids?.includes(cat3.id)
            );
            if (mapping) allocId = mapping.allocationId;
          }

          const entryData = {
            id: 'e-' + Date.now() + Math.random().toString(36).substr(2, 5),
            employeeId: emp.id,
            projectId: proj.id,
            category2Id: cat2.id,
            category3Id: cat3.id,
            allocationId: allocId,
            date,
            hours,
            comment: comment || '',
            submittedAt: new Date().toISOString()
          };

          // 允许“多条信息完全一致”的记录：不做去重，全部写入数据库并在表格中一一展示
          newEntries.push(entryData);
        }
      }
    } catch (e) {
      console.error('Submit Error:', e);
      alert(language === 'zh' ? '提交时发生错误，请检查控制台' : 'An error occurred during submission. Check console.');
      return;
    }

    if (newEntries.length === 0) {
      alert(language === 'zh' ? '没有有效的填报数据' : 'No valid entries');
      return;
    }

    // 为了支持“易修改”：以当前表格为准，替换所选人员在当前日期范围内的记录
    const otherEntries = state.entries.filter(e => {
      const isWithinRange = e.date >= startDate && e.date <= endDate;
      const isSubmittedEmp = selectedEmployees.some(emp => emp.id === e.employeeId);
      return !(isWithinRange && isSubmittedEmp);
    });

    const newState = { ...state, entries: [...newEntries, ...otherEntries] };

    if (state.currentUser) {
      const firstFilledRow = data.find((row: any[]) => row && (String(row[1] || '').trim() || String(row[2] || '').trim() || String(row[3] || '').trim() || String(row[4] || '').trim()));
      if (firstFilledRow) {
        updateUserSetting(`${state.currentUser.id}_lastSelection`, {
          projectName: norm(firstFilledRow[1]),
          cat2Name: norm(firstFilledRow[2]),
          cat3Name: norm(firstFilledRow[3]),
          allocName: norm(firstFilledRow[4]),
        });
      }
    }

    // ★ Fix: 必须在 setState 之前设为 true，否则 React 会在 .then() 异步回调之前
    //   同步处理 setState → 触发 useEffect → 调用 saveDraft，导致提交后仍存有草稿
    isResettingRef.current = true;

    setState(newState);
    await handleSaveToDatabase(newState);
    setModifiedCells(new Set()); // Clear unsaved changes
    
    // 清除草稿，并从数据库重新拉取最新数据回填（保证提交后实时展示）
    if (state.currentUser) {
      const draftKey = `tms_draft_${state.currentUser.id}`;
      // 立即清除 IndexedDB 草稿（不等待 .then）
      await del(draftKey).catch(e => console.warn('Failed to clear draft:', e));
    }

    // 关键：保存后立刻从文件句柄重新读取（共享盘多人同时填报时也能刷新到最新）
    if (newState.fileHandle) {
      try {
        await loadDatabaseFromHandle(newState.fileHandle);
      } catch (e) {
        console.warn('Failed to reload database after submit:', e);
      }
    }

    hasUserEditedRef.current = false;
    isResettingRef.current = false;
    alert(language === 'zh' ? '批量填报成功' : 'Batch entry successful');
  };

  const handleJumpToNextError = () => {
    const hotInstance = hotRef.current?.hotInstance;
    if (!hotInstance) return;

    const data = hotInstance.getData();
    const selected = hotInstance.getSelected(); // [[row1, col1, row2, col2], ...]
    let startRow = 0;
    let startCol = 0;

    if (selected && selected.length > 0) {
      startRow = selected[0][0];
      startCol = selected[0][1];
    }

    // Flatten logic: scan rows, then columns.
    // We want to find the next error after (startRow, startCol)
    for (let r = 0; r < data.length; r++) {
      // Loop through all rows, but start checking from startRow
      const rowIdx = (startRow + r) % data.length;
      const row = data[rowIdx];
      
      const empName = row[0];
      const projName = row[1];
      const cat2Name = row[2];
      const cat3Name = row[3];
      const allocName = row[4];

      // Check if this row has any hours
      let hasHours = false;
      for (let i = 0; i < days.length; i++) {
        if (parseFloat(row[5 + i]) > 0) {
          hasHours = true;
          break;
        }
      }

      if (!hasHours) continue;

      // Validation logic (simplified)
      const isGeneral = projName === 'General time';
      
      // If we are on the startRow, we only want to check columns *after* the current selection if we've already checked the basics.
      // But actually, it's easier to just find the FIRST missing field in the NEXT row that has hours,
      // OR the NEXT missing field in the CURRENT row.

      const fields = [
        { val: empName, col: 0, msg: 'Name' },
        { val: projName, col: 1, msg: 'Project' }
      ];
      if (!isGeneral) {
        fields.push({ val: cat2Name, col: 2, msg: 'Category 2' });
        fields.push({ val: cat3Name, col: 3, msg: 'Category 3' });
        fields.push({ val: allocName, col: 4, msg: 'Allocation' });
      }

      for (const field of fields) {
        // If we are in the startRow, skip fields before or at the current startCol
        if (rowIdx === startRow && field.col <= startCol) continue;

        if (!field.val) {
          hotInstance.selectCell(rowIdx, field.col);
          return;
        }
      }

      // If we found no missing basic fields in this row, but it's not the startRow, 
      // we might have skipped the startRow's subsequent fields. 
      // The logic above handles "next field in current row" and "first field in next rows".
    }

    // If we reach here, we might need to check the startRow's fields from the beginning if we started in the middle
    for (const field of [
      { val: data[startRow][0], col: 0 },
      { val: data[startRow][1], col: 1 },
      { val: data[startRow][2], col: 2 },
      { val: data[startRow][3], col: 3 },
      { val: data[startRow][4], col: 4 }
    ]) {
      const isGeneral = data[startRow][1] === 'General time';
      if (isGeneral && field.col >= 2) continue;
      
      if (!field.val && field.col <= startCol) {
        // This is an error *before* the current selection in the same row
        // (This case is covered by the modulo loop above, but let's be safe)
        // Actually, the modulo loop already covers all rows.
      }
    }
    
    alert(language === 'zh' ? '未发现更多缺失信息' : 'No more missing information found');
  };

  const handleAddRow = () => {
    setHotData(prev => [...prev, ['', '', '', '', '', ...Array(days.length).fill('')]]);
  };

  return (
      <div className="bg-white p-4 rounded-[32px] shadow-sm border border-slate-200 space-y-4 w-full min-w-0 custom-scrollbar">
      <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100 min-w-fit space-y-3">
        {/* Row 1: 提交 & 日期 */}
        <div className="flex flex-wrap items-center gap-2 p-2 rounded-2xl bg-indigo-50 border border-indigo-200">
          <button 
            onClick={handleSubmit} 
            className={cn(
              "flex items-center justify-center gap-2 px-5 py-2.5 rounded-xl font-black text-[11px] shadow-md transition-all whitespace-nowrap",
              modifiedCells.size > 0 
                ? "bg-blue-600 text-white shadow-blue-200 hover:bg-blue-700" 
                : "bg-slate-300 text-slate-500 cursor-not-allowed"
            )}
            disabled={modifiedCells.size === 0}
          >
            <CheckCircle2 size={14} />
            {language === 'zh' ? '提交工时（表格）' : 'Submit Timesheet'}
          </button>

          <button 
            onClick={handleJumpToNextError}
            className="flex items-center justify-center gap-2 bg-amber-500 text-white px-5 py-2.5 rounded-xl font-black text-[11px] shadow-md shadow-amber-200 hover:bg-amber-600 transition-all whitespace-nowrap"
          >
            <ArrowRight size={14} />
            {language === 'zh' ? '跳转缺项' : 'Jump Missing'}
          </button>
          
          <div className="flex items-center gap-2 bg-white px-2 py-1.5 rounded-lg border border-slate-200 shadow-sm">
            <button
              onClick={() => {
                const d = addMonths(parseISO(startDate), -1);
                setStartDate(format(startOfMonth(d), 'yyyy-MM-dd'));
                setEndDate(format(endOfMonth(d), 'yyyy-MM-dd'));
              }}
              className="px-2 py-1 rounded-lg text-[10px] font-black text-slate-600 hover:bg-slate-50"
            >
              {language === 'zh' ? '上月' : 'Prev'}
            </button>
            <input
              type="month"
              value={monthValue}
              onChange={e => setMonthRange(e.target.value)}
              className="bg-transparent text-[10px] font-black text-slate-600 outline-none"
            />
            <button
              onClick={() => {
                const d = addMonths(parseISO(startDate), 1);
                setStartDate(format(startOfMonth(d), 'yyyy-MM-dd'));
                setEndDate(format(endOfMonth(d), 'yyyy-MM-dd'));
              }}
              className="px-2 py-1 rounded-lg text-[10px] font-black text-slate-600 hover:bg-slate-50"
            >
              {language === 'zh' ? '下月' : 'Next'}
            </button>
            <div className="text-[10px] font-bold text-slate-400 ml-1 whitespace-nowrap">
              {startDate}~{endDate}
            </div>
          </div>
        </div>

        {/* Row 2: 加班/请假调整（独立区域，避免与“提交工时”混淆） */}
        <div className="flex flex-wrap items-center gap-2 p-3 rounded-2xl border border-slate-200 bg-transparent">
          <div className="flex items-center gap-2 mr-1">
            <div className="text-[10px] font-black text-slate-700 uppercase tracking-widest">
              {language === 'zh' ? '加班/请假调整' : 'OT/Leave'}
            </div>
            <div className="text-[10px] font-bold text-slate-500">
              {language === 'zh' ? '（不等于工时提交）' : '(Not timesheet submit)'}
            </div>
          </div>
          <select 
            value={quickAdj.employeeId}
            onChange={e => setQuickAdj(prev => ({ ...prev, employeeId: e.target.value }))}
            className="bg-white border border-slate-200 rounded-lg px-2 py-1 text-[10px] font-bold outline-none w-32"
          >
            <option value="">{language === 'zh' ? '选择人员' : 'Select Emp'}</option>
            {visibleEmployees.filter(e => e.id !== 'admin-1').map(e => (
              <option key={e.id} value={e.id}>{getEmpName(e, language)}</option>
            ))}
          </select>
          <input 
            type="date" 
            value={quickAdj.date}
            onChange={e => setQuickAdj(prev => ({ ...prev, date: e.target.value }))}
            className="bg-white border border-slate-200 rounded-lg px-2 py-1 text-[10px] font-bold outline-none"
          />
          <div className="flex items-center gap-1">
            <input 
              type="number" 
              placeholder={language === 'zh' ? '加班' : 'OT'}
              value={quickAdj.overtime}
              onChange={e => setQuickAdj(prev => ({ ...prev, overtime: e.target.value }))}
              className="bg-white border border-slate-200 rounded-lg px-2 py-1 text-[10px] font-bold outline-none w-16"
            />
            <input 
              type="number" 
              placeholder={language === 'zh' ? '请假' : 'LV'}
              value={quickAdj.leave}
              onChange={e => setQuickAdj(prev => ({ ...prev, leave: e.target.value }))}
              className="bg-white border border-slate-200 rounded-lg px-2 py-1 text-[10px] font-bold outline-none w-16"
            />
          </div>
          <button 
            onClick={handleSaveAdj}
            className="bg-slate-900 text-white px-4 py-2 rounded-xl font-black text-[10px] tracking-widest hover:bg-slate-800 transition-all"
            title={language === 'zh' ? '保存加班/请假调整（影响当日上限 8 + 加班 - 请假）' : 'Save overtime/leave adjustment'}
          >
            {language === 'zh' ? '加班/请假保存' : 'Save OT/LV'}
          </button>
        </div>

        {/* Row 3: 人员勾选（查看范围） */}
        <div className="flex items-center gap-2 overflow-x-auto scrollbar-hide">
          <div className="flex flex-wrap gap-1">
            {visibleEmployees.filter(emp => emp.id !== 'admin-1').map(emp => (
              <label key={emp.id} className="flex items-center gap-1 bg-white px-1.5 py-0.5 rounded-md border border-slate-200 cursor-pointer hover:border-blue-300 hover:bg-blue-50/30 transition-all group whitespace-nowrap">
                <input 
                  type="checkbox" 
                  checked={selectedEmployeeIds.includes(emp.id)}
                  onChange={e => {
                    let next: string[];
                    if (e.target.checked) {
                      next = Array.from(new Set([...selectedEmployeeIds, emp.id]));
                    } else {
                      next = selectedEmployeeIds.filter(id => id !== emp.id);
                    }
                    setSelectedEmployeeIds(next);
                    if (state.currentUser) updateUserSetting(`${state.currentUser.id}_selectedEmployeeIds`, next);
                  }}
                  className="w-2.5 h-2.5 rounded border-slate-300 text-blue-600 focus:ring-blue-500 transition-all"
                />
                <span className="text-[9px] font-bold text-slate-500 group-hover:text-blue-700">{getEmpName(emp, language)}</span>
              </label>
            ))}
          </div>
        </div>
      </div>

      {/* 关键：父容器必须允许收缩(min-w-0)，否则 Handsontable 的最小宽度会把容器“撑宽”，滚动条消失 */}
      <div className="rounded-2xl border border-slate-100 bg-white overflow-hidden w-full min-w-0">
        {/* 让 Handsontable 自己产生横向滚动；同时确保这里不会被内容撑宽 */}
        {/* 固定可视宽度（避免容器被表格内容撑宽导致滚动条消失）。大屏最多 1400px，小屏自动降到 100% */}
        <div className="w-full max-w-[1400px] min-w-0">
        <HotTable
          ref={hotRef}
          data={hotData}
          afterChange={(changes, source) => {
            if (changes && source !== 'loadData') {
              hasUserEditedRef.current = true;
              setModifiedCells(prev => {
                const next = new Set(prev);
                changes.forEach(([row, prop]) => {
                  next.add(`${row}-${prop}`);
                });
                return next;
              });

              setHotData(prev => {
                const newData = prev.map(row => [...row]);
                changes.forEach(([row, prop, oldVal, newVal]) => {
                  if (typeof prop === 'number') {
                    if (newData[row]) {
                      newData[row][prop] = newVal;

                      // 1. General 工时自动清空分工
                      if (prop === 1) { // 项目列变更
                        const proj = getProjectByName(newVal as string);
                        if (proj && proj.id === 'p-general') {
                          newData[row][4] = ''; // 清空分工
                        }
                      }
                      
                      // 2. 只有非 General 项目才尝试自动填充分工
                      if (prop === 2 || prop === 3 || prop === 0) {
                        const empName = newData[row][0];
                        const projName = newData[row][1];
                        const cat2Name = newData[row][2];
                        const cat3Name = newData[row][3];
                        
                        const proj = getProjectByName(projName);
                        if (proj && proj.id !== 'p-general') {
                          const emp = getEmployeeByName(empName);
                          const cat2 = getCat2ByName(cat2Name);
                          const cat3 = cat2 ? getCat3ByName(cat3Name, cat2.id) : null;

                          if (emp && cat2 && cat3) {
                            const mapping = state.userMappings.find(m => 
                              (m.userId === emp.id || !m.userId || m.userId === '') && 
                              m.category2Ids?.includes(cat2.id) && 
                              m.category3Ids?.includes(cat3.id)
                            );
                            if (mapping) {
                              const alloc = state.allocations.find(a => a.id === mapping.allocationId);
                              if (alloc) {
                                newData[row][4] = language === 'en' ? (alloc.nameEn || alloc.nameZh) : (alloc.nameZh || alloc.nameEn);
                              }
                            }
                          }
                        }
                      }
                    }
                  }
                });
                saveDraft(newData);
                return newData;
              });
            }
          }}
          colHeaders={[
            'Employee',
            'Project',
            'Category 2',
            'Category 3',
            'Allocation',
            ...days.map(d => format(d, 'MM/dd') + ` (${format(d, 'eee')})`)
          ]}
          columns={tableColumns}
          colWidths={(index: number) => {
            if (index === 0) return colWidths.emp;
            if (index === 1) return colWidths.proj;
            if (index === 2) return colWidths.cat2;
            if (index === 3) return colWidths.cat3;
            if (index === 4) return colWidths.alloc;
            return colWidths.day;
          }}
          cell={cellComments}
          cells={function(row, col) {
            const cellProperties: any = {};
            const hotInstance = (this as any).instance;
            const rowData = hotInstance.getSourceDataAtRow(row);
            if (!rowData) return cellProperties;

            // 权限：若该行属于“仅查看不可编辑”的人员，则整行只读（仍可查看）
            const rowEmpName = rowData[0];
            if (rowEmpName) {
              const emp = getEmployeeByName(rowEmpName);
              if (emp && !editableIdSet.has(emp.id)) {
                cellProperties.readOnly = true;
              }
            }
            
            // 1. 背景色区分：人员固定颜色
            if (col === 0) {
              const empName = rowData[0];
              if (empName) {
                const emp = getEmployeeByName(empName);
                if (emp) {
                  // 根据 ID 生成固定的淡色背景
                  const colors = ['#E0F2FE', '#F3E8FF', '#DCFCE7', '#FFEDD5', '#FEE2E2', '#FCE7F3', '#CCFBF1', '#FEF9C3', '#F1F5F9'];
                  const colorIdx = emp.id.split('').reduce((acc: number, char: string) => acc + char.charCodeAt(0), 0) % colors.length;
                  cellProperties.renderer = function(instance: any, td: any, row: any, col: any, prop: any, value: any, cellProperties: any) {
                    Handsontable.renderers.TextRenderer.apply(this, arguments as any);
                    td.style.backgroundColor = colors[colorIdx];
                    td.style.fontWeight = '900';
                    td.style.color = '#0F172A';
                  };
                }
              }
            }

            // 2. 背景色区分：工作日/休息日
            if (col >= 5) {
              const date = days[col - 5];
              if (date && isWeekend(date)) {
                cellProperties.className = (cellProperties.className || '') + ' weekend-cell';
              }
            }

            // 2.1 工时填报：按“工时投入% = 当天已填总工时 / 应出勤工时”自动上色
            if (col >= 5) {
              const empName = rowData[0];
              const emp = empName ? getEmployeeByName(empName) : null;
              const date = days[col - 5] ? format(days[col - 5], 'yyyy-MM-dd') : '';
              const total = emp?.id && date ? (employeeDailyTotals[emp.id]?.[date] || 0) : 0;
              const expected = emp?.id && date ? (employeeDailyExpectedHours[emp.id]?.[date] || 0) : 0;
              const ratio = expected > 0 ? total / expected : 0;
              const raw = rowData[col];
              const cellHours = typeof raw === 'number' ? raw : parseFloat(String(raw ?? ''));
              if (!isNaN(cellHours) && cellHours > 0) {
                let cls = '';
                if (ratio >= 0.9) cls = 'hours-max';
                else if (ratio >= 0.85) cls = 'hours-high';
                else if (ratio >= 0.7) cls = 'hours-mid';
                else if (ratio >= 0.5) cls = 'hours-midlow';
                else cls = 'hours-low';
                cellProperties.className = (cellProperties.className || '') + ` ${cls}`;
              }
            }

            // 3. General 工时分工不可填
            if (col === 4) {
              const projName = rowData[1];
              const proj = getProjectByName(projName);
              if (proj && proj.id === 'p-general') {
                cellProperties.readOnly = true;
                cellProperties.renderer = function(instance: any, td: any, row: any, col: any, prop: any, value: any, cellProperties: any) {
                  Handsontable.renderers.DropdownRenderer.apply(this, arguments as any);
                  td.style.backgroundColor = '#F1F5F9';
                  td.style.color = '#94A3B8';
                };
              }
            }

            // 4. Unsaved changes indicator
            if (modifiedCells.has(`${row}-${col}`)) {
              cellProperties.className = (cellProperties.className || '') + ' unsaved-cell';
            }

            return cellProperties;
          }}
          afterGetColHeader={(col: number, TH: HTMLElement) => {
            try {
              if (col >= 5) {
                const d = days[col - 5];
                if (d && isWeekend(d)) TH.classList.add('weekend-header');
              }
            } catch (e) {}
          }}
          comments={true}
          fixedColumnsLeft={5}
          // 冻结列错行通常由“自动行高 + 文本换行”引起（冻结区与主体区计算行高不一致）
          // 这里固定行高并关闭 autoRowSize，保证冻结区/非冻结区行高一致
          autoRowSize={false as any}
          rowHeights={34 as any}
          // 进一步避免虚拟渲染带来的冻结区/主体区微小错位：强制渲染全部行
          renderAllRows={true as any}
          viewportRowRenderingOffset={999 as any}
          rowHeaders={true}
          rowHeaderWidth={colWidths.rowHeader}
          height="auto"
          width="100%"
          stretchH="none" 
          contextMenu={true}
          manualColumnResize={true}
          licenseKey="non-commercial-and-evaluation"
          className="custom-hot"
        />
        </div>
      </div>

      <div className="flex items-center gap-2 text-[10px] font-bold text-slate-400">
        <Info size={12} />
        <span>{language === 'zh' ? '提示：支持复制粘贴、拖拽填充，多行填报。' : 'Tip: Supports copy-paste, drag-fill, multiple rows.'}</span>
      </div>
    </div>
  );
}

// --- 1. Entry View (Strict Controls) ---
function ProductionForecastView() {
  const { state, setState, language, handleSaveToDatabase, connectionType, apiBase } = useContext(AppContext)!;
  const projects = state.projects.filter(p => p.id !== 'p-general' && p.id !== 'p-leave');
  const [selectedProjectId, setSelectedProjectId] = useState<string>(projects[0]?.id || '');
  const [selectedBomType, setSelectedBomType] = useState<MaterialType>('Mechanical');
  const [horizonWeeks, setHorizonWeeks] = useState<number>(12);

  const [newBom, setNewBom] = useState({
    equipmentId: '',
    equipmentName: '',
    partCode: '',
    partName: '',
    quantity: '1',
    assemblyHours: '8',
    isCritical: true,
    priority: '50',
    targetDate: format(addWeeks(new Date(), 8), 'yyyy-MM-dd')
  });
  const [newEta, setNewEta] = useState({
    partCode: '',
    etaDate: format(addWeeks(new Date(), 2), 'yyyy-MM-dd'),
    status: 'OnTrack' as MaterialETAPlan['status']
  });
  const [newCapacity, setNewCapacity] = useState({
    lineName: '',
    weeklyHours: '320',
    efficiency: '0.85'
  });

  useEffect(() => {
    if (!selectedProjectId && projects.length > 0) {
      setSelectedProjectId(projects[0].id);
    }
  }, [projects, selectedProjectId]);

  const filteredBoms = useMemo(
    () => state.assemblyBoms.filter(x => x.projectId === selectedProjectId && x.bomType === selectedBomType),
    [state.assemblyBoms, selectedProjectId, selectedBomType]
  );
  const filteredEta = useMemo(
    () => state.materialEtaPlans.filter(x => x.projectId === selectedProjectId),
    [state.materialEtaPlans, selectedProjectId]
  );
  const filteredCapacities = useMemo(
    () => state.productionCapacities.filter(x => x.bomType === selectedBomType),
    [state.productionCapacities, selectedBomType]
  );

  const capacityPerWeek = useMemo(() => {
    const total = filteredCapacities.reduce((sum, c) => sum + c.weeklyHours * Math.max(0, Math.min(1, c.efficiency)), 0);
    return Math.max(80, Math.round(total || 280));
  }, [filteredCapacities]);

  const equipmentForecastPreview = useMemo(() => {
    if (!selectedProjectId) return [] as Omit<ProductionForecastItem, 'id' | 'generatedAt'>[];
    const today = new Date();
    const horizonStart = startOfWeek(today, { weekStartsOn: 1 });
    const equipmentMap = new Map<string, AssemblyBOMItem[]>();
    filteredBoms.forEach(item => {
      if (!equipmentMap.has(item.equipmentId)) equipmentMap.set(item.equipmentId, []);
      equipmentMap.get(item.equipmentId)!.push(item);
    });

    const candidates = Array.from(equipmentMap.entries()).map(([equipmentId, items]) => {
      const equipmentName = items[0]?.equipmentName || equipmentId;
      const partCodes = Array.from(new Set(items.map(i => i.partCode)));
      const criticalPartCodes = Array.from(new Set(items.filter(i => i.isCritical).map(i => i.partCode)));
      const priority = Math.min(...items.map(i => i.priority || 50));
      const requiredHours = Math.max(1, items.reduce((sum, i) => sum + (Number(i.assemblyHours) || 0), 0));
      const targetDate = items.map(i => i.targetDate).find(Boolean);

      const partEtas = partCodes.map(code => {
        const plans = filteredEta.filter(e => e.partCode === code);
        if (!plans.length) return { code, eta: null as Date | null, status: 'Missing' as const };
        const earliest = plans
          .map(p => parseISO(p.etaDate))
          .sort((a, b) => a.getTime() - b.getTime())[0];
        const risky = plans.some(p => p.status === 'Delayed');
        return { code, eta: earliest, status: risky ? 'Delayed' as const : 'Ok' as const };
      });

      const readinessReady = partEtas.filter(p => p.eta && p.eta.getTime() <= today.getTime()).length;
      const readiness = partCodes.length > 0 ? readinessReady / partCodes.length : 0;
      const blockedParts = partEtas.filter(p => !p.eta || p.status === 'Delayed').map(p => p.code);
      const focusParts = criticalPartCodes.length > 0 ? criticalPartCodes : partCodes;
      const focusEtas = partEtas.filter(p => focusParts.includes(p.code)).map(p => p.eta).filter(Boolean) as Date[];
      const earliestReady = focusEtas.length ? new Date(Math.max(...focusEtas.map(d => d.getTime()))) : addWeeks(today, 12);
      const risk: 'Low' | 'Medium' | 'High' =
        blockedParts.length > 0
          ? 'High'
          : earliestReady.getTime() > addWeeks(today, 4).getTime()
            ? 'Medium'
            : 'Low';

      return {
        projectId: selectedProjectId,
        equipmentId,
        equipmentName,
        bomType: selectedBomType,
        earliestReady,
        requiredHours,
        readiness,
        blockedParts,
        risk,
        priority,
        targetDate
      };
    });

    candidates.sort((a, b) => a.priority - b.priority || a.earliestReady.getTime() - b.earliestReady.getTime());

    let cursor = horizonStart;
    return candidates.map(c => {
      const earliestStart = startOfWeek(c.earliestReady, { weekStartsOn: 1 });
      const plannedStart = c.earliestReady.getTime() > cursor.getTime() ? earliestStart : cursor;
      const durationWeeks = Math.max(1, Math.ceil(c.requiredHours / capacityPerWeek));
      const plannedEnd = addDays(addWeeks(plannedStart, durationWeeks), -1);
      cursor = addDays(plannedEnd, 1);
      return {
        projectId: c.projectId,
        equipmentId: c.equipmentId,
        equipmentName: c.equipmentName,
        bomType: c.bomType,
        earliestStart: format(earliestStart, 'yyyy-MM-dd'),
        plannedStart: format(plannedStart, 'yyyy-MM-dd'),
        plannedEnd: format(plannedEnd, 'yyyy-MM-dd'),
        requiredHours: c.requiredHours,
        capacityPerWeek,
        readiness: Number((c.readiness * 100).toFixed(0)),
        risk: c.risk,
        blockedParts: c.blockedParts
      };
    });
  }, [selectedProjectId, selectedBomType, filteredBoms, filteredEta, capacityPerWeek]);

  const timelineWeeks = useMemo(() => {
    const start = startOfWeek(new Date(), { weekStartsOn: 1 });
    return Array.from({ length: horizonWeeks }).map((_, i) => {
      const ws = addWeeks(start, i);
      return { idx: i, start: ws, label: format(ws, 'MM-dd') };
    });
  }, [horizonWeeks]);

  const saveForecast = () => {
    const generatedAt = new Date().toISOString();
    const nextForecasts: ProductionForecastItem[] = equipmentForecastPreview.map((f, idx) => ({
      ...f,
      id: `pf-${Date.now()}-${idx}`,
      generatedAt
    }));
    setState(prev => {
      const remaining = prev.productionForecasts.filter(x => !(x.projectId === selectedProjectId && x.bomType === selectedBomType));
      const newState = { ...prev, productionForecasts: [...nextForecasts, ...remaining] };
      handleSaveToDatabase(newState);
      return newState;
    });
    alert(language === 'zh' ? '预测排班已保存' : 'Forecast schedule saved');
  };

  const addBomItem = () => {
    if (!selectedProjectId || !newBom.equipmentId || !newBom.partCode) return;
    const item: AssemblyBOMItem = {
      id: `ab-${Date.now()}`,
      projectId: selectedProjectId,
      equipmentId: newBom.equipmentId.trim(),
      equipmentName: newBom.equipmentName.trim() || newBom.equipmentId.trim(),
      bomType: selectedBomType,
      partCode: newBom.partCode.trim(),
      partName: newBom.partName.trim() || newBom.partCode.trim(),
      quantity: Math.max(1, Number(newBom.quantity) || 1),
      assemblyHours: Math.max(1, Number(newBom.assemblyHours) || 8),
      isCritical: newBom.isCritical,
      priority: Math.max(1, Number(newBom.priority) || 50),
      targetDate: newBom.targetDate
    };
    setState(prev => {
      const newState = { ...prev, assemblyBoms: [item, ...prev.assemblyBoms] };
      handleSaveToDatabase(newState);
      return newState;
    });
  };

  const addEtaPlan = () => {
    if (!selectedProjectId || !newEta.partCode) return;
    const eta: MaterialETAPlan = {
      id: `eta-${Date.now()}`,
      projectId: selectedProjectId,
      partCode: newEta.partCode.trim(),
      etaDate: newEta.etaDate,
      status: newEta.status
    };
    setState(prev => {
      const newState = { ...prev, materialEtaPlans: [eta, ...prev.materialEtaPlans] };
      handleSaveToDatabase(newState);
      return newState;
    });
  };

  const addCapacityPlan = () => {
    if (!newCapacity.lineName) return;
    const cap: ProductionCapacityPlan = {
      id: `cap-${Date.now()}`,
      bomType: selectedBomType,
      lineName: newCapacity.lineName.trim(),
      weeklyHours: Math.max(40, Number(newCapacity.weeklyHours) || 320),
      efficiency: Math.max(0.3, Math.min(1, Number(newCapacity.efficiency) || 0.85))
    };
    setState(prev => {
      const newState = { ...prev, productionCapacities: [cap, ...prev.productionCapacities] };
      handleSaveToDatabase(newState);
      return newState;
    });
  };

  const endpointSpecs = [
    { method: 'GET', path: '/api/production/bom', desc: '获取项目装配BOM' },
    { method: 'POST', path: '/api/production/bom', desc: '新增或批量导入装配BOM' },
    { method: 'GET', path: '/api/material/eta', desc: '获取物料到货计划与风险' },
    { method: 'POST', path: '/api/material/eta', desc: '写入到货预测或供应商承诺' },
    { method: 'GET', path: '/api/production/capacity', desc: '获取产线周产能' },
    { method: 'POST', path: '/api/production/forecast/generate', desc: '触发12周预测排班计算' },
    { method: 'GET', path: '/api/production/forecast', desc: '获取预测排班结果与风险解释' }
  ];

  return (
    <div className="space-y-6 w-full max-w-full pb-10">
      <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-6 space-y-4">
        <div className="flex flex-wrap items-center gap-2">
          <select value={selectedProjectId} onChange={e => setSelectedProjectId(e.target.value)} className="input-field py-2 px-3 h-10 text-xs font-bold min-w-[220px]">
            {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
          {(['Mechanical', 'Electrical', 'Standard'] as MaterialType[]).map(type => (
            <button key={type} onClick={() => setSelectedBomType(type)} className={cn("px-3 py-1.5 rounded-xl text-xs font-black border", selectedBomType === type ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200")}>
              {type === 'Mechanical' ? (language === 'zh' ? '机械BOM' : 'Mechanical BOM') : type === 'Electrical' ? (language === 'zh' ? '电气BOM' : 'Electrical BOM') : (language === 'zh' ? '标准件BOM' : 'Standard BOM')}
            </button>
          ))}
          <input type="number" min={4} max={16} value={horizonWeeks} onChange={e => setHorizonWeeks(Math.max(4, Math.min(16, Number(e.target.value) || 12)))} className="input-field py-2 px-3 h-10 text-xs font-bold w-24" />
          <span className="text-xs font-bold text-slate-400">{language === 'zh' ? '预测周数' : 'Horizon(weeks)'}</span>
          <button onClick={saveForecast} className="bg-blue-600 text-white px-4 py-2 rounded-xl text-xs font-black flex items-center gap-2">
            <Save size={12} /> {language === 'zh' ? '生成并保存预测' : 'Generate & Save'}
          </button>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
          <div className="rounded-2xl border border-slate-100 p-4 bg-slate-50/60">
            <p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? 'BOM条目' : 'BOM Items'}</p>
            <p className="text-2xl font-black text-slate-900">{filteredBoms.length}</p>
          </div>
          <div className="rounded-2xl border border-slate-100 p-4 bg-slate-50/60">
            <p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? '到货计划' : 'ETA Plans'}</p>
            <p className="text-2xl font-black text-slate-900">{filteredEta.length}</p>
          </div>
          <div className="rounded-2xl border border-slate-100 p-4 bg-slate-50/60">
            <p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? '周产能(小时)' : 'Weekly Capacity(h)'}</p>
            <p className="text-2xl font-black text-slate-900">{capacityPerWeek}</p>
          </div>
          <div className="rounded-2xl border border-slate-100 p-4 bg-slate-50/60">
            <p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? '预测设备数' : 'Predicted Equipments'}</p>
            <p className="text-2xl font-black text-slate-900">{equipmentForecastPreview.length}</p>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-4 space-y-2">
          <h4 className="text-sm font-black">{language === 'zh' ? '新增装配BOM' : 'Add Assembly BOM'}</h4>
          <input value={newBom.equipmentId} onChange={e => setNewBom(prev => ({ ...prev, equipmentId: e.target.value }))} placeholder={language === 'zh' ? '设备ID' : 'Equipment ID'} className="input-field h-9 py-1 px-2 text-xs" />
          <input value={newBom.equipmentName} onChange={e => setNewBom(prev => ({ ...prev, equipmentName: e.target.value }))} placeholder={language === 'zh' ? '设备名称' : 'Equipment Name'} className="input-field h-9 py-1 px-2 text-xs" />
          <input value={newBom.partCode} onChange={e => setNewBom(prev => ({ ...prev, partCode: e.target.value }))} placeholder={language === 'zh' ? '零件编码' : 'Part Code'} className="input-field h-9 py-1 px-2 text-xs" />
          <input value={newBom.partName} onChange={e => setNewBom(prev => ({ ...prev, partName: e.target.value }))} placeholder={language === 'zh' ? '零件名称' : 'Part Name'} className="input-field h-9 py-1 px-2 text-xs" />
          <div className="grid grid-cols-3 gap-2">
            <input value={newBom.quantity} onChange={e => setNewBom(prev => ({ ...prev, quantity: e.target.value }))} placeholder={language === 'zh' ? '数量' : 'Qty'} className="input-field h-9 py-1 px-2 text-xs" />
            <input value={newBom.assemblyHours} onChange={e => setNewBom(prev => ({ ...prev, assemblyHours: e.target.value }))} placeholder={language === 'zh' ? '工时' : 'Hours'} className="input-field h-9 py-1 px-2 text-xs" />
            <input value={newBom.priority} onChange={e => setNewBom(prev => ({ ...prev, priority: e.target.value }))} placeholder={language === 'zh' ? '优先级' : 'Priority'} className="input-field h-9 py-1 px-2 text-xs" />
          </div>
          <button onClick={addBomItem} className="btn-primary py-2 px-3 text-xs"><Plus size={12} /> {language === 'zh' ? '添加BOM行' : 'Add BOM Row'}</button>
        </div>
        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-4 space-y-2">
          <h4 className="text-sm font-black">{language === 'zh' ? '新增到货计划' : 'Add ETA Plan'}</h4>
          <input value={newEta.partCode} onChange={e => setNewEta(prev => ({ ...prev, partCode: e.target.value }))} placeholder={language === 'zh' ? '零件编码' : 'Part Code'} className="input-field h-9 py-1 px-2 text-xs" />
          <input type="date" value={newEta.etaDate} onChange={e => setNewEta(prev => ({ ...prev, etaDate: e.target.value }))} className="input-field h-9 py-1 px-2 text-xs" />
          <select value={newEta.status} onChange={e => setNewEta(prev => ({ ...prev, status: e.target.value as MaterialETAPlan['status'] }))} className="input-field h-9 py-1 px-2 text-xs">
            <option value="OnTrack">OnTrack</option>
            <option value="Risky">Risky</option>
            <option value="Delayed">Delayed</option>
          </select>
          <button onClick={addEtaPlan} className="btn-primary py-2 px-3 text-xs"><Plus size={12} /> {language === 'zh' ? '添加ETA' : 'Add ETA'}</button>
        </div>
        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-4 space-y-2">
          <h4 className="text-sm font-black">{language === 'zh' ? '产能配置' : 'Capacity Setup'}</h4>
          <input value={newCapacity.lineName} onChange={e => setNewCapacity(prev => ({ ...prev, lineName: e.target.value }))} placeholder={language === 'zh' ? '产线名称' : 'Line Name'} className="input-field h-9 py-1 px-2 text-xs" />
          <input value={newCapacity.weeklyHours} onChange={e => setNewCapacity(prev => ({ ...prev, weeklyHours: e.target.value }))} placeholder={language === 'zh' ? '周工时' : 'Weekly Hours'} className="input-field h-9 py-1 px-2 text-xs" />
          <input value={newCapacity.efficiency} onChange={e => setNewCapacity(prev => ({ ...prev, efficiency: e.target.value }))} placeholder={language === 'zh' ? '效率(0-1)' : 'Efficiency'} className="input-field h-9 py-1 px-2 text-xs" />
          <button onClick={addCapacityPlan} className="btn-primary py-2 px-3 text-xs"><Plus size={12} /> {language === 'zh' ? '添加产能' : 'Add Capacity'}</button>
        </div>
      </div>

      <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-4 overflow-auto">
        <h4 className="text-sm font-black mb-3">{language === 'zh' ? '未来排班预测（12周粗排）' : '12-Week Forecast Scheduling'}</h4>
        <table className="w-full text-xs border-collapse">
          <thead>
            <tr className="bg-slate-50">
              <th className="p-2 text-left">Equipment</th>
              <th className="p-2 text-left">{language === 'zh' ? '最早可开工' : 'Earliest Start'}</th>
              <th className="p-2 text-left">{language === 'zh' ? '计划开工' : 'Planned Start'}</th>
              <th className="p-2 text-left">{language === 'zh' ? '计划完工' : 'Planned End'}</th>
              <th className="p-2 text-left">{language === 'zh' ? '需求工时' : 'Req.Hours'}</th>
              <th className="p-2 text-left">{language === 'zh' ? '齐套率' : 'Readiness'}</th>
              <th className="p-2 text-left">{language === 'zh' ? '风险' : 'Risk'}</th>
            </tr>
          </thead>
          <tbody>
            {equipmentForecastPreview.map(row => (
              <tr key={row.equipmentId} className="border-t border-slate-100">
                <td className="p-2 font-bold">{row.equipmentName}</td>
                <td className="p-2">{row.earliestStart}</td>
                <td className="p-2">{row.plannedStart}</td>
                <td className="p-2">{row.plannedEnd}</td>
                <td className="p-2">{Math.round(row.requiredHours)}</td>
                <td className="p-2">{row.readiness}%</td>
                <td className={cn("p-2 font-bold", row.risk === 'High' ? "text-rose-600" : row.risk === 'Medium' ? "text-amber-600" : "text-emerald-600")}>{row.risk}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="mt-4 space-y-2">
          {equipmentForecastPreview.map(task => {
            const s = parseISO(task.plannedStart);
            const e = parseISO(task.plannedEnd);
            return (
              <div key={`g-${task.equipmentId}`} className="flex items-center gap-2">
                <div className="w-40 text-[10px] font-bold text-slate-600 truncate">{task.equipmentName}</div>
                <div className="flex-1 grid gap-1" style={{ gridTemplateColumns: `repeat(${timelineWeeks.length}, minmax(20px, 1fr))` }}>
                  {timelineWeeks.map(w => {
                    const we = addDays(w.start, 6);
                    const overlap = s.getTime() <= we.getTime() && e.getTime() >= w.start.getTime();
                    return (
                      <div key={`${task.equipmentId}-${w.idx}`} className={cn("h-4 rounded-sm text-[8px] flex items-center justify-center", overlap ? "bg-indigo-500 text-white" : "bg-slate-100 text-slate-300")}>
                        {overlap ? '■' : ''}
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
          <div className="flex gap-1 pl-44">
            {timelineWeeks.map(w => (
              <span key={`wk-${w.idx}`} className="text-[8px] font-bold text-slate-400 min-w-[24px] text-center">{w.label}</span>
            ))}
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-4">
          <h4 className="text-sm font-black mb-3">{language === 'zh' ? '计算流程图（MVP）' : 'Calculation Flow (MVP)'}</h4>
          <div className="space-y-2 text-xs">
            <div className="p-3 rounded-xl bg-blue-50 font-bold">{language === 'zh' ? '1. 读取项目装配BOM' : '1. Load assembly BOM'}</div>
            <div className="text-center text-slate-300">↓</div>
            <div className="p-3 rounded-xl bg-indigo-50 font-bold">{language === 'zh' ? '2. 计算关键件齐套最早日期' : '2. Compute earliest kitting date'}</div>
            <div className="text-center text-slate-300">↓</div>
            <div className="p-3 rounded-xl bg-violet-50 font-bold">{language === 'zh' ? '3. 结合周产能做12周粗排' : '3. 12-week rough-cut capacity scheduling'}</div>
            <div className="text-center text-slate-300">↓</div>
            <div className="p-3 rounded-xl bg-emerald-50 font-bold">{language === 'zh' ? '4. 输出风险与计划开完工窗口' : '4. Output risk and start/end windows'}</div>
          </div>
        </div>
        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-4 overflow-auto">
          <h4 className="text-sm font-black mb-3">{language === 'zh' ? '接口清单（建议）' : 'API List (Recommended)'}</h4>
          <table className="w-full text-xs border-collapse">
            <thead>
              <tr className="bg-slate-50">
                <th className="p-2 text-left">Method</th>
                <th className="p-2 text-left">Path</th>
                <th className="p-2 text-left">{language === 'zh' ? '说明' : 'Description'}</th>
              </tr>
            </thead>
            <tbody>
              {endpointSpecs.map(api => (
                <tr key={api.path} className="border-t border-slate-100">
                  <td className="p-2 font-black text-blue-700">{api.method}</td>
                  <td className="p-2 font-mono">{api.path}</td>
                  <td className="p-2">{api.desc}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

export function ProductionForecastBoard() {
  const { state, setState, language, handleSaveToDatabase } = useContext(AppContext)!;
  const projects = state.projects.filter(p => p.id !== 'p-general' && p.id !== 'p-leave');
  const [selectedProjectId, setSelectedProjectId] = useState<string>(projects[0]?.id || '');
  const [selectedBomType, setSelectedBomType] = useState<'Mechanical' | 'Electrical'>('Mechanical');
  // Link to Device Model Assembly BOM (strongly related)
  const [selectedDeviceModelId, setSelectedDeviceModelId] = useState<string>((state.deviceModels?.[0]?.id) || '');
  const [selectedDeviceAssemblyTplId, setSelectedDeviceAssemblyTplId] = useState<string>('');
  const [plannedDeviceCount, setPlannedDeviceCount] = useState<number>(1);
  const [blockReason, setBlockReason] = useState<Record<string, string>>({});
  const [blockCategory, setBlockCategory] = useState<Record<string, string>>({});
  const [eventActionFilter, setEventActionFilter] = useState<'all' | 'refresh_readiness' | 'assembly_status_change'>('all');
  const [eventNodeFilter, setEventNodeFilter] = useState('');
  const [eventStatusFilter, setEventStatusFilter] = useState<'all' | 'NotReady' | 'ReadyToAssemble' | 'Assembling' | 'Done' | 'Blocked'>('all');
  const [eventDays, setEventDays] = useState(30);

  useEffect(() => {
    if (!selectedProjectId && projects.length > 0) setSelectedProjectId(projects[0].id);
  }, [projects, selectedProjectId]);

  useEffect(() => {
    if (!selectedDeviceModelId && (state.deviceModels || []).length > 0) {
      setSelectedDeviceModelId(state.deviceModels[0].id);
    }
  }, [state.deviceModels, selectedDeviceModelId]);

  const deviceAssemblyTplOptions = useMemo(() => {
    const dm = selectedDeviceModelId;
    if (!dm) return [];
    return (state.deviceBomTemplates || [])
      .filter((t: any) => t.deviceModelId === dm
        && t.bomType === selectedBomType
        && (t.templateKind || 'purchasing') === 'assembly')
      .sort((a: any, b: any) => String(a.version).localeCompare(String(b.version)))
      .map((t: any) => ({ id: t.id, label: `${t.version} · ${t.name || ''}`.trim() }));
  }, [state.deviceBomTemplates, selectedDeviceModelId, selectedBomType]);

  useEffect(() => {
    if (!deviceAssemblyTplOptions.length) {
      setSelectedDeviceAssemblyTplId('');
      return;
    }
    if (!deviceAssemblyTplOptions.some(x => x.id === selectedDeviceAssemblyTplId)) {
      setSelectedDeviceAssemblyTplId(deviceAssemblyTplOptions[0].id);
    }
  }, [deviceAssemblyTplOptions, selectedDeviceAssemblyTplId]);

  const expandAssemblyToPurchasingAgg = (nodes: any[]) => {
    const items = (nodes || []).filter(n => n.levelPath && n.partCodeModel && (Number(n.quantityPerDevice) || 0) > 0);
    const byPath = new Map<string, any>();
    items.forEach(n => byPath.set(String(n.levelPath), n));
    const hasChild = new Set<string>();
    items.forEach(n => { if (n.parentLevelPath) hasChild.add(String(n.parentLevelPath)); });
    const cum = new Map<string, number>();
    const getCum = (p: string): number => {
      if (cum.has(p)) return cum.get(p)!;
      const node = byPath.get(p);
      if (!node) return 1;
      const parent = String(node.parentLevelPath || '');
      const parentCum = parent ? getCum(parent) : 1;
      const local = Number(node.quantityPerDevice) || 0;
      const v = parentCum * local;
      cum.set(p, v);
      return v;
    };
    items.forEach(n => getCum(String(n.levelPath)));
    const isLeaf = (n: any) => {
      const fn = String(n.fileName || '').toLowerCase();
      if (fn.endsWith('.ipt')) return true;
      return !hasChild.has(String(n.levelPath));
    };
    const agg = new Map<string, { qty: number; sample: any }>();
    items.forEach(n => {
      if (!isLeaf(n)) return;
      const part = String(n.partCodeModel || '').trim();
      if (!part) return;
      const q = cum.get(String(n.levelPath)) || 0;
      if (q <= 0) return;
      const prev = agg.get(part);
      agg.set(part, { qty: (prev?.qty || 0) + q, sample: prev?.sample || n });
    });
    return { agg, byPath, cum, hasChild, items };
  };

  const generateProjectNeedsFromDeviceModel = async () => {
    if (!selectedProjectId) return alert('请先选择项目');
    if (!selectedDeviceModelId) return alert('请先选择设备型号');
    if (!selectedDeviceAssemblyTplId) return alert('请先选择装配BOM版本');
    const deviceCount = Math.max(1, Math.floor(Number(plannedDeviceCount) || 1));
    const nodes = (state.deviceBomLines || []).filter((l: any) => l.templateId === selectedDeviceAssemblyTplId);
    if (!nodes.length) return alert('该装配BOM版本没有明细');

    const tpl: any = (state.deviceBomTemplates || []).find((t: any) => t.id === selectedDeviceAssemblyTplId);
    const tag = `From DeviceModelBOM ${selectedDeviceModelId} ${tpl?.version || ''} ×${deviceCount}`;
    const overwrite = window.confirm(`将按“设备数量=${deviceCount}”展开装配BOM，并写入：\n- 项目装配BOM（用于排班预测）\n- 采购需求（用于齐套/到货跟踪）\n\n重复生成将覆盖同一设备型号的旧记录。\n\n继续吗？`);
    if (!overwrite) return;

    const { agg, items, hasChild, cum } = expandAssemblyToPurchasingAgg(nodes);

    // Build project assemblyBoms by top-level nodes (level depth=1), equipmentId uses DM prefix to avoid conflict
    const topNodes = items.filter((n: any) => !String(n.parentLevelPath || '').trim());
    const subtreePartsAgg = (rootPath: string) => {
      const map = new Map<string, { qty: number; sample: any }>();
      items.forEach((n: any) => {
        const lp = String(n.levelPath || '');
        if (!(lp === rootPath || lp.startsWith(rootPath + '.'))) return;
        const fn = String(n.fileName || '').toLowerCase();
        const leaf = fn.endsWith('.ipt') || !hasChild.has(lp);
        if (!leaf) return;
        const part = String(n.partCodeModel || '').trim();
        if (!part) return;
        const q = cum.get(lp) || 0;
        if (q <= 0) return;
        const prev = map.get(part);
        map.set(part, { qty: (prev?.qty || 0) + q, sample: prev?.sample || n });
      });
      return map;
    };

    const now = new Date().toISOString();
    const newAssemblyBoms: AssemblyBOMItem[] = [];
    topNodes.forEach((tn: any, idx: number) => {
      const root = String(tn.levelPath);
      const equipId = `${selectedDeviceModelId}.${root}`;
      const equipName = String(tn.name || tn.partCodeModel || root);
      const partMap = subtreePartsAgg(root);
      Array.from(partMap.entries()).forEach(([part, info]) => {
        newAssemblyBoms.push({
          id: `ab-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
          projectId: selectedProjectId,
          equipmentId: equipId,
          equipmentName: equipName,
          bomType: selectedBomType as any,
          partCode: part,
          partName: String(info.sample?.name || part),
          quantity: Number((info.qty * deviceCount).toFixed(6)),
          assemblyHours: 0,
          isCritical: false,
          priority: idx
        });
      });
    });

    // Build procurement requirements (aggregate overall)
    const newReqs: MaterialRequirement[] = Array.from(agg.entries()).map(([part, info]) => {
      const reqQty = Number((info.qty * deviceCount).toFixed(6));
      return {
        id: `req-bomdm-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        projectId: selectedProjectId,
        bomId: 'bom-default',
        creatorId: state.currentUser?.id || 'system',
        type: selectedBomType as any,
        code: `BOMDM-${Math.floor(Math.random() * 100000)}`,
        investmentOrder: '',
        csOrder: '',
        costCenter: '',
        name: String(info.sample?.name || part),
        model: part,
        drawingNo: '',
        revision: String(info.sample?.revision || 'NA'),
        brand: '',
        specs: String(info.sample?.specs || ''),
        quantity: reqQty,
        unit: String(info.sample?.unit || ''),
        unitPrice: 0,
        totalPrice: 0,
        supplier: String(info.sample?.supplier || ''),
        prNumber: '',
        prCreatedAt: '',
        poNumber: '',
        poCreatedAt: '',
        actualSupplier: '',
        prStatus: '',
        needByDate: '',
        deliveryTime: '',
        expectedArrival: format(new Date(), 'yyyy-MM-dd'),
        receivedAt: '',
        warehouseStatus: '',
        repairArrivalAt: '',
        inboundQuantity: 0,
        outboundQuantity: 0,
        inventoryQuantity: 0,
        userName: '',
        useDate: '',
        qualityFeedback: '',
        sourceTemplateVersion: `${tpl.version}${tpl.name ? ` · ${tpl.name}` : ''}`.trim(),
        changeReason: '从设备BOM导入',
        changeReasonNote: tag,
        urgency: 'Medium',
        status: 'Pending Review',
        comments: tag,
        version: 1,
        createdAt: now,
        updatedAt: now,
      } as any;
    }).filter(r => (Number((r as any).quantity) || 0) > 0);

    const newState = (() => {
      // overwrite previous generated records for this project+device model
      const keptBoms = (state.assemblyBoms || []).filter(b => !(b.projectId === selectedProjectId && String(b.equipmentId || '').startsWith(`${selectedDeviceModelId}.`) && b.bomType === (selectedBomType as any)));
      const keptReqs = (state.materialRequirements || []).filter(r => !(r.projectId === selectedProjectId && String((r as any).comments || '').includes(`From DeviceModelBOM ${selectedDeviceModelId}`)));
      return {
        ...state,
        assemblyBoms: [...newAssemblyBoms, ...keptBoms],
        materialRequirements: [...newReqs, ...keptReqs]
      } as AppState;
    })();

    setState(newState);
    await handleSaveToDatabase(newState);
    alert(`已生成：装配BOM明细 ${newAssemblyBoms.length} 条；采购需求 ${newReqs.length} 条。`);
  };

  const filteredBoms = useMemo(
    () => state.assemblyBoms.filter(x => x.projectId === selectedProjectId && x.bomType === selectedBomType),
    [state.assemblyBoms, selectedProjectId, selectedBomType]
  );
  const assemblySourceTypes = useMemo<MaterialType[]>(
    () => (selectedBomType === 'Mechanical' ? ['Mechanical', 'Standard'] : ['Electrical']),
    [selectedBomType]
  );
  const procurementDeliveryRows = useMemo(() => {
    return state.materialRequirements
      .filter(r => r.projectId === selectedProjectId && assemblySourceTypes.includes(r.type))
      .map(r => ({
        partCode: `${r.model || r.code || ''}`.trim(),
        etaDate: r.deliveryTime || r.expectedArrival || '',
        receivedAt: r.receivedAt || '',
        warehouseStatus: r.warehouseStatus || '',
        status: r.status
      }))
      .filter(r => !!r.partCode);
  }, [state.materialRequirements, selectedProjectId, assemblySourceTypes]);
  const existingStatusMap = useMemo(() => {
    const map = new Map<string, ProductionForecastItem>();
    state.productionForecasts
      .filter(x => x.projectId === selectedProjectId && x.bomType === selectedBomType)
      .forEach(x => map.set(x.equipmentId, x));
    return map;
  }, [state.productionForecasts, selectedProjectId, selectedBomType]);

  const safeDate = (s?: string) => {
    if (!s) return null;
    try {
      const d = parseISO(s);
      return isNaN(d.getTime()) ? null : d;
    } catch {
      return null;
    }
  };
  const getParentCode = (code: string) => {
    const idx = code.lastIndexOf('.');
    return idx > 0 ? code.slice(0, idx) : '';
  };
  const inferImportedBomType = (row: Record<string, any>, sheetName?: string): MaterialType | null => {
    const rawType = String(
      row['BOM分类'] ||
      row['分类'] ||
      row['物料分类'] ||
      row['Type'] ||
      row['type'] ||
      row['bomType'] ||
      row['Category'] ||
      sheetName ||
      ''
    ).trim().toLowerCase();
    if (!rawType) return null;
    if (rawType.includes('electrical') || rawType.includes('electric') || rawType.includes('电')) return 'Electrical';
    if (rawType.includes('standard') || rawType.includes('std') || rawType.includes('标准')) return 'Standard';
    if (rawType.includes('mechanical') || rawType.includes('mech') || rawType.includes('机')) return 'Mechanical';
    return null;
  };

  const readinessRows = useMemo(() => {
    if (!selectedProjectId) return [] as ProductionForecastItem[];
    const today = new Date();
    const nodeMap = new Map<string, AssemblyBOMItem[]>();
    filteredBoms.forEach(item => {
      if (!nodeMap.has(item.equipmentId)) nodeMap.set(item.equipmentId, []);
      nodeMap.get(item.equipmentId)!.push(item);
    });
    const allNodeCodes = Array.from(nodeMap.keys()).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    const getSubtreeCodes = (nodeCode: string) => allNodeCodes.filter(code => code === nodeCode || code.startsWith(`${nodeCode}.`));

    return allNodeCodes.map((nodeCode, idx) => {
      const subtreeItems = getSubtreeCodes(nodeCode).flatMap(code => nodeMap.get(code) || []);
      const nodeName = nodeMap.get(nodeCode)?.[0]?.equipmentName || nodeCode;
      const partCodes = Array.from(new Set(subtreeItems.map(i => i.partCode).filter(Boolean)));
      const criticalPartCodes = Array.from(new Set(subtreeItems.filter(i => i.isCritical).map(i => i.partCode)));
      const focusPartCodes = criticalPartCodes.length ? criticalPartCodes : partCodes;
      const partEtas = focusPartCodes.map(code => {
        const plans = procurementDeliveryRows.filter(e => e.partCode === code);
        if (!plans.length) return { code, eta: null as Date | null, delayed: true };
        const earliest = plans.map(p => safeDate(p.etaDate)).filter(Boolean).sort((a: any, b: any) => a.getTime() - b.getTime())[0] || null;
        const delayed = plans.some(p => {
          if (p.receivedAt || p.warehouseStatus.includes('到货') || p.warehouseStatus.includes('进库') || p.status === 'Delivered') return false;
          const d = safeDate(p.etaDate);
          return !!d && d.getTime() < today.getTime();
        });
        return { code, eta: earliest, delayed };
      });
      const blocked = Array.from(new Set([
        ...partEtas.filter(x => !x.eta).map(x => x.code),
        ...partEtas.filter(x => x.delayed).map(x => x.code)
      ]));
      const readyCount = partEtas.filter(x => x.eta && x.eta.getTime() <= today.getTime()).length;
      const readiness = focusPartCodes.length ? Math.round((readyCount / focusPartCodes.length) * 100) : 100;
      const latestEta = partEtas.filter(x => x.eta).map(x => x.eta as Date).sort((a, b) => b.getTime() - a.getTime())[0];
      const earliestStartDate = blocked.length > 0 || !latestEta ? '' : format(latestEta, 'yyyy-MM-dd');
      const existing = existingStatusMap.get(nodeCode);
      const autoStatus: ProductionForecastItem['assemblyStatus'] = blocked.length ? 'NotReady' : 'ReadyToAssemble';
      const status = (existing?.assemblyStatus && !['NotReady', 'ReadyToAssemble'].includes(existing.assemblyStatus)) ? existing.assemblyStatus : autoStatus;

      return {
        id: existing?.id || `pf-${Date.now()}-${idx}`,
        projectId: selectedProjectId,
        equipmentId: nodeCode,
        equipmentName: nodeName,
        parentNodeCode: getParentCode(nodeCode) || undefined,
        level: nodeCode.split('.').length,
        bomType: selectedBomType,
        earliestStart: earliestStartDate,
        plannedStart: earliestStartDate,
        plannedEnd: earliestStartDate,
        requiredHours: 0,
        capacityPerWeek: 0,
        readiness,
        risk: blocked.length ? 'High' : readiness >= 100 ? 'Low' : 'Medium',
        blockedParts: blocked,
        assemblyStatus: status,
        startedAt: existing?.startedAt,
        finishedAt: existing?.finishedAt,
        blockerReason: existing?.blockerReason,
        generatedAt: new Date().toISOString()
      } as ProductionForecastItem;
    }).sort((a, b) => a.equipmentId.localeCompare(b.equipmentId, undefined, { numeric: true }));
  }, [selectedProjectId, selectedBomType, filteredBoms, procurementDeliveryRows, existingStatusMap]);

  const readyRows = useMemo(() => readinessRows.filter(r => r.assemblyStatus === 'ReadyToAssemble' || r.assemblyStatus === 'Assembling'), [readinessRows]);
  const blockedRows = useMemo(() => readinessRows.filter(r => r.assemblyStatus === 'Blocked' || r.assemblyStatus === 'NotReady'), [readinessRows]);
  const blockerStats = useMemo(() => {
    const map = new Map<string, number>();
    blockedRows.forEach(r => {
      const category = (r.blockerReason || '').split(':')[0].trim() || '未分类';
      map.set(category, (map.get(category) || 0) + 1);
    });
    return Array.from(map.entries())
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count);
  }, [blockedRows]);
  const currentUserId = state.currentUser?.id || 'system';
  const projectEmployees = useMemo(
    () => getVisibleEmployees(state.currentUser, state.employees),
    [state.currentUser, state.employees]
  );

  const recentWorkflowEvents = useMemo(
    () => (state.workflowEvents || []).filter(e => e.projectId === selectedProjectId).slice(0, 8),
    [state.workflowEvents, selectedProjectId]
  );
  const filteredWorkflowEvents = useMemo(() => {
    const minTime = addDays(new Date(), -Math.max(1, eventDays)).getTime();
    return (state.workflowEvents || [])
      .filter(e => e.projectId === selectedProjectId)
      .filter(e => parseISO(e.createdAt).getTime() >= minTime)
      .filter(e => eventActionFilter === 'all' ? true : e.action === eventActionFilter)
      .filter(e => eventNodeFilter.trim() ? e.entityId.toLowerCase().includes(eventNodeFilter.trim().toLowerCase()) : true)
      .filter(e => eventStatusFilter === 'all' ? true : (e.fromStatus === eventStatusFilter || e.toStatus === eventStatusFilter))
      .slice(0, 200);
  }, [state.workflowEvents, selectedProjectId, eventDays, eventActionFilter, eventNodeFilter, eventStatusFilter]);
  const blockReasonOptions = ['缺料', '质量异常', '工装夹具', '人力不足', '工艺问题', '其他'];
  const statusLabel = (s?: ProductionForecastItem['assemblyStatus']) => {
    if (s === 'NotReady') return '未就绪';
    if (s === 'ReadyToAssemble') return '可开工';
    if (s === 'Assembling') return '装配中';
    if (s === 'Done') return '已完工';
    if (s === 'Blocked') return '阻塞';
    return '-';
  };
  const actionLabel = (a: string) => {
    if (a === 'refresh_readiness') return '刷新可开工清单';
    if (a === 'assembly_status_change') return '装配状态变更';
    return a;
  };
  const exportWorkflowEventsCsv = () => {
    const headers = ['时间', '项目', '实体', '动作', '前状态', '后状态', '说明', '操作人'];
    const rows = filteredWorkflowEvents.map(ev => {
      const operator = projectEmployees.find(x => x.id === ev.operatorId);
      return [
        format(parseISO(ev.createdAt), 'yyyy-MM-dd HH:mm:ss'),
        ev.projectId,
        ev.entityId,
        actionLabel(ev.action),
        statusLabel(ev.fromStatus as any),
        statusLabel(ev.toStatus as any),
        (ev.message || '').replace(/[\r\n,]/g, ' '),
        getEmpName(operator, language)
      ];
    });
    const csv = [headers, ...rows].map(r => r.map(v => `"${String(v ?? '')}"`).join(',')).join('\n');
    const blob = new Blob(["\uFEFF" + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `工作流事件_${selectedProjectId}_${format(new Date(), 'yyyyMMdd_HHmm')}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };


  const appendWorkflowEvent = (prev: AppState, payload: Omit<WorkflowEvent, 'id' | 'createdAt'>): AppState => {
    const event: WorkflowEvent = {
      id: `wf-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      createdAt: new Date().toISOString(),
      ...payload
    };
    return { ...prev, workflowEvents: [event, ...(prev.workflowEvents || [])] };
  };

  const saveForecast = () => {
    setState(prev => {
      const remaining = prev.productionForecasts.filter(x => !(x.projectId === selectedProjectId && x.bomType === selectedBomType));
      let newState = { ...prev, productionForecasts: [...readinessRows, ...remaining] };
      newState = appendWorkflowEvent(newState, {
        module: 'forecast',
        projectId: selectedProjectId,
        entityType: 'project',
        entityId: selectedProjectId,
        action: 'refresh_readiness',
        message: language === 'zh'
          ? `刷新可开工清单，节点数: ${readinessRows.length}`
          : `Refresh readiness list, nodes: ${readinessRows.length}`,
        operatorId: currentUserId
      });
      handleSaveToDatabase(newState);
      return newState;
    });
    alert(language === 'zh' ? '可开工预测已保存' : 'Readiness forecast saved');
  };

  const updateAssemblyStatus = (row: ProductionForecastItem, next: ProductionForecastItem['assemblyStatus']) => {
    if (row.assemblyStatus === 'Done' && next !== 'Done') {
      alert('已完工节点不可回退状态。');
      return;
    }
    if (next === 'Done' && row.assemblyStatus !== 'Assembling') {
      alert(language === 'zh' ? '请先“开始”后再“完工”。' : 'Please Start before Done.');
      return;
    }
    if (next === 'Assembling' && row.assemblyStatus !== 'ReadyToAssemble') {
      alert(language === 'zh' ? '仅“可开工”状态可开始装配。' : 'Only ReadyToAssemble can be started.');
      return;
    }
    if (next === 'Blocked' && !(blockCategory[row.equipmentId] || blockReason[row.equipmentId] || row.blockerReason)) {
      alert(language === 'zh' ? '请先选择阻塞原因或填写备注。' : 'Please select blocker reason first.');
      return;
    }

    setState(prev => {
      const others = prev.productionForecasts.filter(x => !(x.projectId === selectedProjectId && x.bomType === selectedBomType && x.equipmentId === row.equipmentId));
      const updated: ProductionForecastItem = {
        ...row,
        assemblyStatus: next,
        startedAt: next === 'Assembling' ? new Date().toISOString() : row.startedAt,
        finishedAt: next === 'Done' ? new Date().toISOString() : row.finishedAt,
        blockerReason: next === 'Blocked'
          ? `${blockCategory[row.equipmentId] || ''}${blockReason[row.equipmentId] ? `:${blockReason[row.equipmentId]}` : ''}`.replace(/^:/, '') || row.blockerReason || (language === 'zh' ? '待补充原因' : 'Reason pending')
          : row.blockerReason,
        generatedAt: new Date().toISOString()
      };
      let newState = { ...prev, productionForecasts: [updated, ...others] };
      newState = appendWorkflowEvent(newState, {
        module: 'forecast',
        projectId: selectedProjectId,
        entityType: 'assembly_node',
        entityId: row.equipmentId,
        action: 'assembly_status_change',
        fromStatus: row.assemblyStatus,
        toStatus: next,
        message: next === 'Blocked' ? updated.blockerReason : undefined,
        operatorId: currentUserId
      });
      handleSaveToDatabase(newState);
      return newState;
    });
  };

  const importBomTreeFromExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !selectedProjectId) return;
    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'binary' });
        const mapped = wb.SheetNames.flatMap((sheetName, sheetIdx) => {
          const ws = wb.Sheets[sheetName];
          const rows: any[] = XLSX.utils.sheet_to_json(ws);
          return rows.map((r, idx) => ({
            id: `ab-import-${Date.now()}-${sheetIdx}-${idx}`,
            projectId: selectedProjectId,
            equipmentId: String(r['节点编号'] || r['Node Code'] || r['nodeCode'] || '').trim(),
            equipmentName: String(r['节点名称'] || r['Node Name'] || r['nodeName'] || '').trim(),
            bomType: inferImportedBomType(r, sheetName) || selectedBomType,
            partCode: String(r['物料编码'] || r['Part Code'] || r['partCode'] || '').trim(),
            partName: String(r['物料名称'] || r['Part Name'] || r['partName'] || '').trim(),
            quantity: Math.max(1, Number(r['数量'] || r['Qty'] || r['quantity'] || 1)),
            assemblyHours: 0,
            isCritical: String(r['关键件'] || r['Critical'] || 'Y').toUpperCase() !== 'N',
            priority: Number(r['优先级'] || r['Priority'] || 50) || 50
          } as AssemblyBOMItem));
        }).filter(x => x.equipmentId && x.partCode);
        if (!mapped.length) {
          alert(language === 'zh' ? '未识别到有效BOM行，请检查表头。' : 'No valid BOM rows found.');
          return;
        }
        setState(prev => {
          const newState = { ...prev, assemblyBoms: [...mapped, ...prev.assemblyBoms] };
          handleSaveToDatabase(newState);
          return newState;
        });
      } catch {
        alert(language === 'zh' ? 'BOM导入失败，请检查Excel格式。' : 'BOM import failed.');
      }
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div className="space-y-6 w-full max-w-full pb-10">
      <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-6 space-y-4">
        <div className="flex flex-wrap items-center gap-2">
          <select value={selectedProjectId} onChange={e => setSelectedProjectId(e.target.value)} className="input-field py-2 px-3 h-10 text-xs font-bold min-w-[220px]">
            {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
          {(['Mechanical', 'Electrical'] as ('Mechanical' | 'Electrical')[]).map(type => (
            <button key={type} onClick={() => setSelectedBomType(type)} className={cn("px-3 py-1.5 rounded-xl text-xs font-black border", selectedBomType === type ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200")}>
              {type === 'Mechanical' ? (language === 'zh' ? '机械装配BOM' : 'Mechanical Assembly BOM') : (language === 'zh' ? '电气装配BOM' : 'Electrical Assembly BOM')}
            </button>
          ))}
          <button onClick={saveForecast} className="bg-blue-600 text-white px-4 py-2 rounded-xl text-xs font-black flex items-center gap-2">
            <Save size={12} /> {language === 'zh' ? '刷新并保存可开工列表' : 'Refresh & Save Readiness'}
          </button>
          <label className="btn-secondary py-2 px-3 text-xs cursor-pointer flex items-center gap-2">
            <Upload size={12} />
            {language === 'zh' ? '导入树状BOM并自动分类' : 'Import BOM Tree & Auto Classify'}
            <input type="file" accept=".xlsx,.xls" className="hidden" onChange={importBomTreeFromExcel} />
          </label>
        </div>

        <div className="rounded-2xl border border-slate-100 bg-slate-50/60 p-4">
          <div className="text-[10px] font-black text-slate-400 uppercase mb-2">按设备型号BOM展开需求（强关联）</div>
          <div className="grid grid-cols-1 md:grid-cols-5 gap-2 items-end">
            <div className="space-y-1">
              <div className="text-[10px] font-black text-slate-400">设备型号（ID）</div>
              <select value={selectedDeviceModelId} onChange={e => setSelectedDeviceModelId(e.target.value)} className="input-field h-9 py-1 px-2 text-xs">
                {(() => {
                  const all = (state.deviceModels || []) as DeviceModel[];
                  return all.map(dm => (
                    <option key={dm.id} value={dm.id}>{formatDeviceModelLabel(dm, all)}</option>
                  ));
                })()}
              </select>
            </div>
            <div className="space-y-1 md:col-span-2">
              <div className="text-[10px] font-black text-slate-400">装配BOM版本</div>
              <select value={selectedDeviceAssemblyTplId} onChange={e => setSelectedDeviceAssemblyTplId(e.target.value)} className="input-field h-9 py-1 px-2 text-xs">
                {deviceAssemblyTplOptions.length === 0 && <option value="">（该型号暂无装配BOM）</option>}
                {deviceAssemblyTplOptions.map(op => (
                  <option key={op.id} value={op.id}>{op.label}</option>
                ))}
              </select>
            </div>
            <div className="space-y-1">
              <div className="text-[10px] font-black text-slate-400">计划台数</div>
              <input
                type="number"
                min={1}
                value={plannedDeviceCount}
                onChange={e => setPlannedDeviceCount(Math.max(1, Math.floor(Number(e.target.value) || 1)))}
                className="input-field h-9 py-1 px-2 text-xs"
              />
            </div>
            <button
              onClick={generateProjectNeedsFromDeviceModel}
              disabled={!selectedDeviceAssemblyTplId}
              className={cn("h-9 px-3 rounded-xl text-xs font-black border",
                selectedDeviceAssemblyTplId ? "bg-slate-900 text-white border-slate-900" : "bg-white text-slate-400 border-slate-200")}
            >
              生成项目需求/节点
            </button>
          </div>
          <div className="text-[11px] text-slate-500 mt-2">
            会覆盖同项目下该设备型号（{selectedDeviceModelId}）之前生成的装配节点/采购需求；用于“排班预测”与“齐套判断”。
          </div>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
          <div className="rounded-2xl border border-slate-100 p-4 bg-slate-50/60"><p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? 'BOM条目' : 'BOM Items'}</p><p className="text-2xl font-black text-slate-900">{filteredBoms.length}</p></div>
          <div className="rounded-2xl border border-slate-100 p-4 bg-slate-50/60"><p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? '采购交付数据' : 'Procurement Delivery Rows'}</p><p className="text-2xl font-black text-slate-900">{procurementDeliveryRows.length}</p></div>
          <div className="rounded-2xl border border-slate-100 p-4 bg-slate-50/60"><p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? '可开工单元' : 'Ready Units'}</p><p className="text-2xl font-black text-slate-900">{readyRows.length}</p></div>
          <div className="rounded-2xl border border-slate-100 p-4 bg-slate-50/60"><p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? '节点总数' : 'Total Nodes'}</p><p className="text-2xl font-black text-slate-900">{readinessRows.length}</p></div>
        </div>
      </div>

      <div className="grid grid-cols-1 gap-4">
        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-4 space-y-2">
          <h4 className="text-sm font-black">{language === 'zh' ? '到货来源说明' : 'Delivery Source'}</h4>
          <p className="text-xs text-slate-600 leading-relaxed">
            {language === 'zh'
              ? '到货计划直接读取采购管理中的物料交付日期（交付日期/预计到货/收货状态），此页面不再手工维护ETA。'
              : 'Delivery dates are read directly from Procurement material delivery fields (delivery/ETA/receipt status). No manual ETA maintenance here.'}
          </p>
          <p className="text-xs text-slate-600 leading-relaxed">
            {language === 'zh'
              ? 'BOM树节点物料不再手工新增，统一通过导入 BOM 表生成；系统优先按 Excel 内的分类字段识别机械/电气/标准件，其次按工作表名称识别。'
              : 'BOM tree materials are no longer added manually. They come from imported BOM tables; the system classifies rows by BOM category columns first, then by sheet name.'}
          </p>
          <p className="text-[11px] font-bold text-slate-400">
            {language === 'zh'
              ? `当前口径：${selectedBomType === 'Mechanical' ? '机械装配 = 机械BOM + 标准件BOM' : '电气装配 = 电气BOM'}`
              : `Scope: ${selectedBomType === 'Mechanical' ? 'Mechanical assembly = Mechanical BOM + Standard parts BOM' : 'Electrical assembly = Electrical BOM'}`}
          </p>
        </div>
      </div>

      <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-4 overflow-auto">
        <h4 className="text-sm font-black mb-3">{language === 'zh' ? '装配单元可开工判断（树状）' : 'Assembly Unit Readiness (Tree)'}</h4>
        <table className="w-full text-xs border-collapse">
          <thead><tr className="bg-slate-50"><th className="p-2 text-left">{language === 'zh' ? '节点' : 'Node'}</th><th className="p-2 text-left">{language === 'zh' ? '父节点' : 'Parent'}</th><th className="p-2 text-left">{language === 'zh' ? '可开工日期' : 'Ready Date'}</th><th className="p-2 text-left">{language === 'zh' ? '可开工周' : 'Ready Week'}</th><th className="p-2 text-left">{language === 'zh' ? '齐套率' : 'Readiness'}</th><th className="p-2 text-left">{language === 'zh' ? '缺料' : 'Blocked Parts'}</th><th className="p-2 text-left">{language === 'zh' ? '状态' : 'Status'}</th><th className="p-2 text-left">{language === 'zh' ? '闭环操作' : 'Actions'}</th></tr></thead>
          <tbody>
            {readinessRows.map(row => (
              <tr key={row.equipmentId} className="border-t border-slate-100">
                <td className="p-2 font-bold" style={{ paddingLeft: `${Math.max(8, (row.level || 1) * 12)}px` }}>{row.equipmentId} {row.equipmentName}</td>
                <td className="p-2 text-slate-500">{row.parentNodeCode || '-'}</td>
                <td className="p-2">{row.earliestStart || '-'}</td>
                <td className="p-2">{row.earliestStart ? format(startOfWeek(parseISO(row.earliestStart), { weekStartsOn: 1 }), 'yyyy-MM-dd') : '-'}</td>
                <td className="p-2">{row.readiness}%</td>
                <td className="p-2 text-rose-600">{row.blockedParts.slice(0, 3).join(', ') || '-'}</td>
                <td className={cn("p-2 font-black", row.assemblyStatus === 'Done' ? "text-emerald-600" : row.assemblyStatus === 'Assembling' ? "text-blue-600" : row.assemblyStatus === 'Blocked' ? "text-rose-600" : row.assemblyStatus === 'ReadyToAssemble' ? "text-amber-600" : "text-slate-500")}>{statusLabel(row.assemblyStatus)}</td>
                <td className="p-2">
                  <div className="flex flex-wrap gap-1">
                    <button onClick={() => updateAssemblyStatus(row, 'Assembling')} className="px-2 py-0.5 rounded-lg text-[10px] font-black border border-blue-200 text-blue-600">{language === 'zh' ? '开始' : 'Start'}</button>
                    <button onClick={() => updateAssemblyStatus(row, 'Done')} className="px-2 py-0.5 rounded-lg text-[10px] font-black border border-emerald-200 text-emerald-600">{language === 'zh' ? '完工' : 'Done'}</button>
                    <button
                      onClick={() => updateAssemblyStatus(row, 'Blocked')}
                      disabled={!(blockCategory[row.equipmentId] || blockReason[row.equipmentId] || row.blockerReason)}
                      className={cn(
                        "px-2 py-0.5 rounded-lg text-[10px] font-black border",
                        (blockCategory[row.equipmentId] || blockReason[row.equipmentId] || row.blockerReason)
                          ? "border-rose-200 text-rose-600"
                          : "border-slate-200 text-slate-300 cursor-not-allowed"
                      )}
                    >
                      {language === 'zh' ? '阻塞' : 'Block'}
                    </button>
                  </div>
                  {row.assemblyStatus !== 'Done' && (
                    <div className="grid grid-cols-2 gap-2 mt-1">
                      <select
                        value={blockCategory[row.equipmentId] || ''}
                        onChange={e => setBlockCategory(prev => ({ ...prev, [row.equipmentId]: e.target.value }))}
                        className="input-field h-7 py-0.5 px-2 text-[10px]"
                      >
                        <option value="">阻塞分类</option>
                        {blockReasonOptions.map(op => <option key={op} value={op}>{op}</option>)}
                      </select>
                      <input
                        value={blockReason[row.equipmentId] || ''}
                        onChange={e => setBlockReason(prev => ({ ...prev, [row.equipmentId]: e.target.value }))}
                        placeholder="备注(可选)"
                        className="input-field h-7 py-0.5 px-2 text-[10px]"
                      />
                    </div>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-4">
          <h4 className="text-sm font-black mb-3">{language === 'zh' ? '今日可开工单元' : 'Ready To Assemble Today'}</h4>
          <div className="space-y-2">
            {readyRows.length === 0 && <p className="text-xs text-slate-400">{language === 'zh' ? '暂无可开工单元' : 'No ready unit.'}</p>}
            {readyRows.map(row => (
              <div key={`ready-${row.equipmentId}`} className="p-3 rounded-xl border border-emerald-100 bg-emerald-50/60">
                <p className="text-xs font-black text-emerald-700">{row.equipmentId} {row.equipmentName}</p>
                <p className="text-[11px] text-emerald-600">{language === 'zh' ? '可开工日期' : 'Ready Date'}: {row.earliestStart || '-'}</p>
              </div>
            ))}
          </div>
        </div>
        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-4">
          <h4 className="text-sm font-black mb-3">{language === 'zh' ? '阻塞清单' : 'Blocked List'}</h4>
          <div className="space-y-2">
            {blockedRows.length === 0 && <p className="text-xs text-slate-400">{language === 'zh' ? '当前无阻塞' : 'No blocked nodes.'}</p>}
            {blockedRows.map(row => (
              <div key={`blocked-${row.equipmentId}`} className="p-3 rounded-xl border border-rose-100 bg-rose-50/60">
                <p className="text-xs font-black text-rose-700">{row.equipmentId} {row.equipmentName}</p>
                <p className="text-[11px] text-rose-600">{language === 'zh' ? '缺料/原因' : 'Blocked By'}: {(row.blockedParts.join(', ') || row.blockerReason || '-')}</p>
              </div>
            ))}
          </div>
          <div className="mt-3 pt-3 border-t border-slate-100">
            <p className="text-[10px] font-black text-slate-400 uppercase mb-2">阻塞分类统计</p>
            <div className="flex flex-wrap gap-2">
              {blockerStats.length === 0 && <span className="text-xs text-slate-400">暂无</span>}
              {blockerStats.map(item => (
                <span key={item.name} className="px-2 py-1 rounded-lg text-[10px] font-black border border-rose-100 bg-rose-50 text-rose-700">
                  {item.name} {item.count}
                </span>
              ))}
            </div>
          </div>
        </div>
        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-4">
          <div className="flex items-center justify-between mb-3">
            <h4 className="text-sm font-black">{language === 'zh' ? '近期工作流事件' : 'Recent Workflow Events'}</h4>
            <button onClick={exportWorkflowEventsCsv} className="btn-secondary py-1 px-2 text-[10px] font-black">
              导出CSV
            </button>
          </div>
          <div className="grid grid-cols-2 gap-2 mb-2">
            <select value={eventActionFilter} onChange={e => setEventActionFilter(e.target.value as any)} className="input-field h-8 py-1 px-2 text-[10px]">
              <option value="all">全部动作</option>
              <option value="refresh_readiness">刷新可开工清单</option>
              <option value="assembly_status_change">装配状态变更</option>
            </select>
            <select value={eventStatusFilter} onChange={e => setEventStatusFilter(e.target.value as any)} className="input-field h-8 py-1 px-2 text-[10px]">
              <option value="all">全部状态</option>
              <option value="NotReady">未就绪</option>
              <option value="ReadyToAssemble">可开工</option>
              <option value="Assembling">装配中</option>
              <option value="Done">已完工</option>
              <option value="Blocked">阻塞</option>
            </select>
            <input value={eventNodeFilter} onChange={e => setEventNodeFilter(e.target.value)} placeholder="节点筛选（如1.1）" className="input-field h-8 py-1 px-2 text-[10px]" />
            <input type="number" min={1} max={365} value={eventDays} onChange={e => setEventDays(Math.max(1, Math.min(365, Number(e.target.value) || 30)))} className="input-field h-8 py-1 px-2 text-[10px]" />
          </div>
          <div className="space-y-2 max-h-[260px] overflow-y-auto">
            {filteredWorkflowEvents.length === 0 && <p className="text-xs text-slate-400">{language === 'zh' ? '暂无事件' : 'No events.'}</p>}
            {filteredWorkflowEvents.map(ev => {
              const operator = projectEmployees.find(x => x.id === ev.operatorId);
              return (
                <div key={ev.id} className="p-2 rounded-xl border border-slate-100 bg-slate-50/70">
                  <p className="text-[11px] font-black text-slate-700">{actionLabel(ev.action)}</p>
                  <p className="text-[10px] text-slate-500">{ev.entityId} {ev.fromStatus ? `: ${ev.fromStatus} -> ${ev.toStatus}` : ''}</p>
                  <p className="text-[10px] text-slate-400">{format(parseISO(ev.createdAt), 'yyyy-MM-dd HH:mm')} · {getEmpName(operator, language)}</p>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}

export function ProcurementShortageDashboard({
  state,
  language,
  projectIdFilter,
  onProjectFilterChange
}: {
  state: AppState;
  language: 'zh' | 'en';
  projectIdFilter: string;
  onProjectFilterChange: (id: string) => void;
}) {
  const [bomTypeFilter, setBomTypeFilter] = useState<MaterialType | 'All'>('All');

  const today = useMemo(() => {
    const d = new Date();
    // normalize to date-only for comparisons
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }, []);

  const toDate = (s?: string) => {
    if (!s) return null;
    try {
      const d = parseISO(s);
      if (isNaN(d.getTime())) return null;
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    } catch {
      return null;
    }
  };

  const dayDiff = (a: Date, b: Date) => Math.floor((a.getTime() - b.getTime()) / 86400000);

  const typeLabel = (t: MaterialType) => {
    if (language === 'en') return t;
    if (t === 'Mechanical') return '机械';
    if (t === 'Electrical') return '电气';
    if (t === 'Standard') return '标准件';
    if (t === 'Spare') return '备件';
    return t;
  };

  const projectOptions = useMemo(
    () => state.projects.filter(p => p.id !== 'p-general' && p.id !== 'p-leave'),
    [state.projects]
  );
  const projectNameById = useMemo(() => {
    const m: Record<string, string> = {};
    projectOptions.forEach(p => (m[p.id] = p.name));
    return m;
  }, [projectOptions]);

  type RiskStatus = 'satisfied' | 'overdue' | 'at_risk' | 'on_track' | 'missing_info';
  type RiskRow = MaterialRequirement & {
    materialKey: string;
    requiredQty: number;
    onHand: number;
    reservedQty: number;
    shortageQty: number;
    availableAfterAll: number;
    riskStatus: RiskStatus;
    overdueDays: number;
  };

  const rows = useMemo(() => {
    const list = state.materialRequirements.filter(r => {
      if (projectIdFilter && r.projectId !== projectIdFilter) return false;
      if (bomTypeFilter !== 'All' && r.type !== bomTypeFilter) return false;
      return true;
    });

    // 物料唯一识别：
    // - 机械加工件：图号 + 版本号
    // - 其他：型号/零件代号(model/code) + 版本号
    const keyOf = (r: MaterialRequirement) => {
      const drawing = ((r as any).drawingNo || '').toString().trim();
      const revision = ((r as any).revision || '').toString().trim();
      const base = drawing || (r.model || r.code || r.name || '').toString().trim();
      return `${base}`.toLowerCase() + `|${revision}`.toLowerCase();
    };

    // 1) material-level onHand (use inventoryQuantity first; fallback inbound-outbound)
    const onHandByKey: Record<string, number> = {};
    list.forEach(r => {
      const k = keyOf(r);
      if (!k) return;
      const inv =
        typeof (r as any).inventoryQuantity === 'number'
          ? (r as any).inventoryQuantity
          : (typeof (r as any).inboundQuantity === 'number' && typeof (r as any).outboundQuantity === 'number')
            ? (r as any).inboundQuantity - (r as any).outboundQuantity
            : 0;
      onHandByKey[k] = Math.max(onHandByKey[k] ?? 0, Number(inv) || 0);
    });

    // 2) group rows by materialKey, allocate reservation by needByDate asc
    const byKey: Record<string, MaterialRequirement[]> = {};
    list.forEach(r => {
      const k = keyOf(r);
      if (!k) return;
      if (!byKey[k]) byKey[k] = [];
      byKey[k].push(r);
    });

    const out: RiskRow[] = [];
    Object.entries(byKey).forEach(([k, items]) => {
      const sorted = [...items].sort((a, b) => {
        const ad = toDate((a as any).needByDate);
        const bd = toDate((b as any).needByDate);
        if (!ad && !bd) return (a.id || '').localeCompare(b.id || '');
        if (!ad) return 1;
        if (!bd) return -1;
        return ad.getTime() - bd.getTime();
      });

      let remaining = onHandByKey[k] ?? 0;
      sorted.forEach(r => {
        const requiredQty = Number((r as any).quantity) || 0;
        const reservedQty = Math.max(0, Math.min(requiredQty, remaining));
        remaining -= reservedQty;
        const shortageQty = Math.max(0, requiredQty - reservedQty);

        const need = toDate((r as any).needByDate);
        const eta = toDate((r as any).expectedArrival);
        const committed = toDate((r as any).deliveryTime);

        let riskStatus: RiskStatus = 'on_track';
        let overdueDays = 0;
        if (!need) {
          riskStatus = 'missing_info';
        } else if (shortageQty <= 0) {
          riskStatus = 'satisfied';
        } else if (today.getTime() > need.getTime()) {
          riskStatus = 'overdue';
          overdueDays = dayDiff(today, need);
        } else {
          const willLate =
            (eta && eta.getTime() > need.getTime()) ||
            (!eta && committed && committed.getTime() > need.getTime());
          riskStatus = willLate ? 'at_risk' : 'on_track';
        }

        out.push({
          ...(r as any),
          materialKey: k,
          requiredQty,
          onHand: onHandByKey[k] ?? 0,
          reservedQty,
          shortageQty,
          availableAfterAll: remaining,
          riskStatus,
          overdueDays
        });
      });
    });

    // sort: overdue desc, at_risk, on_track, missing_info, satisfied
    const rank = (s: RiskStatus) =>
      s === 'overdue' ? 0 : s === 'at_risk' ? 1 : s === 'on_track' ? 2 : s === 'missing_info' ? 3 : 4;
    out.sort((a, b) => {
      const ra = rank(a.riskStatus);
      const rb = rank(b.riskStatus);
      if (ra !== rb) return ra - rb;
      if (a.riskStatus === 'overdue' && b.riskStatus === 'overdue') return (b.overdueDays || 0) - (a.overdueDays || 0);
      const ad = toDate((a as any).needByDate);
      const bd = toDate((b as any).needByDate);
      if (ad && bd) return ad.getTime() - bd.getTime();
      return (a.id || '').localeCompare(b.id || '');
    });

    return out;
  }, [state.materialRequirements, projectIdFilter, bomTypeFilter, today]);

  const kpis = useMemo(() => {
    const base = {
      overdue: 0,
      atRisk: 0,
      missingInfo: 0,
      shortage: 0,
      satisfied: 0
    };
    rows.forEach(r => {
      if (r.riskStatus === 'overdue') base.overdue += 1;
      else if (r.riskStatus === 'at_risk') base.atRisk += 1;
      else if (r.riskStatus === 'missing_info') base.missingInfo += 1;
      else if (r.riskStatus === 'satisfied') base.satisfied += 1;
      if (r.shortageQty > 0) base.shortage += 1;
    });
    return base;
  }, [rows]);

  const overdueRows = useMemo(() => rows.filter(r => r.riskStatus === 'overdue'), [rows]);
  const atRiskRows = useMemo(() => rows.filter(r => r.riskStatus === 'at_risk' || r.riskStatus === 'on_track' || r.riskStatus === 'missing_info').filter(r => r.shortageQty > 0), [rows]);

  const statusPill = (r: RiskRow) => {
    const txt =
      r.riskStatus === 'overdue'
        ? (language === 'zh' ? `延期${r.overdueDays}天` : `Overdue ${r.overdueDays}d`)
        : r.riskStatus === 'at_risk'
          ? (language === 'zh' ? '有风险' : 'At Risk')
          : r.riskStatus === 'missing_info'
            ? (language === 'zh' ? '缺需求时间' : 'Need date missing')
            : r.riskStatus === 'satisfied'
              ? (language === 'zh' ? '已满足' : 'Satisfied')
              : (language === 'zh' ? '推进中' : 'On Track');

    const cls =
      r.riskStatus === 'overdue'
        ? 'bg-rose-50 text-rose-700'
        : r.riskStatus === 'at_risk'
          ? 'bg-amber-50 text-amber-700'
          : r.riskStatus === 'missing_info'
            ? 'bg-slate-100 text-slate-600'
            : r.riskStatus === 'satisfied'
              ? 'bg-emerald-50 text-emerald-700'
              : 'bg-blue-50 text-blue-700';

    return <span className={cn("text-[10px] font-black px-2 py-1 rounded-lg", cls)}>{txt}</span>;
  };

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center gap-2">
        <div className="flex items-center gap-2">
          <span className="text-[10px] font-black text-slate-400">{language === 'zh' ? '项目筛选' : 'Project'}</span>
          <select
            value={projectIdFilter}
            onChange={e => onProjectFilterChange(e.target.value)}
            className="input-field h-9 py-1 px-2 text-[11px] font-black min-w-[180px]"
          >
            <option value="">{language === 'zh' ? '全部项目' : 'All Projects'}</option>
            {projectOptions.map(p => (
              <option key={p.id} value={p.id}>{p.name}</option>
            ))}
          </select>
        </div>
        <div className="flex items-center gap-2">
          <span className="text-[10px] font-black text-slate-400">BOM</span>
          <select
            value={bomTypeFilter}
            onChange={e => setBomTypeFilter(e.target.value as any)}
            className="input-field h-9 py-1 px-2 text-[11px] font-black min-w-[160px]"
          >
            <option value="All">{language === 'zh' ? '全部类型' : 'All Types'}</option>
            {(['Mechanical', 'Electrical', 'Standard', 'Spare'] as MaterialType[]).map(t => (
              <option key={t} value={t}>{typeLabel(t)}</option>
            ))}
          </select>
        </div>
      </div>

      <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
        <div className="p-4 rounded-2xl border border-slate-100 bg-white">
          <div className="text-[10px] font-black text-slate-400">{language === 'zh' ? '缺口项' : 'Shortage Items'}</div>
          <div className="text-2xl font-black text-slate-900">{kpis.shortage}</div>
        </div>
        <div className="p-4 rounded-2xl border border-rose-100 bg-rose-50/40">
          <div className="text-[10px] font-black text-rose-500">{language === 'zh' ? '已延期' : 'Overdue'}</div>
          <div className="text-2xl font-black text-rose-700">{kpis.overdue}</div>
        </div>
        <div className="p-4 rounded-2xl border border-amber-100 bg-amber-50/40">
          <div className="text-[10px] font-black text-amber-600">{language === 'zh' ? '有风险' : 'At Risk'}</div>
          <div className="text-2xl font-black text-amber-700">{kpis.atRisk}</div>
        </div>
        <div className="p-4 rounded-2xl border border-slate-100 bg-slate-50/40">
          <div className="text-[10px] font-black text-slate-500">{language === 'zh' ? '缺需求时间' : 'Need date missing'}</div>
          <div className="text-2xl font-black text-slate-700">{kpis.missingInfo}</div>
        </div>
        <div className="p-4 rounded-2xl border border-emerald-100 bg-emerald-50/40">
          <div className="text-[10px] font-black text-emerald-600">{language === 'zh' ? '已满足' : 'Satisfied'}</div>
          <div className="text-2xl font-black text-emerald-700">{kpis.satisfied}</div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden">
          <div className="p-4 border-b border-slate-50 flex items-center justify-between">
            <div className="text-xs font-black text-slate-700">{language === 'zh' ? '延期清单（按需求时间）' : 'Overdue List (Need-by Date)'}</div>
            <div className="text-[10px] font-black text-slate-400">{overdueRows.length}</div>
          </div>
          <div className="max-h-[420px] overflow-auto">
            <table className="w-full text-left text-[11px]">
              <thead className="sticky top-0 bg-white border-b border-slate-100">
                <tr className="text-[10px] font-black text-slate-400">
                  <th className="px-3 py-2">{language === 'zh' ? '项目' : 'Project'}</th>
                  <th className="px-3 py-2">{language === 'zh' ? '物料' : 'Item'}</th>
                  <th className="px-3 py-2">{language === 'zh' ? '需求/可用' : 'Req/Avail'}</th>
                  <th className="px-3 py-2">{language === 'zh' ? '需求时间' : 'Need-by'}</th>
                  <th className="px-3 py-2">{language === 'zh' ? '状态' : 'Status'}</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {overdueRows.length === 0 && (
                  <tr><td colSpan={5} className="px-3 py-6 text-slate-400 text-xs">{language === 'zh' ? '暂无延期' : 'No overdue items.'}</td></tr>
                )}
                {overdueRows.map(r => (
                  <tr key={r.id} className="hover:bg-slate-50/60">
                    <td className="px-3 py-2 font-bold text-slate-700">
                      <div className="text-[10px] text-slate-400">{typeLabel(r.type)}</div>
                      {projectNameById[r.projectId] || r.projectId}
                    </td>
                    <td className="px-3 py-2">
                      <div className="font-black text-slate-800">{r.model || r.code || r.name}</div>
                      <div className="text-[10px] text-slate-400 truncate">{r.supplier || r.actualSupplier || ''}</div>
                    </td>
                    <td className="px-3 py-2">
                      <div className="font-black text-slate-800">{r.requiredQty}</div>
                      <div className="text-[10px] text-slate-400">{language === 'zh' ? `可用 ${r.reservedQty}` : `Avail ${r.reservedQty}`}</div>
                      {r.shortageQty > 0 && <div className="text-[10px] font-black text-rose-600">{language === 'zh' ? `缺口 ${r.shortageQty}` : `Gap ${r.shortageQty}`}</div>}
                    </td>
                    <td className="px-3 py-2 font-black text-slate-700">{(r as any).needByDate || '-'}</td>
                    <td className="px-3 py-2">{statusPill(r)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden">
          <div className="p-4 border-b border-slate-50 flex items-center justify-between">
            <div className="text-xs font-black text-slate-700">{language === 'zh' ? '缺口/风险清单（可用库存口径）' : 'Shortage / Risk List (Available Inventory)'}</div>
            <div className="text-[10px] font-black text-slate-400">{atRiskRows.length}</div>
          </div>
          <div className="max-h-[420px] overflow-auto">
            <table className="w-full text-left text-[11px]">
              <thead className="sticky top-0 bg-white border-b border-slate-100">
                <tr className="text-[10px] font-black text-slate-400">
                  <th className="px-3 py-2">{language === 'zh' ? '项目' : 'Project'}</th>
                  <th className="px-3 py-2">{language === 'zh' ? '物料' : 'Item'}</th>
                  <th className="px-3 py-2">{language === 'zh' ? '缺口' : 'Gap'}</th>
                  <th className="px-3 py-2">{language === 'zh' ? '需求/ETA/交期' : 'Need/ETA/Commit'}</th>
                  <th className="px-3 py-2">{language === 'zh' ? '状态' : 'Status'}</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {atRiskRows.length === 0 && (
                  <tr><td colSpan={5} className="px-3 py-6 text-slate-400 text-xs">{language === 'zh' ? '暂无缺口' : 'No shortage items.'}</td></tr>
                )}
                {atRiskRows.map(r => (
                  <tr key={r.id} className="hover:bg-slate-50/60">
                    <td className="px-3 py-2 font-bold text-slate-700">
                      <div className="text-[10px] text-slate-400">{typeLabel(r.type)}</div>
                      {projectNameById[r.projectId] || r.projectId}
                    </td>
                    <td className="px-3 py-2">
                      <div className="font-black text-slate-800">{r.model || r.code || r.name}</div>
                      <div className="text-[10px] text-slate-400 truncate">{r.supplier || r.actualSupplier || ''}</div>
                    </td>
                    <td className="px-3 py-2">
                      <div className={cn("font-black", r.shortageQty > 0 ? "text-rose-700" : "text-slate-700")}>
                        {r.shortageQty.toFixed(1)}
                      </div>
                      <div className="text-[10px] text-slate-400">{language === 'zh' ? `需求 ${r.requiredQty}` : `Req ${r.requiredQty}`}</div>
                    </td>
                    <td className="px-3 py-2 text-[10px] font-black text-slate-600">
                      <div>{language === 'zh' ? '需:' : 'Need:'} {(r as any).needByDate || '-'}</div>
                      <div>{language === 'zh' ? 'ETA:' : 'ETA:'} {(r as any).expectedArrival || '-'}</div>
                      <div>{language === 'zh' ? '交期:' : 'Commit:'} {(r as any).deliveryTime || '-'}</div>
                    </td>
                    <td className="px-3 py-2">{statusPill(r)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

function KitReadinessView() {
  const { state, language } = useContext(AppContext)!;
  const projects = state.projects.filter(p => p.id !== 'p-general' && p.id !== 'p-leave');
  const [selectedProjectId, setSelectedProjectId] = useState<string>(projects[0]?.id || '');
  const [selectedBomType, setSelectedBomType] = useState<'Mechanical' | 'Electrical'>('Mechanical');
  const [expandedNodes, setExpandedNodes] = useState<Set<string>>(new Set());

  useEffect(() => {
    if ((!selectedProjectId || !projects.some(p => p.id === selectedProjectId)) && projects.length > 0) {
      setSelectedProjectId(projects[0].id);
    }
  }, [projects, selectedProjectId]);

  const safeDate = (s?: string) => {
    if (!s) return null;
    try {
      const d = parseISO(s);
      return isNaN(d.getTime()) ? null : d;
    } catch {
      return null;
    }
  };

  const assemblySourceTypes = useMemo<MaterialType[]>(
    () => (selectedBomType === 'Mechanical' ? ['Mechanical', 'Standard'] : ['Electrical']),
    [selectedBomType]
  );

  const filteredBoms = useMemo(
    () => state.assemblyBoms.filter(x => x.projectId === selectedProjectId && x.bomType === selectedBomType),
    [state.assemblyBoms, selectedProjectId, selectedBomType]
  );

  const procurementDeliveryRows = useMemo(() => {
    return (state.materialRequirements || [])
      .filter(r => r.projectId === selectedProjectId && assemblySourceTypes.includes(r.type))
      .map(r => ({
        partCode: `${r.model || r.code || ''}`.trim(),
        etaDate: r.deliveryTime || r.expectedArrival || '',
        receivedAt: r.receivedAt || '',
        warehouseStatus: r.warehouseStatus || '',
        status: r.status,
        name: r.name || '',
        drawingNo: (r as any).drawingNo || '',
        revision: (r as any).revision || '',
      }))
      .filter(r => !!r.partCode);
  }, [state.materialRequirements, selectedProjectId, assemblySourceTypes]);

  const readinessRows = useMemo(() => {
    if (!selectedProjectId) return [] as Array<{
      equipmentId: string; equipmentName: string; level: number;
      readiness: number; blockedParts: string[];
      partDetails: Array<{ partCode: string; name: string; demandQty: number; etaDate: string; receivedAt: string; warehouseStatus: string; delayed: boolean }>;
    }>;
    const today = new Date();
    const nodeMap = new Map<string, AssemblyBOMItem[]>();
    filteredBoms.forEach(item => {
      if (!nodeMap.has(item.equipmentId)) nodeMap.set(item.equipmentId, []);
      nodeMap.get(item.equipmentId)!.push(item);
    });
    const allNodeCodes = Array.from(nodeMap.keys()).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    const getSubtreeCodes = (nodeCode: string) => allNodeCodes.filter(code => code === nodeCode || code.startsWith(`${nodeCode}.`));

    return allNodeCodes.map(nodeCode => {
      const subtreeItems = getSubtreeCodes(nodeCode).flatMap(code => nodeMap.get(code) || []);
      const nodeName = nodeMap.get(nodeCode)?.[0]?.equipmentName || nodeCode;
      const demandByCode = new Map<string, number>();
      subtreeItems.forEach(i => {
        const k = getMaterialKey({ drawingNo: (i as any).drawingNo, revision: (i as any).revision, model: i.partCode });
        demandByCode.set(k, (demandByCode.get(k) || 0) + (i.quantity || 0));
      });
      const deliveryByKey = new Map<string, typeof procurementDeliveryRows[0]>();
      procurementDeliveryRows.forEach(r => {
        const k = getMaterialKey({ drawingNo: (r as any).drawingNo, revision: (r as any).revision, model: r.partCode, name: (r as any).name });
        if (!deliveryByKey.has(k) || (!deliveryByKey.get(k)!.receivedAt && r.receivedAt)) deliveryByKey.set(k, r);
      });

      const partCodes = Array.from(new Set(subtreeItems.map(i => i.partCode).filter(Boolean)));
      const criticalPartCodes = Array.from(new Set(subtreeItems.filter(i => i.isCritical).map(i => i.partCode)));
      const focusPartCodes = criticalPartCodes.length ? criticalPartCodes : partCodes;
      const focusPartKeys = Array.from(new Set(
        subtreeItems.filter(i => focusPartCodes.includes(i.partCode)).map(i =>
          getMaterialKey({ drawingNo: (i as any).drawingNo, revision: (i as any).revision, model: i.partCode })
        )
      ));

      const partEtas = focusPartKeys.map(key => {
        const plan = deliveryByKey.get(key);
        const code = key.split('|')[0];
        const demandQty = demandByCode.get(key) || 0;
        if (!plan) return { code, eta: null as Date | null, delayed: true, etaDate: '', demandQty, name: '', warehouseStatus: '', receivedAt: '' };
        const eta = safeDate(plan.etaDate);
        const delayed = !(plan.receivedAt || (plan.warehouseStatus || '').includes('到货') || (plan.warehouseStatus || '').includes('进库') || plan.status === 'Delivered')
          && eta && eta.getTime() < today.getTime();
        return { code, eta, delayed, etaDate: plan.etaDate || '', demandQty, name: plan.name || '', warehouseStatus: plan.warehouseStatus || '', receivedAt: plan.receivedAt || '' };
      });

      const blocked = partEtas.filter(x => !x.eta || x.delayed);
      const blockedParts = blocked.map(x => `${x.code}${x.etaDate ? ' (ETA:' + x.etaDate + ')' : ''}`);
      const readyCount = partEtas.filter(x => x.eta && !x.delayed).length;
      const readiness = focusPartKeys.length ? Math.round((readyCount / focusPartKeys.length) * 100) : 100;
      const partDetails = blocked.map(x => ({
        partCode: x.code,
        name: x.name,
        demandQty: x.demandQty,
        etaDate: x.etaDate,
        receivedAt: x.receivedAt,
        warehouseStatus: x.warehouseStatus,
        delayed: x.delayed,
      }));

      return { equipmentId: nodeCode, equipmentName: nodeName, level: nodeCode.split('.').length, readiness, blockedParts, partDetails };
    }).sort((a, b) => {
      if (a.readiness < 100 && b.readiness >= 100) return -1;
      if (a.readiness >= 100 && b.readiness < 100) return 1;
      if (a.readiness !== b.readiness) return a.readiness - b.readiness;
      return a.equipmentId.localeCompare(b.equipmentId, undefined, { numeric: true });
    });
  }, [filteredBoms, procurementDeliveryRows, selectedProjectId]);

  const totalNodes = readinessRows.length;
  const readyNodes = readinessRows.filter(r => r.readiness >= 100).length;
  const blockedNodes = readinessRows.filter(r => r.readiness < 100).length;
  const avgReadiness = totalNodes ? Math.round(readinessRows.reduce((s, r) => s + r.readiness, 0) / totalNodes) : 0;

  const toggleExpand = (nodeId: string) => {
    setExpandedNodes(prev => {
      const next = new Set(prev);
      if (next.has(nodeId)) next.delete(nodeId); else next.add(nodeId);
      return next;
    });
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <h2 className="text-lg font-black">{language === 'zh' ? '齐套检查' : 'Kit Readiness'}</h2>
        <div className="flex items-center gap-2">
          <select value={selectedProjectId} onChange={e => setSelectedProjectId(e.target.value)} className="input-field h-9 py-1 px-3 text-sm min-w-[160px]">
            {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
          <div className="flex bg-slate-100 rounded-lg p-0.5">
            {(['Mechanical', 'Electrical'] as const).map(t => (
              <button key={t} onClick={() => setSelectedBomType(t)} className={cn("px-3 py-1.5 rounded-md text-xs font-black transition-all", selectedBomType === t ? "bg-white shadow text-slate-800" : "text-slate-500 hover:text-slate-700")}>
                {t === 'Mechanical' ? (language === 'zh' ? '机械' : 'Mechanical') : (language === 'zh' ? '电气' : 'Electrical')}
              </button>
            ))}
          </div>
        </div>
      </div>

      <div className="grid grid-cols-4 gap-3">
        <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-4 text-center"><p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? '装配单元' : 'Units'}</p><p className="text-2xl font-black mt-1">{totalNodes}</p></div>
        <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-4 text-center"><p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? '齐套' : 'Ready'}</p><p className="text-2xl font-black mt-1 text-green-600">{readyNodes}</p></div>
        <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-4 text-center"><p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? '缺料' : 'Blocked'}</p><p className="text-2xl font-black mt-1 text-rose-600">{blockedNodes}</p></div>
        <div className={cn("rounded-2xl border shadow-sm p-4 text-center", avgReadiness > 90 ? "bg-green-50 border-green-100" : avgReadiness > 70 ? "bg-yellow-50 border-yellow-100" : "bg-red-50 border-red-100")}><p className="text-[10px] font-black text-slate-400 uppercase">{language === 'zh' ? '平均齐套率' : 'Avg Readiness'}</p><p className={cn("text-2xl font-black mt-1", avgReadiness > 90 ? "text-green-700" : avgReadiness > 70 ? "text-yellow-700" : "text-red-700")}>{avgReadiness}%</p></div>
      </div>

      <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead><tr className="bg-slate-50 border-b border-slate-100"><th className="text-left py-2 px-3 font-black text-slate-500 uppercase text-[10px]">{language === 'zh' ? '装配单元' : 'Unit'}</th><th className="text-center py-2 px-3 font-black text-slate-500 uppercase text-[10px] w-20">{language === 'zh' ? '齐套率' : 'Readiness'}</th><th className="text-left py-2 px-3 font-black text-slate-500 uppercase text-[10px]">{language === 'zh' ? '缺料明细' : 'Shortages'}</th></tr></thead>
            <tbody>
              {readinessRows.length === 0 && <tr><td colSpan={3} className="py-8 text-center text-slate-400">{language === 'zh' ? '暂无数据' : 'No data'}</td></tr>}
              {readinessRows.map(row => {
                const isExpanded = expandedNodes.has(row.equipmentId);
                const hasBlocked = row.blockedParts.length > 0;
                return (
                  <React.Fragment key={row.equipmentId}>
                    <tr className={cn("border-b border-slate-50 hover:bg-slate-50/50", hasBlocked && "bg-red-50/30")}>
                      <td className="py-2 px-3"><span className="font-black">{row.equipmentId}</span>{row.equipmentName && <span className="text-slate-500 ml-1">{row.equipmentName}</span>}</td>
                      <td className="py-2 px-3 text-center"><span className={cn("px-2 py-0.5 rounded-lg text-[10px] font-black", row.readiness > 90 ? "bg-green-100 text-green-700" : row.readiness > 70 ? "bg-yellow-100 text-yellow-700" : "bg-red-100 text-red-700")}>{row.readiness}%</span></td>
                      <td className="py-2 px-3">
                        {hasBlocked ? (
                          <div className="flex flex-wrap gap-1">
                            {(isExpanded ? row.blockedParts : row.blockedParts.slice(0, 5)).map((bp, i) => <span key={i} className="px-1.5 py-0.5 rounded-md text-[10px] bg-rose-100 text-rose-700 font-medium">{bp}</span>)}
                            {row.blockedParts.length > 5 && <button onClick={() => toggleExpand(row.equipmentId)} className="px-1.5 py-0.5 rounded-md text-[10px] font-black text-blue-600 hover:underline">{isExpanded ? (language === 'zh' ? '收起' : 'Collapse') : `${language === 'zh' ? '更多' : '+'}${row.blockedParts.length - 5}`}</button>}
                          </div>
                        ) : <span className="text-green-600 font-medium text-[10px]">{language === 'zh' ? '齐套' : 'Ready'}</span>}
                      </td>
                    </tr>
                    {isExpanded && row.partDetails.length > 0 && (
                      <tr key={`${row.equipmentId}-detail`}>
                        <td colSpan={3} className="py-2 px-3 bg-slate-50/50">
                          <table className="w-full text-[10px]">
                            <thead><tr className="border-b border-slate-200"><th className="text-left py-1 px-2 font-black text-slate-500">{language === 'zh' ? '物料编码' : 'Part Code'}</th><th className="text-left py-1 px-2 font-black text-slate-500">{language === 'zh' ? '名称' : 'Name'}</th><th className="text-center py-1 px-2 font-black text-slate-500">{language === 'zh' ? '需求量' : 'Demand'}</th><th className="text-center py-1 px-2 font-black text-slate-500">{language === 'zh' ? '预计到货' : 'ETA'}</th><th className="text-center py-1 px-2 font-black text-slate-500">{language === 'zh' ? '状态' : 'Status'}</th></tr></thead>
                            <tbody>
                              {row.partDetails.map((pd, i) => (
                                <tr key={i} className={cn("border-b border-slate-100", pd.delayed && "bg-red-50/40")}>
                                  <td className="py-1 px-2 font-medium">{pd.partCode}</td>
                                  <td className="py-1 px-2 text-slate-600">{pd.name}</td>
                                  <td className="py-1 px-2 text-center">{pd.demandQty}</td>
                                  <td className="py-1 px-2 text-center">{pd.etaDate || '-'}</td>
                                  <td className="py-1 px-2 text-center">{pd.receivedAt ? <span className="px-1.5 py-0.5 rounded-md bg-green-100 text-green-700 font-medium">{language === 'zh' ? '已到货' : 'Received'}</span> : pd.delayed ? <span className="px-1.5 py-0.5 rounded-md bg-red-100 text-red-700 font-medium">{language === 'zh' ? '延迟' : 'Delayed'}</span> : <span className="px-1.5 py-0.5 rounded-md bg-yellow-100 text-yellow-700 font-medium">{language === 'zh' ? '待到货' : 'Pending'}</span>}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// --- Procurement View ---
export function ProcurementView({ defaultTab = 'requirements' }: { defaultTab?: 'kanban' | 'requirements' | 'bidding' | 'purchasing' | 'analytics' } = {}) {
  const { state, setState, language, handleSaveToDatabase } = useContext(AppContext)!;
  const [activeTab, setActiveTab] = useState<'kanban' | 'requirements' | 'bidding' | 'purchasing' | 'analytics'>(defaultTab);
  const [selectedProjectId, setSelectedProjectId] = useState<string>('');
  const [selectedBomType, setSelectedBomType] = useState<MaterialType>('Mechanical');
  const [selectedBomId, setSelectedBomId] = useState<string>('bom-default');
  
  const t = translations[language];

  const canEditRequirements = state.currentUser?.role === 'admin' || state.currentUser?.role === 'pm' || state.currentUser?.role === 'member';
  const isPurchaser = state.currentUser?.role === 'purchaser' || state.currentUser?.role === 'admin';
  const isApprover = state.currentUser?.role === 'approver' || state.currentUser?.role === 'admin';

  const getBomTypeLabel = (type: MaterialType) => {
    if (language === 'en') return type;
    if (type === 'Mechanical') return '机械';
    if (type === 'Electrical') return '电气';
    if (type === 'Standard') return '标准件';
    if (type === 'Spare') return '备件';
    return type;
  };

  const bomOptions = useMemo(() => {
    if (!selectedProjectId) return [] as Array<{ id: string; name: string; type: MaterialType; sourceLabel?: string }>;
    const list = (state.materialBoms || [])
      .filter(b => b.projectId === selectedProjectId)
      .sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''));
    const opts = list.map(b => ({
      id: b.id,
      name: b.name,
      type: b.type,
      // 统一名称来源：优先使用“设备BOM中心”的当前模板名称（避免改名后采购页仍显示旧名）
      sourceLabel: (() => {
        const tplId = (b as any).importedFromTemplateId;
        if (tplId) {
          const tpl = (state.deviceBomTemplates || []).find((t: any) => t.id === tplId);
          if (tpl) return `来源：${tpl.version || ''}${tpl.name ? ` · ${tpl.name}` : ''}`.trim();
        }
        if ((b as any).importedFromTemplateVersion) {
          return `来源：${(b as any).importedFromTemplateVersion}${(b as any).importedFromTemplateName ? ` · ${(b as any).importedFromTemplateName}` : ''}`.trim();
        }
        return undefined;
      })()
    }));

    // 兼容历史数据：若某类型已有采购行但没有 materialBoms 实体，则补出默认 BOM 占位
    (['Mechanical', 'Electrical', 'Standard', 'Spare'] as MaterialType[]).forEach(type => {
      const hasReq = (state.materialRequirements || []).some((r: any) => r.projectId === selectedProjectId && r.type === type);
      const hasBom = opts.some(o => o.id === 'bom-default' && o.type === type);
      if (hasReq && !hasBom) {
        opts.unshift({
          id: 'bom-default',
          type,
          name: language === 'zh' ? `默认BOM · ${getBomTypeLabel(type)}` : `Default BOM · ${type}`
        });
      }
    });
    return opts;
  }, [state.materialBoms, state.deviceBomTemplates, state.materialRequirements, selectedProjectId, language]);

  useEffect(() => {
    if (!selectedProjectId) {
      setSelectedBomType('Mechanical');
      setSelectedBomId('bom-default');
      return;
    }
    if (bomOptions.length === 0) {
      setSelectedBomType('Mechanical');
      setSelectedBomId('bom-default');
      return;
    }
    if (!bomOptions.some(o => o.id === selectedBomId && o.type === selectedBomType)) {
      setSelectedBomId(bomOptions[0].id);
      setSelectedBomType(bomOptions[0].type);
    }
  }, [selectedProjectId, selectedBomType, bomOptions, selectedBomId]);

  const handleCreateBom = () => {
    if (!selectedProjectId) {
      alert(language === 'zh' ? '请先选择项目' : 'Please select a project first');
      return;
    }
    const nextType = (window.prompt(
      language === 'zh' ? '请输入 BOM 类型：Mechanical / Electrical / Standard / Spare' : 'Enter BOM type: Mechanical / Electrical / Standard / Spare',
      selectedBomType
    ) || '').trim() as MaterialType;
    if (!(['Mechanical', 'Electrical', 'Standard', 'Spare'] as string[]).includes(nextType)) {
      alert(language === 'zh' ? 'BOM类型无效' : 'Invalid BOM type');
      return;
    }
    const defaultName = nextType === 'Mechanical'
      ? (language === 'zh' ? '机械BOM' : 'Mechanical BOM')
      : nextType === 'Electrical'
        ? (language === 'zh' ? '电气BOM' : 'Electrical BOM')
        : nextType === 'Standard'
          ? (language === 'zh' ? '标准件BOM' : 'Standard BOM')
          : (language === 'zh' ? '备件BOM' : 'Spare BOM');

    const name = (window.prompt(language === 'zh' ? '请输入新BOM名称' : 'Enter BOM name', defaultName) || '').trim();
    if (!name) return;

    const bom: MaterialBOM = {
      id: `bom-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      projectId: selectedProjectId,
      type: nextType,
      name,
      createdAt: new Date().toISOString()
    };
    const newState = { ...state, materialBoms: [bom, ...(state.materialBoms || [])] };
    setState(newState);
    handleSaveToDatabase(newState);
    setSelectedBomType(nextType);
    setSelectedBomId(bom.id);
  };

  const filteredRequirements = useMemo(() => {
    if (!selectedProjectId) return [];
    let list = (state.materialRequirements || []) as any[];
    list = list.filter(r => r.projectId === selectedProjectId);
    list = list.filter(r => r.type === selectedBomType);
    list = list.filter(r => ((r as any).bomId || 'bom-default') === selectedBomId);
    // Filter by user role/ownership if needed
    if (state.currentUser?.role === 'member') {
      // Members only see projects where they are potentially involved? 
      // For now let's assume they see all for their project
    }
    return list;
  }, [state.materialRequirements, selectedProjectId, selectedBomType, selectedBomId, state.currentUser]);

  const stats = useMemo(() => {
    const all = state.materialRequirements || [];
    const counts = {
      total: all.length,
      pending: all.filter(r => r.status === 'Pending Review').length,
      approved: all.filter(r => r.status === 'Approved').length,
      bidding: all.filter(r => r.status === 'Bidding').length,
      po: all.filter(r => r.status === 'PO Issued').length,
      delivered: all.filter(r => r.status === 'Delivered').length,
    };
    return counts;
  }, [state.materialRequirements]);

  return (
    <div className="space-y-6 w-full max-w-full pb-10">
      <div className="flex items-center gap-2">
        <TabButton
          active={activeTab === 'analytics'}
          onClick={() => setActiveTab('analytics')}
          icon={<AlertTriangle size={14} />}
          label={language === 'zh' ? '缺料/延期看板' : 'Shortage & Delay'}
        />
        <TabButton
          active={activeTab === 'requirements'}
          onClick={() => setActiveTab('requirements')}
          icon={<List size={14} />}
          label={language === 'zh' ? '采购表格' : 'Tracking Table'}
        />
      </div>

      {activeTab === 'analytics' && (
        <ProcurementShortageDashboard
          state={state}
          language={language}
          projectIdFilter={selectedProjectId}
          onProjectFilterChange={setSelectedProjectId}
        />
      )}

      {activeTab === 'requirements' && (
        <div className="w-full">
          <TrackingTable 
            requirements={filteredRequirements} 
            projectId={selectedProjectId} 
            bomType={selectedBomType}
            bomId={selectedBomId}
            bomOptions={bomOptions}
            projectOptions={state.projects.filter(p => p.id !== 'p-general' && p.id !== 'p-leave')}
            onProjectChange={setSelectedProjectId}
            onBomChange={(id, type) => {
              setSelectedBomId(id);
              setSelectedBomType(type);
            }}
            onCreateBom={handleCreateBom}
            canCreateBom={canEditRequirements}
          />
        </div>
      )}
    </div>
  );
}

function TabButton({ active, onClick, icon, label }: { active: boolean, onClick: () => void, icon: React.ReactNode, label: string }) {
  return (
    <button 
      onClick={onClick}
      className={cn(
        "flex items-center gap-2 px-6 py-2.5 rounded-xl font-black text-sm transition-all",
        active ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:bg-slate-200/50"
      )}
    >
      {icon}
      {label}
    </button>
  );
}

function ProcurementKanban({ requirements }: { requirements: MaterialRequirement[] }) {
  const { language } = useContext(AppContext)!;
  const t = translations[language];

  const columns = [
    { status: 'Pending Review', label: language === 'zh' ? '待审核' : 'Review' },
    { status: 'Approved', label: language === 'zh' ? '已核准' : 'Approved' },
    { status: 'Bidding', label: language === 'zh' ? '处理中' : 'Processing' },
    { status: 'PO Issued', label: language === 'zh' ? '已下单' : 'Ordered' },
    { status: 'In Transit', label: language === 'zh' ? '运输中' : 'Transit' },
  ];

  return (
    <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-6">
      {columns.map(col => (
        <div key={col.status} className="space-y-4">
          <div className="flex items-center justify-between px-2">
            <h3 className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-slate-300" />
              {col.label}
            </h3>
            <span className="text-[10px] font-black text-slate-300">
              {requirements.filter(r => r.status === col.status).length}
            </span>
          </div>
          <div className="space-y-4 min-h-[400px]">
            {requirements.filter(r => r.status === col.status).map(req => (
              <MaterialCard key={req.id} requirement={req} />
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

function MaterialCard({ requirement }: { requirement: MaterialRequirement }) {
  const { state, language } = useContext(AppContext)!;
  const project = state.projects.find(p => p.id === requirement.projectId);
  const t = translations[language];

  return (
    <motion.div 
      layoutId={requirement.id}
      className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm hover:shadow-md transition-all cursor-pointer group"
    >
      <div className="flex items-start justify-between gap-1 mb-3">
        <span className={cn(
          "text-[9px] font-black uppercase px-2 py-0.5 rounded-full",
          requirement.type === 'Mechanical'
            ? "bg-purple-50 text-purple-600"
            : requirement.type === 'Electrical'
              ? "bg-blue-50 text-blue-600"
              : requirement.type === 'Spare'
                ? "bg-rose-50 text-rose-700"
                : "bg-amber-50 text-amber-700"
        )}>
          {requirement.type === 'Mechanical'
            ? t.mechanical
            : requirement.type === 'Electrical'
              ? t.electrical
              : requirement.type === 'Spare'
                ? (language === 'zh' ? '备件' : 'Spare')
                : (language === 'zh' ? '标准件' : 'Standard')}
        </span>
        <span className={cn(
          "text-[9px] font-black uppercase px-2 py-0.5 rounded-full",
          requirement.urgency === 'Urgent' ? "bg-rose-50 text-rose-600" : 
          requirement.urgency === 'High' ? "bg-orange-50 text-orange-600" : "bg-slate-50 text-slate-400"
        )}>
          {requirement.urgency}
        </span>
      </div>
      <h4 className="font-black text-slate-900 text-sm line-clamp-2 leading-tight group-hover:text-blue-600">
        {requirement.name}
      </h4>
      <p className="text-[10px] text-slate-400 font-bold mt-1 truncate">{requirement.specs}</p>
      
      <div className="mt-4 pt-4 border-t border-slate-50 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="w-5 h-5 bg-slate-100 rounded flex items-center justify-center">
            <Briefcase size={10} className="text-slate-400" />
          </div>
          <span className="text-[10px] font-bold text-slate-500 truncate max-w-[80px]">
            {project?.name || 'Unknown'}
          </span>
        </div>
        <div className="flex items-center gap-1 text-slate-900 font-black">
          <span className="text-xs">{requirement.quantity}</span>
          <span className="text-[8px] text-slate-400 uppercase">PCS</span>
        </div>
      </div>
      <div className="mt-2 flex items-center gap-1.5 text-[9px] font-bold text-slate-400">
        <Clock size={10} />
        {format(parseISO(requirement.expectedArrival), 'MM/dd')}
      </div>
    </motion.div>
  );
}

function TrackingTable({
  requirements,
  projectId,
  bomType,
  bomId,
  bomOptions,
  projectOptions,
  onProjectChange,
  onBomChange,
  onCreateBom,
  canCreateBom
}: {
  requirements: MaterialRequirement[],
  projectId: string,
  bomType: MaterialType,
  bomId: string,
  bomOptions: Array<{ id: string; name: string; type: MaterialType; sourceLabel?: string }>,
  projectOptions: Project[],
  onProjectChange: (id: string) => void,
  onBomChange: (id: string, type: MaterialType) => void,
  onCreateBom: () => void,
  canCreateBom: boolean
}) {
  const { state, setState, language, handleSaveToDatabase } = useContext(AppContext)!;
  const t = translations[language];
  const hotRef = useRef<any>(null);

  const getBomTypeLabel = (type: MaterialType) => {
    if (language === 'en') return type;
    if (type === 'Mechanical') return '机械';
    if (type === 'Electrical') return '电气';
    if (type === 'Standard') return '标准件';
    if (type === 'Spare') return '备件';
    return type;
  };

  const renameCurrentBom = async () => {
    if (!projectId) {
      alert(language === 'zh' ? '请先选择项目' : 'Please select project');
      return;
    }
    // 默认BOM id 固定为 bom-default，但每个“项目+类型”都应可独立命名，因此匹配需包含 projectId/type
    const existing = (state.materialBoms || []).find((b: any) =>
      b.id === bomId && b.projectId === projectId && b.type === bomType
    );
    const defaultName = bomType === 'Mechanical'
      ? '机械BOM'
      : bomType === 'Electrical'
        ? '电气BOM'
        : bomType === 'Standard'
          ? '标准件BOM'
          : '备件BOM';
    const currentName = existing?.name || (bomOptions?.find(b => b.id === bomId && b.type === bomType)?.name) || (bomId === 'bom-default' ? '默认BOM' : bomId);
    const nextName = (window.prompt(language === 'zh' ? '请输入BOM表名称' : 'Enter BOM name', currentName || defaultName) || '').trim();
    if (!nextName) return;
    const now = new Date().toISOString();
    const list = [...(state.materialBoms || [])];
    const idx = list.findIndex((b: any) => b.id === bomId && b.projectId === projectId && b.type === bomType);
    if (idx >= 0) {
      list[idx] = { ...list[idx], projectId, type: bomType, name: nextName, createdAt: (list[idx] as any).createdAt || now };
    } else {
      list.unshift({ id: bomId, projectId, type: bomType, name: nextName, createdAt: now } as any);
    }
    const newState = { ...state, materialBoms: list };
    setState(newState);
    await handleSaveToDatabase(newState);
    alert(language === 'zh' ? 'BOM名称已更新' : 'BOM name updated');
  };

  const clearCurrentBomRows = async () => {
    if (!projectId) {
      alert(language === 'zh' ? '请先选择项目' : 'Please select project');
      return;
    }
    const targetName = bomOptions?.find(b => b.id === bomId && b.type === bomType)?.name || bomId;
    const count = (state.materialRequirements || []).filter((r: any) =>
      r.projectId === projectId && r.type === bomType && ((r.bomId || 'bom-default') === bomId)
    ).length;
    const ok = window.confirm(language === 'zh'
      ? `确认清空当前BOM所有采购行？\n\n项目：${projectId}\nBOM：${targetName}\n行数：${count}\n\n此操作不可撤销。`
      : `Clear all rows in current BOM?\nProject: ${projectId}\nBOM: ${targetName}\nRows: ${count}`);
    if (!ok) return;
    const kept = (state.materialRequirements || []).filter((r: any) =>
      !(r.projectId === projectId && r.type === bomType && ((r.bomId || 'bom-default') === bomId))
    );
    const newState = { ...state, materialRequirements: kept };
    setState(newState);
    // Server mode: authoritative delete to avoid “删除后刷新又回来”
    try {
      if (connectionType === 'server' && apiBase) {
        const base = String(apiBase || '').replace(/\/+$/, '');
        await fetch(`${base}/api/material-requirements/clear`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ projectId, bomId, bomType, operator: state.currentUser?.id || null })
        });
      }
    } catch (e) {
      // ignore; fallback to save diff (may not delete on outdated backend)
      console.warn('Server clear failed, fallback to diff deleteIds', e);
    }
    await handleSaveToDatabase(newState);
    alert(language === 'zh' ? '已清空当前BOM采购行' : 'Cleared');
  };

  const bomCounts = useMemo(() => {
    if (!projectId) return [] as Array<{ id: string; name: string; type: MaterialType; sourceLabel?: string; count: number }>;
    const counts = new Map<string, number>();
    (state.materialRequirements || [])
      .filter(r => r.projectId === projectId)
      .forEach((r: any) => {
        const bid = (r.bomId || 'bom-default') as string;
        const key = `${r.type}::${bid}`;
        counts.set(key, (counts.get(key) || 0) + 1);
      });
    const fallback = [{ id: 'bom-default', type: bomType, name: language === 'zh' ? '默认BOM' : 'Default BOM' }];
    return (bomOptions.length ? bomOptions : fallback)
      .map(b => ({ id: b.id, type: b.type, name: b.name, sourceLabel: (b as any).sourceLabel, count: counts.get(`${b.type}::${b.id}`) || 0 }));
  }, [state.materialRequirements, projectId, bomType, bomOptions, language]);

  const procurementBomSummary = useMemo(() => {
    if (!projectId) return { bomCount: 0, rowCount: 0 };
    return {
      bomCount: bomCounts.length,
      rowCount: bomCounts.reduce((sum, b) => sum + (b.count || 0), 0)
    };
  }, [projectId, bomCounts]);

  // --- 从“设备BOM中心 → 采购BOM（列表）”导入到当前项目采购BOM（下拉选择，不用输入框） ---
  const [showDeviceBomImport, setShowDeviceBomImport] = useState(false);
  const [importDeviceModelId, setImportDeviceModelId] = useState<string>('');
  const [importTemplateId, setImportTemplateId] = useState<string>('');
  const [importDeviceCount, setImportDeviceCount] = useState<number>(1);

  const deviceModelOptions = useMemo(() => (state.deviceModels || []) as DeviceModel[], [state.deviceModels]);

  const devicePurchasingTplOptions = useMemo(() => {
    if (!importDeviceModelId) return [] as DeviceBOMTemplate[];
    return (state.deviceBomTemplates || [])
      .filter((t: any) =>
        t.deviceModelId === importDeviceModelId
        && t.bomType === bomType
        && (t.templateKind || 'purchasing') === 'purchasing'
      )
      .sort((a: any, b: any) => String(a.version).localeCompare(String(b.version)));
  }, [state.deviceBomTemplates, importDeviceModelId, bomType]);

  useEffect(() => {
    if (!showDeviceBomImport) return;
    if (!importDeviceModelId) {
      const first = deviceModelOptions?.[0]?.id || '';
      if (first) setImportDeviceModelId(first);
    }
  }, [showDeviceBomImport]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    if (!showDeviceBomImport) return;
    if (devicePurchasingTplOptions.length === 0) {
      setImportTemplateId('');
      return;
    }
    if (!devicePurchasingTplOptions.some(t => t.id === importTemplateId)) {
      setImportTemplateId(devicePurchasingTplOptions[devicePurchasingTplOptions.length - 1].id);
    }
  }, [devicePurchasingTplOptions, showDeviceBomImport]); // eslint-disable-line react-hooks/exhaustive-deps

  const doImportFromDevicePurchasingBom = async () => {
    if (!projectId) {
      alert(language === 'zh' ? '请先选择项目' : 'Please select project');
      return;
    }
    if (!importDeviceModelId) return;
    const tpl = devicePurchasingTplOptions.find(t => t.id === importTemplateId);
    if (!tpl) {
      alert(language === 'zh' ? '请选择采购BOM版本' : 'Select template');
      return;
    }
    const cnt = Math.max(1, Math.floor(Number(importDeviceCount) || 1));
    const overwrite = window.confirm(language === 'zh'
      ? `确认导入？\n- 项目：${projectId}\n- BOM：${bomCounts.find(b => b.id === bomId && b.type === bomType)?.name || bomId}\n- 设备型号：${importDeviceModelId}\n- 版本：${tpl.version}\n- 台数：${cnt}\n\n将覆盖当前项目/当前BOM下的所有采购行。`
      : 'Import and overwrite current BOM rows?');
    if (!overwrite) return;

    const lines = (state.deviceBomLines || []).filter((l: any) => l.templateId === tpl.id);
    if (!lines.length) {
      alert(language === 'zh' ? '该采购BOM版本没有明细行。' : 'No lines found.');
      return;
    }

    const nowIso = new Date().toISOString();
    const today = format(new Date(), 'yyyy-MM-dd');
    const tag = `Imported from Device Purchasing BOM: ${importDeviceModelId} ${tpl.version} ×${cnt}`;

    // overwrite current bom rows
    // Server mode: authoritative clear to ensure “覆盖导入后刷新不倍增”
    try {
      if (connectionType === 'server' && apiBase) {
        const base = String(apiBase || '').replace(/\/+$/, '');
        await fetch(`${base}/api/material-requirements/clear`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ projectId, bomId, bomType, operator: state.currentUser?.id || null })
        });
      }
    } catch (e) {
      console.warn('Server clear (before import) failed; continue with diff deleteIds', e);
    }

    const kept = (state.materialRequirements || []).filter((r: any) =>
      !(r.projectId === projectId && r.type === bomType && ((r.bomId || 'bom-default') === bomId))
    );

    const imported: MaterialRequirement[] = lines.map((l: any, idx: number) => {
      const q = (Number(l.quantityPerDevice) || 0) * cnt;
      return {
        id: `req-bomimp-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 6)}`,
        projectId,
        bomId,
        creatorId: state.currentUser?.id || 'system',
        type: bomType as MaterialType,
        code: `BOM-${Math.floor(1000 + Math.random() * 9000)}`,
        investmentOrder: '',
        csOrder: '',
        costCenter: '',
        name: String(l.name || l.partCodeModel || 'ITEM'),
        model: String(l.partCodeModel || ''),
        drawingNo: String((l as any).drawingNo || ''),
        revision: String(l.revision || 'NA'),
        brand: String((l as any).brand || ''),
        specs: String(l.specs || ''),
        quantity: Number.isFinite(q) ? q : 0,
        unit: String(l.unit || ''),
        unitPrice: 0,
        totalPrice: 0,
        supplier: String(l.supplier || ''),
        prNumber: '',
        prCreatedAt: '',
        poNumber: '',
        poCreatedAt: '',
        actualSupplier: '',
        prStatus: '',
        needByDate: '',
        deliveryTime: '',
        expectedArrival: today,
        receivedAt: '',
        warehouseStatus: '',
        repairArrivalAt: '',
        inboundQuantity: 0,
        outboundQuantity: 0,
        inventoryQuantity: 0,
        userName: '',
        useDate: '',
        qualityFeedback: '',
        sourceTemplateVersion: `${currentTemplate.version}${currentTemplate.name ? ` · ${currentTemplate.name}` : ''}`.trim(),
        changeReason: '由设备BOM生成',
        changeReasonNote: `From DeviceBOM ${currentTemplate.deviceModelId} ${currentTemplate.bomType} ${currentTemplate.version} ×${deviceCount}`,
        urgency: 'Medium',
        status: 'Pending Review',
        comments: tag,
        version: 1,
        createdAt: nowIso,
        updatedAt: nowIso,
        customFields: {},
      } as any;
    }).filter(r => String((r as any).model || '').trim());

      // also persist import source on BOM entity (用于采购页显示 BOM 名称/来源)
    const ensureBomEntity = (): MaterialBOM[] => {
      const list = [...(state.materialBoms || [])];
        // 默认BOM id 固定为 bom-default，但每个“项目+类型”都应可独立命名/记录来源
        const idx = list.findIndex(b => b.id === bomId && b.projectId === projectId && b.type === bomType);
      const now = new Date().toISOString();
      const bomNameDefault = language === 'zh' ? '默认BOM' : 'Default BOM';
      const update = (b: MaterialBOM) => ({
        ...b,
        importedFromDeviceModelId: importDeviceModelId,
        importedFromTemplateId: tpl.id,
        importedFromTemplateVersion: tpl.version,
        importedFromTemplateName: tpl.name,
        importedAt: now,
        // 如果BOM名字还是“默认/机械BOM”等通用名，则自动更新为更明确的名称
        name: (() => {
          const n = String(b.name || '').trim();
          const generic = [bomNameDefault, '机械BOM', '电气BOM', '标准件BOM', '备件BOM', 'Mechanical BOM', 'Electrical BOM', 'Standard BOM', 'Spare BOM'];
          if (!n || generic.includes(n)) {
            return `${tpl.version} · ${tpl.name || bomNameDefault}`.trim();
          }
          return n;
        })(),
      });
      if (idx >= 0) {
        list[idx] = update(list[idx]);
        return list;
      }
      // bom-default 之前可能不存在实体，这里创建一个以便保存“来源信息/名称”
      list.unshift(update({
        id: bomId,
        projectId,
        type: bomType,
        name: bomId === 'bom-default' ? bomNameDefault : (bomCounts.find(b => b.id === bomId)?.name || bomId),
        createdAt: now
      } as any));
      return list;
    };

    const newState = { ...state, materialRequirements: [...imported, ...kept], materialBoms: ensureBomEntity() };
    setState(newState);
    await handleSaveToDatabase(newState);
    setShowDeviceBomImport(false);
    alert(language === 'zh' ? `导入完成：${imported.length} 行（已覆盖当前BOM）。` : `Imported ${imported.length} rows.`);
  };

  // NOTE: “生成测试行”仅用于开发验证，生产环境已移除。

  // --- Persist template changes (incl. column widths) reliably ---
  // Column-resize can trigger rapid successive updates; without serialization/debounce,
  // concurrent file writes may finish out-of-order and cause widths not to persist.
  const stateRef = useRef(state);
  useEffect(() => {
    stateRef.current = state;
  }, [state]);
  const templateSaveTimerRef = useRef<any>(null);
  const scheduleTemplateSave = (nextTemplate: MaterialTrackingTemplateColumn[]) => {
    if (templateSaveTimerRef.current) clearTimeout(templateSaveTimerRef.current);
    templateSaveTimerRef.current = setTimeout(() => {
      const latest = stateRef.current;
      const newState = { ...latest, materialTrackingTemplate: nextTemplate };
      void handleSaveToDatabase(newState);
    }, 500);
  };
  useEffect(() => {
    return () => {
      if (templateSaveTimerRef.current) clearTimeout(templateSaveTimerRef.current);
    };
  }, []);

  const role = state.currentUser?.role;
  const isAdmin = role === 'admin';
  const isEngineer = isAdmin || role === 'member' || role === 'pm';
  const isPurchaser = isAdmin || role === 'purchaser';
  const isWarehouse = isAdmin || role === 'warehouse';
  const hasSelectedProject = !!projectId;
  const employees = state.employees.filter(e => e.id !== 'admin-1');

  const [showTemplateEditor, setShowTemplateEditor] = useState(false);
  const [newHeaderName, setNewHeaderName] = useState('');
  const [newHeaderType, setNewHeaderType] = useState<MaterialColumnDataType>('text');
  const [newHeaderGroup, setNewHeaderGroup] = useState<MaterialColumnGroup>('custom');
  const [newHeaderAssignees, setNewHeaderAssignees] = useState<string[]>([]);
  const [hotData, setHotData] = useState<any[]>([]);
  const hotDataSeededRef = useRef(false);

  const fullTemplate = useMemo(
    () => {
      const tpl = (state.materialTrackingTemplate?.length ? state.materialTrackingTemplate : DEFAULT_MATERIAL_TRACKING_TEMPLATE).map(x => ({ ...x }));
      // Auto-migrate: ensure built-in "name" column exists (for requirement prerequisite)
      if (!tpl.some(c => c.key === 'name')) {
        const modelIdx = tpl.findIndex(c => c.key === 'model');
        const insertAt = modelIdx >= 0 ? modelIdx + 1 : Math.min(5, tpl.length);
        tpl.splice(insertAt, 0, { id: 'mt-name', key: 'name', titleZh: '名称', titleEn: 'Name', dataType: 'text', width: 180, group: 'engineering', assigneeIds: [], enabled: true, builtIn: true });
      }
      // Auto-migrate: ensure built-in "reportedAt" column exists (traceability)
      if (!tpl.some(c => c.key === 'reportedAt')) {
        const prAtIdx = tpl.findIndex(c => c.key === 'prCreatedAt');
        const insertAt = prAtIdx >= 0 ? prAtIdx + 1 : Math.min(10, tpl.length);
        tpl.splice(insertAt, 0, { id: 'mt-reportedAt', key: 'reportedAt', titleZh: '申报时间', titleEn: 'Reported At', dataType: 'date', width: 110, group: 'purchasing', assigneeIds: [], enabled: true, builtIn: true });
      }
      return tpl;
    },
    [state.materialTrackingTemplate]
  );
  const activeTemplate = useMemo(() => fullTemplate.filter(col => col.enabled), [fullTemplate]);
  const idColumnIndex = activeTemplate.length;
  const keyIndexMap = useMemo(() => {
    const map: Record<string, number> = {};
    activeTemplate.forEach((col, idx) => { map[col.key] = idx; });
    return map;
  }, [activeTemplate]);
  const statusOptions = useMemo(() => ([
    '需求释放', '招标中', 'PR创建', 'PR审批', 'PO创建', 'PO审批', '已到货'
  ]), []);
  const privilegedPriceUsers = useMemo(
    () => state.employees
      .filter(e => {
        const n = `${e.nameEn || ''}`.toLowerCase();
        return n.includes('yihong') || n.includes('weijun') || n.includes('jianning');
      })
      .map(e => e.id),
    [state.employees]
  );

  const getColumnLabel = (col: MaterialTrackingTemplateColumn) => language === 'zh' ? col.titleZh : col.titleEn;
  const normalizeHeaderToKey = (name: string) =>
    `custom_${name.trim().toLowerCase().replace(/[^a-z0-9\u4e00-\u9fa5]+/g, '_')}_${Date.now()}`;
  const builtInMaterialKeys = new Set([
    'investmentOrder', 'csOrder', 'costCenter', 'quantity', 'model', 'name', 'supplier', 'prNumber', 'prCreatedAt',
    'reportedAt',
    'poNumber', 'poCreatedAt', 'actualSupplier', 'prStatus', 'unitPrice', 'totalPrice',
    'unit', 'drawingNo', 'revision', 'needByDate', 'deliveryTime', 'receivedAt', 'warehouseStatus', 'repairArrivalAt',
    'inboundQuantity', 'outboundQuantity', 'inventoryQuantity', 'userName', 'useDate', 'qualityFeedback'
  ]);

  const getColumnValue = (req: MaterialRequirement, key: string) => {
    const builtin = (req as any)[key];
    if (builtin !== undefined) return builtin;
    return req.customFields?.[key] ?? '';
  };

  const defaultByType = (type: MaterialColumnDataType) => (type === 'numeric' ? 0 : '');

  const canEditByGroup = (group: MaterialColumnGroup) => {
    if (group === 'engineering') return isEngineer;
    if (group === 'purchasing') return isPurchaser;
    if (group === 'warehouse') return isWarehouse;
    if (group === 'quality') return isEngineer || isWarehouse;
    return isAdmin;
  };

  const canEditColumn = (col: MaterialTrackingTemplateColumn) => {
    if (!hasSelectedProject) return false;
    if (isAdmin) return true;
    if (col.assigneeIds.length > 0) {
      return !!state.currentUser && col.assigneeIds.includes(state.currentUser.id);
    }
    return canEditByGroup(col.group);
  };

  // 需求前置条件：只有“需求字段”填完，后续列才允许填写
  // 你确认的需求字段：零件代号/型号、名称、数量、需求日期（needByDate）
  const requirementPrereqKeys = useMemo(() => (['model', 'name', 'quantity', 'needByDate']), []);
  const prereqKeySet = useMemo(() => new Set(requirementPrereqKeys), [requirementPrereqKeys]);

  const isFilled = (key: string, v: any) => {
    if (v === null || v === undefined) return false;
    if (key === 'quantity') {
      const n = Number(v);
      return Number.isFinite(n) && n > 0;
    }
    const s = String(v).trim();
    return s.length > 0;
  };

  const getMissingPrereqsForRow = (rowIndex: number) => {
    const row = hotData[rowIndex];
    if (!row) return requirementPrereqKeys;
    const missing: string[] = [];
    for (const k of requirementPrereqKeys) {
      const idx = keyIndexMap[k];
      const v = idx === undefined ? undefined : row[idx];
      if (!isFilled(k, v)) missing.push(k);
    }
    return missing;
  };

  const getKeyLabelByKey = (key: string) => {
    const col = activeTemplate.find(c => c.key === key);
    if (col) return getColumnLabel(col);
    // fallback for robustness
    if (key === 'model') return language === 'zh' ? '零件代号/型号' : 'Part Code/Model';
    if (key === 'name') return language === 'zh' ? '名称' : 'Name';
    if (key === 'quantity') return language === 'zh' ? '数量' : 'Qty';
    if (key === 'needByDate') return language === 'zh' ? '需求日期' : 'Need-by Date';
    return key;
  };
  const canViewColumn = (col: MaterialTrackingTemplateColumn) => {
    if (isAdmin) return true; // 超级管理可见全部
    if (col.key === 'unitPrice' || col.key === 'totalPrice') {
      return !!state.currentUser && privilegedPriceUsers.includes(state.currentUser.id);
    }
    return true;
  };
  const editableColumnMap = useMemo(() => activeTemplate.map(col => canEditColumn(col)), [activeTemplate, state.currentUser, hasSelectedProject, role]);
  const hiddenColumnIndexes = useMemo(() => {
    const hidden = [idColumnIndex];
    activeTemplate.forEach((col, idx) => {
      if (!canViewColumn(col)) hidden.push(idx);
    });
    return hidden;
  }, [activeTemplate, idColumnIndex, state.currentUser, privilegedPriceUsers, isAdmin]);
  const quantityLikeKeys = useMemo(() => new Set(['quantity', 'inboundQuantity', 'outboundQuantity', 'inventoryQuantity']), []);

  const setTemplateAndSave = (nextTemplate: MaterialTrackingTemplateColumn[]) => {
    setState(prev => ({ ...prev, materialTrackingTemplate: nextTemplate }));
    scheduleTemplateSave(nextTemplate);
  };

  const syncVisibleColumnWidthsToTemplate = () => {
    const hotInstance = hotRef.current?.hotInstance;
    if (!hotInstance) return;
    const nextTemplate = fullTemplate.map(col => ({ ...col }));
    let changed = false;

    activeTemplate.forEach((col, physicalIdx) => {
      const visualIdx = hotInstance.toVisualColumn ? hotInstance.toVisualColumn(physicalIdx) : physicalIdx;
      if (visualIdx === undefined || visualIdx === null || visualIdx < 0) return;
      const width = hotInstance.getColWidth(visualIdx);
      if (typeof width === 'number' && width > 20) {
        const target = nextTemplate.find(t => t.id === col.id);
        if (target && target.width !== width) {
          target.width = width;
          changed = true;
        }
      }
    });

    if (changed) {
      setTemplateAndSave(nextTemplate);
    }
  };

  const getCellByKey = (row: any[], key: string) => {
    const idx = keyIndexMap[key];
    if (idx === undefined) return '';
    return row[idx];
  };

  const calcAutoStatus = (row: any[]) => {
    const warehouseStatus = `${getCellByKey(row, 'warehouseStatus') || ''}`;
    const receivedAt = `${getCellByKey(row, 'receivedAt') || ''}`;
    const poNumber = `${getCellByKey(row, 'poNumber') || ''}`;
    const poCreatedAt = `${getCellByKey(row, 'poCreatedAt') || ''}`;
    const prNumber = `${getCellByKey(row, 'prNumber') || ''}`;
    const prCreatedAt = `${getCellByKey(row, 'prCreatedAt') || ''}`;
    const qty = Number(getCellByKey(row, 'quantity')) || 0;
    const model = `${getCellByKey(row, 'model') || ''}`;
    const investmentOrder = `${getCellByKey(row, 'investmentOrder') || ''}`;
    const bidderHint = `${getCellByKey(row, 'actualSupplier') || ''}`;

    if (warehouseStatus.includes('到货') || warehouseStatus.includes('进库') || !!receivedAt) return '已到货';
    if (poNumber && poCreatedAt) return 'PO审批';
    if (poNumber) return 'PO创建';
    if (prNumber && prCreatedAt) return 'PR审批';
    if (prNumber) return 'PR创建';
    if (bidderHint) return '招标中';
    if (qty > 0 || model || investmentOrder) return '需求释放';
    return '';
  };

  useEffect(() => {
    // 性能优化：切换 BOM 时避免 React 侧大量 setState 触发 Handsontable 重建。
    // 1) 不再手工补空行（交给 Handsontable 的 minRows/minSpareRows）
    // 2) 表格已初始化后，直接 hot.loadData + suspendRender（切换更顺滑）
    const rows = requirements.map((req) => {
      const row: any[] = [];
      activeTemplate.forEach(col => row.push(getColumnValue(req, col.key)));
      row.push(req.id);
      return row;
    });

    const hotInstance = hotRef.current?.hotInstance;
    if (hotInstance && hotDataSeededRef.current) {
      try {
        // 性能优化（且避免 updateSettings 切换插件导致的 getSortOrderOfColumn 报错）：
        // 不用 updateSettings 来开关插件，改为直接 enable/disable plugin。
        const pFilters = hotInstance.getPlugin?.('filters');
        const pDropdown = hotInstance.getPlugin?.('dropdownMenu');
        const pSorting = hotInstance.getPlugin?.('columnSorting');
        try { pFilters?.disablePlugin?.(); } catch {}
        try { pDropdown?.disablePlugin?.(); } catch {}
        try { pSorting?.disablePlugin?.(); } catch {}

        hotInstance.batch?.(() => {
          hotInstance.suspendRender?.();
          hotInstance.loadData(rows);
          hotInstance.resumeRender?.();
        });

        // 让“先显示出来”，再异步恢复插件（避免选择项目后卡很久才刷出）
        window.setTimeout(() => {
          try { pFilters?.enablePlugin?.(); } catch {}
          try { pDropdown?.enablePlugin?.(); } catch {}
          try { pSorting?.enablePlugin?.(); } catch {}
          // 防止排序插件残留状态导致内部方法访问 undefined
          try { hotInstance.getPlugin?.('columnSorting')?.clearSort?.(); } catch {}
        }, 0);
        hotInstance.refreshDimensions?.();
        hotInstance.render?.();
      } catch (e) {
        setHotData(rows);
      }
      return;
    }

    hotDataSeededRef.current = true;
    setHotData(rows);
  }, [requirements, activeTemplate]);

  useEffect(() => {
    const timer = setTimeout(() => {
      const hotInstance = hotRef.current?.hotInstance;
      if (hotInstance) {
        hotInstance.refreshDimensions?.();
        hotInstance.render?.();
      }
    }, 60);
    return () => clearTimeout(timer);
  }, [hotData, projectId, activeTemplate.length]);

  const handleSave = () => {
    if (!hasSelectedProject) {
      alert(language === 'zh' ? '请选择具体项目再保存' : 'Please select a specific project to save');
      return;
    }
    const hotInstance = hotRef.current?.hotInstance;
    if (!hotInstance) return;

    const currentData = hotInstance.getData();
    const updatedRequirements = [...state.materialRequirements];
    let unauthorizedChangesDetected = false;
    let hasChanges = false;
    const validationErrors: string[] = [];
    const today = format(new Date(), 'yyyy-MM-dd');

    currentData.forEach((row: any) => {
      const id = row[idColumnIndex];
      const reqData: Partial<MaterialRequirement> = { customFields: {} };
      let hasAnyValue = false;
      const existingReq = id ? updatedRequirements.find(r => r.id === id) : undefined;

      if (existingReq) {
        activeTemplate.forEach((col, idx) => {
          if (editableColumnMap[idx]) return;
          const currentVal = row[idx];
          const originalVal = getColumnValue(existingReq, col.key);
          if ((currentVal ?? '') !== (originalVal ?? '')) {
            unauthorizedChangesDetected = true;
          }
        });
      }

      activeTemplate.forEach((col, idx) => {
        const canWrite = editableColumnMap[idx] || col.key === 'totalPrice';
        if (!canWrite) return;
        const rawVal = row[idx];
        const hasValue = col.dataType === 'numeric'
          ? !isNaN(Number(rawVal)) && Number(rawVal) > 0
          : rawVal !== '' && rawVal !== null && rawVal !== undefined;
        if (hasValue) hasAnyValue = true;
        const value = col.dataType === 'numeric' ? (Number(rawVal) || 0) : rawVal;

        if (builtInMaterialKeys.has(col.key)) {
          (reqData as any)[col.key] = value;
        } else {
          (reqData.customFields as any)[col.key] = value;
        }
      });

      const qtyVal = Number(getCellByKey(row, 'quantity')) || 0;
      const priceVal = Number(getCellByKey(row, 'unitPrice')) || 0;
      (reqData as any).totalPrice = qtyVal * priceVal;
      (reqData as any).prStatus = calcAutoStatus(row);

      if (!hasAnyValue) return;

      // --- Validation: CS Order / Cost Center ---
      // Rule: 归属来源优先级：
      // 1) 行 csOrder（手工） 2) 行 costCenter 3) 项目 csOrder（可为临时号 项目号+00x）
      // 需求阶段允许不手填 csOrder/costCenter；若项目层已有 csOrder，则保存时自动继承到该行。
      // 只有当进入后续阶段（PR/PO/入库/出库 等）仍无任何归属时，才阻止保存。
      const csOrderVal = `${getCellByKey(row, 'csOrder') || ''}`.trim();
      const costCenterVal = `${getCellByKey(row, 'costCenter') || ''}`.trim();
      const projectCsOrder = (() => {
        // procurement 页是按 projectId 过滤的，projectId 为当前选中项目
        const p = state.projects.find(p => p.id === projectId);
        return `${(p as any)?.csOrder || ''}`.trim();
      })();

      // Auto inherit: if both empty, inherit project csOrder (if present)
      let effectiveCsOrder = csOrderVal;
      let effectiveCostCenter = costCenterVal;
      if (!effectiveCsOrder && !effectiveCostCenter && projectCsOrder) {
        effectiveCsOrder = projectCsOrder;
        // 写回到保存数据：避免后续追踪丢失归属
        (reqData as any).csOrder = projectCsOrder;
      }

      const outboundVal = Number(getCellByKey(row, 'outboundQuantity')) || 0;
      const isNewRow = !id;
      const existingHasOwner = !!(existingReq && ((existingReq as any).csOrder || (existingReq as any).costCenter));
      const prNumberVal = `${getCellByKey(row, 'prNumber') || ''}`.trim();
      const prCreatedAtVal = `${getCellByKey(row, 'prCreatedAt') || ''}`.trim();
      const poNumberVal = `${getCellByKey(row, 'poNumber') || ''}`.trim();
      const poCreatedAtVal = `${getCellByKey(row, 'poCreatedAt') || ''}`.trim();
      const receivedAtVal = `${getCellByKey(row, 'receivedAt') || ''}`.trim();
      const warehouseStatusVal = `${getCellByKey(row, 'warehouseStatus') || ''}`.trim();
      const inboundVal = Number(getCellByKey(row, 'inboundQuantity')) || 0;

      const hasFollowUp =
        outboundVal > 0 ||
        inboundVal > 0 ||
        !!receivedAtVal ||
        !!warehouseStatusVal ||
        !!poNumberVal ||
        !!poCreatedAtVal ||
        !!prNumberVal ||
        !!prCreatedAtVal;

      const hasOwnerNow = !!effectiveCsOrder || !!effectiveCostCenter;

      if (hasFollowUp && !hasOwnerNow) {
        validationErrors.push(language === 'zh'
          ? '该行进入后续阶段（PR/PO/入库/出库）前必须补齐 CS Order 或 成本中心（或在项目上维护 CS Order 供自动继承）。'
          : 'Before PR/PO/Inbound/Outbound, CS Order or Cost Center is required (or maintain project CS Order for auto inherit).');
        return;
      }

      if (!isNewRow && !existingHasOwner && !hasOwnerNow) {
        // Legacy data: allow saving other fields without forcing backfill, unless follow-up data exists (handled above)
      }

      if ((reqData as any).deliveryTime && !(reqData as any).expectedArrival) {
        (reqData as any).expectedArrival = (reqData as any).deliveryTime;
      }
      if ((reqData as any).warehouseStatus) {
        const ws = (reqData as any).warehouseStatus;
        (reqData as any).status = (ws === 'Already in stock' || ws === '已进库' || ws === '已到货') ? 'Delivered' : 'PO Issued';
      }
      if (!(reqData as any).name) {
        (reqData as any).name = (reqData as any).model || 'ITEM';
      }

      if (id) {
        const idx = updatedRequirements.findIndex(r => r.id === id);
        if (idx > -1) {
          // Only mark changed when the editable fields differ (avoid “点保存就更新一遍”带来的卡顿与追溯噪音)
          let changed = false;
          const base = updatedRequirements[idx] as any;
          activeTemplate.forEach((col, cIdx) => {
            const canWrite = editableColumnMap[cIdx] || col.key === 'totalPrice';
            if (!canWrite) return;
            const rawVal = row[cIdx];
            const nv = col.dataType === 'numeric' ? (Number(rawVal) || 0) : (rawVal ?? '');
            const ov = getColumnValue(base, col.key);
            const ovn = col.dataType === 'numeric' ? (Number(ov) || 0) : (ov ?? '');
            if (String(nv) !== String(ovn)) changed = true;
          });
          // auto fields
          if (String((reqData as any).prStatus ?? '') !== String((base as any).prStatus ?? '')) changed = true;
          if (Number((reqData as any).totalPrice || 0) !== Number((base as any).totalPrice || 0)) changed = true;
          // auto-inherit CS Order
          if ((reqData as any).csOrder !== undefined && String((reqData as any).csOrder || '') !== String((base as any).csOrder || '')) changed = true;

          if (!changed) return;
          updatedRequirements[idx] = {
            ...updatedRequirements[idx],
            bomId: (updatedRequirements[idx] as any).bomId || bomId,
            ...reqData,
            customFields: { ...(updatedRequirements[idx].customFields || {}), ...(reqData.customFields || {}) },
            reportedAt: today,
            updatedAt: new Date().toISOString()
          };
          hasChanges = true;
        }
      } else {
        const newReq: MaterialRequirement = {
          id: `req-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
          projectId,
          bomId,
          creatorId: state.currentUser?.id || 'system',
          type: bomType,
          code: `M-${Math.floor(1000 + Math.random() * 9000)}`,
          specs: '',
          urgency: 'Medium',
          investmentOrder: '',
          name: '',
          model: '',
          brand: '',
          supplier: '',
          quantity: 0,
          unit: '',
          unitPrice: 0,
          totalPrice: 0,
          reportedAt: today,
          expectedArrival: format(new Date(), 'yyyy-MM-dd'),
          status: 'Pending Review',
          comments: '',
          version: 1,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
          customFields: {},
          ...reqData
        };
        updatedRequirements.push(newReq);
        hasChanges = true;
      }
    });

    if (unauthorizedChangesDetected) {
      alert(language === 'zh' ? '检测到未授权列被修改，已拦截保存。请联系管理员调整列分配。' : 'Unauthorized column changes detected. Save was blocked.');
      return;
    }
    if (validationErrors.length) {
      alert(language === 'zh'
        ? `保存失败：\n${Array.from(new Set(validationErrors)).slice(0, 5).join('\n')}`
        : `Save blocked:\n${Array.from(new Set(validationErrors)).slice(0, 5).join('\n')}`);
      return;
    }

    if (hasChanges) {
      const newState = { ...state, materialRequirements: updatedRequirements };
      setState(newState);
      handleSaveToDatabase(newState);
      alert(language === 'zh' ? '保存成功' : 'Saved successfully');
    }
  };

  const handleImportExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);
      const pick = (row: any, keys: string[]) => {
        for (const k of keys) {
          const v = row?.[k];
          if (v !== undefined && v !== null && String(v).trim() !== '') return v;
        }
        return '';
      };

      const normalize = (v: any) => String(v ?? '').trim();
      const normalizeKey = (v: any) => normalize(v).replace(/\s+/g, ' ').toLowerCase();
      const normRev = (v: any) => normalize(v) || 'NA';
      const getMaterialNo = (row: any) =>
        normalize(
          pick(row, [
            '零件代号/型号', '零件代号 / 型号',
            '型号/图号', '型号 / 图号',
            '图号/型号', '图号 / 型号',
            '零件代号', '零件代号/型号(必填)',
            '型号', '图号', '料号', '物料号', 'Part No', 'PartNo', 'P/N', 'PN', 'Model', 'Drawing No'
          ])
        );
      const getProjectCode = (row: any) =>
        normalize(pick(row, ['项目编号', '项目', 'Project', 'Project ID', 'ProjectId']));
      const getCSOrder = (row: any) =>
        normalize(pick(row, ['CS Order', 'CSOrder', 'CSORDER', 'CS_ORDER', 'cs order', 'csorder']));
      const getCostCenter = (row: any) =>
        normalize(pick(row, ['成本中心', 'Cost Center', 'CostCenter', 'COST CENTER', 'cost center']));
      const getNeedByDate = (row: any) =>
        normalize(pick(row, ['需求时间', '需求日期', 'Need-by Date', 'Need Date', 'NeedByDate', 'needByDate']));

      const getUniqKey = (r: { projectId: string; csOrder?: string; costCenter?: string; materialNo: string; revision: string }) => {
        const ref = r.csOrder ? `CS:${normalizeKey(r.csOrder)}` : `CC:${normalizeKey(r.costCenter || '')}`;
        return `${normalizeKey(r.projectId)}|${ref}|${normalizeKey(r.materialNo)}|${normalizeKey(r.revision || 'NA')}`;
      };

      // Build existing key map
      const existingByKey = new Map<string, MaterialRequirement>();
      (state.materialRequirements || []).forEach((r: any) => {
        const materialNo = normalize(r.model || r.drawingNo || r.code || r.name || '');
        const revision = normRev(r.revision);
        const key = getUniqKey({ projectId: r.projectId, csOrder: r.csOrder, costCenter: r.costCenter, materialNo, revision });
        existingByKey.set(key, r);
      });

      // Parse rows
      const errors: Array<{ row: number; message: string }> = [];
      const parsed: Array<{
        row: number;
        projectId: string;
        materialNo: string;
        revision: string;
        csOrder?: string;
        costCenter?: string;
        quantity: number;
        needByDate?: string;
      }> = [];
      const newProjects: Project[] = [];
      const projectById = new Map(state.projects.map(p => [p.id, p] as const));

      data.forEach((row: any, idx: number) => {
        const rowNo = idx + 2; // header is row 1
        const projectCode = getProjectCode(row) || projectId;
        if (!projectCode) {
          errors.push({ row: rowNo, message: '缺少 项目编号' });
          return;
        }
        let project = state.projects.find(p => p.id === projectCode || p.name === projectCode);
        if (!project) {
          // Auto-create (lightweight) project record
          project = { id: projectCode, name: projectCode, csOrder: '', budgets: {}, managerId: state.currentUser?.id || 'admin-1', status: 'active' } as any;
          if (!projectById.has(project.id)) {
            projectById.set(project.id, project);
            newProjects.push(project);
          }
        }

        const csOrder = getCSOrder(row);
        const costCenter = getCostCenter(row);
        // Ensure project has a CS Order (required for project config). If none from sheet, generate temporary: 项目号+00x
        if (!(project as any).csOrder) {
          (project as any).csOrder = csOrder || `${project.id}001`;
        }
        if (!csOrder && !costCenter) {
          errors.push({ row: rowNo, message: 'CS Order 与 成本中心 至少填一个' });
          return;
        }
        const materialNo = getMaterialNo(row);
        if (!materialNo) {
          errors.push({ row: rowNo, message: '缺少 零件代号/型号' });
          return;
        }
        const qtyRaw = pick(row, ['数量', 'Qty', 'Quantity', '需求数量', '需求量']);
        const qty = Number(qtyRaw);
        if (!Number.isFinite(qty) || qty <= 0) {
          errors.push({ row: rowNo, message: '数量必须为 >0 的数字' });
          return;
        }
        const revision = normRev(pick(row, ['版本', '版本号', 'Revision', 'Rev']));
        const needByDate = getNeedByDate(row);

        parsed.push({
          row: rowNo,
          projectId: project.id,
          materialNo,
          revision,
          csOrder: csOrder || undefined,
          costCenter: costCenter || undefined,
          quantity: qty,
          needByDate: needByDate || undefined,
        });
      });

      if (errors.length) {
        const top = errors.slice(0, 10).map(e => `第${e.row}行：${e.message}`).join('\n');
        alert(`${language === 'zh' ? '导入失败：存在必填或格式错误' : 'Import failed'}\n${top}${errors.length > 10 ? `\n...(${errors.length} errors)` : ''}`);
        return;
      }

      // Count duplicates
      const importKeyCounts = new Map<string, number>();
      parsed.forEach(r => {
        const k = getUniqKey({ projectId: r.projectId, csOrder: r.csOrder, costCenter: r.costCenter, materialNo: r.materialNo, revision: r.revision });
        importKeyCounts.set(k, (importKeyCounts.get(k) || 0) + 1);
      });
      const dupInFile = Array.from(importKeyCounts.values()).filter(n => n > 1).reduce((a, b) => a + (b - 1), 0);
      const dupWithExisting = Array.from(importKeyCounts.keys()).filter(k => existingByKey.has(k)).length;

      // Choose strategy if duplicates exist
      let strategy: 'sum' | 'overwrite' = 'sum';
      if (dupInFile > 0 || dupWithExisting > 0) {
        const choice = window.prompt(
          `检测到重复主键：\n- 文件内部重复：${dupInFile} 条\n- 与现有数据重复：${dupWithExisting} 条\n\n请输入：\n1 = 数量求和合并\n2 = 后者整行覆盖\n0 = 取消导入`,
          '1'
        );
        if (!choice || choice.trim() === '0') return;
        strategy = choice.trim() === '2' ? 'overwrite' : 'sum';
        const remember = window.confirm(language === 'zh' ? '是否记住本次选择（下次默认）？' : 'Remember this choice for next time?');
        if (remember) localStorage.setItem('tms_import_dup_strategy', strategy);
      }

      // Apply remembered strategy only when no duplicates prompt shown (optional)
      if ((dupInFile === 0 && dupWithExisting === 0)) {
        const remembered = localStorage.getItem('tms_import_dup_strategy');
        if (remembered === 'sum' || remembered === 'overwrite') strategy = remembered as any;
      }

      // First, resolve duplicates within file
      const grouped = new Map<string, any>();
      parsed.forEach(r => {
        const key = getUniqKey({ projectId: r.projectId, csOrder: r.csOrder, costCenter: r.costCenter, materialNo: r.materialNo, revision: r.revision });
        const prev = grouped.get(key);
        if (!prev) {
          grouped.set(key, { ...r });
          return;
        }
        if (strategy === 'overwrite') {
          grouped.set(key, { ...r }); // keep the latter whole row
        } else {
          // sum
          grouped.set(key, {
            ...prev,
            quantity: (Number(prev.quantity) || 0) + (Number(r.quantity) || 0),
            needByDate: prev.needByDate || r.needByDate,
            csOrder: prev.csOrder || r.csOrder,
            costCenter: prev.costCenter || r.costCenter
          });
        }
      });

      const now = new Date().toISOString();
      const today = format(new Date(), 'yyyy-MM-dd');
      const updatedRequirements = [...state.materialRequirements];
      let createdCount = 0;
      let updatedCount = 0;

      const fillIfEmpty = (oldVal: any, newVal: any) => (normalize(oldVal) ? oldVal : (normalize(newVal) ? newVal : oldVal));

      grouped.forEach((r, key) => {
        const existing = existingByKey.get(key);
        if (existing) {
          if (strategy === 'overwrite') {
            const idx = updatedRequirements.findIndex(x => x.id === existing.id);
            if (idx >= 0) {
              updatedRequirements[idx] = {
                ...existing,
                projectId: r.projectId,
                bomId: (existing as any).bomId || bomId,
                type: existing.type || (bomType as any),
                csOrder: r.csOrder || '',
                costCenter: r.costCenter || '',
                model: r.materialNo,
                revision: r.revision,
                quantity: Number(r.quantity) || 0,
                needByDate: r.needByDate || '',
                comments: 'Imported (overwrite)',
                reportedAt: today,
                updatedAt: now,
              } as any;
              updatedCount++;
            }
          } else {
            // sum (quantity add, other fields fill-if-empty)
            const idx = updatedRequirements.findIndex(x => x.id === existing.id);
            if (idx >= 0) {
              const next = { ...existing } as any;
              next.quantity = (Number(next.quantity) || 0) + (Number(r.quantity) || 0);
              next.csOrder = fillIfEmpty(next.csOrder, r.csOrder || '');
              next.costCenter = fillIfEmpty(next.costCenter, r.costCenter || '');
              next.model = fillIfEmpty(next.model, r.materialNo);
              next.revision = fillIfEmpty(next.revision, r.revision);
              next.needByDate = fillIfEmpty(next.needByDate, r.needByDate || '');
              next.reportedAt = today;
              next.updatedAt = now;
              next.comments = 'Imported (sum merge)';
              updatedRequirements[idx] = next;
              updatedCount++;
            }
          }
        } else {
          const newReq: MaterialRequirement = {
            id: `req-import-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
            projectId: r.projectId,
            bomId,
            creatorId: state.currentUser?.id || 'system',
            type: bomType as MaterialType,
            code: `IMP-${Math.floor(Math.random() * 100000)}`,
            investmentOrder: '',
            csOrder: r.csOrder || '',
            costCenter: r.costCenter || '',
            name: r.materialNo || 'Imported Item',
            model: r.materialNo || '',
            drawingNo: '',
            revision: r.revision,
            brand: '',
            specs: '',
            quantity: Number(r.quantity) || 0,
            unit: '',
            unitPrice: 0,
            totalPrice: 0,
            supplier: '',
            expectedArrival: format(new Date(), 'yyyy-MM-dd'),
            needByDate: r.needByDate || '',
            urgency: 'Medium',
            status: 'Pending Review',
            comments: 'Imported from Excel',
            reportedAt: today,
            version: 1,
            createdAt: now,
            updatedAt: now
          };
          updatedRequirements.unshift(newReq);
          createdCount++;
        }
      });

      const nextState = {
        ...state,
        projects: [...state.projects, ...newProjects],
        materialRequirements: updatedRequirements
      };
      setState(nextState);
      handleSaveToDatabase(nextState);
      alert(language === 'zh'
        ? `导入完成：新增 ${createdCount} 条，更新 ${updatedCount} 条。`
        : `Import done: created ${createdCount}, updated ${updatedCount}.`);
    };
    reader.readAsBinaryString(file);
  };

  const addCustomHeader = () => {
    if (!newHeaderName.trim()) return;
    const key = normalizeHeaderToKey(newHeaderName);
    const newCol: MaterialTrackingTemplateColumn = {
      id: `mt-${Date.now()}`,
      key,
      titleZh: newHeaderName.trim(),
      titleEn: newHeaderName.trim(),
      dataType: newHeaderType,
      width: 120,
      group: newHeaderGroup,
      assigneeIds: newHeaderAssignees,
      enabled: true,
      builtIn: false
    };
    setTemplateAndSave([...fullTemplate, newCol]);
    setNewHeaderName('');
    setNewHeaderAssignees([]);
  };

  return (
    <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden flex flex-col h-[85vh] min-h-[600px]">
      <div className="p-4 border-b border-slate-50 space-y-2">
        <div className="flex flex-col gap-2">
          {/* Row 0: 操作区 */}
          <div className="flex flex-wrap items-center gap-2">
            <button onClick={handleSave} className="bg-blue-600 text-white px-4 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-widest shadow-lg shadow-blue-200 hover:bg-blue-700 transition-all flex items-center gap-2">
              <Save size={12} /> {t.save}
            </button>
            <button
              onClick={() => void renameCurrentBom()}
              className="btn-secondary py-1.5 px-3 text-[10px] flex items-center gap-1"
              disabled={!projectId}
              title={language === 'zh' ? '修改当前BOM表名称（会影响顶部BOM按钮显示）' : 'Rename current BOM'}
            >
              <Pencil size={12} /> {language === 'zh' ? '重命名BOM' : 'Rename BOM'}
            </button>
            <button
              onClick={() => void clearCurrentBomRows()}
              className="bg-rose-600 text-white px-4 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-widest shadow-lg shadow-rose-200 hover:bg-rose-700 transition-all flex items-center gap-2"
              disabled={!projectId}
              title={language === 'zh' ? '清空当前项目/当前BOM下的全部采购行' : 'Clear all rows in current project/BOM'}
            >
              <Trash2 size={12} /> {language === 'zh' ? '清空BOM' : 'Clear BOM'}
            </button>
            {isAdmin && (
              <button
                onClick={() => {
                  if (showTemplateEditor) {
                    syncVisibleColumnWidthsToTemplate();
                  }
                  setShowTemplateEditor(v => !v);
                }}
                className="btn-secondary py-1.5 px-3 text-[10px] flex items-center gap-1"
              >
                <SettingsIcon size={12} />
                {language === 'zh' ? '模板设置' : 'Template'}
              </button>
            )}
            <div className="relative">
              <button className="btn-secondary py-1.5 px-3 text-[10px] flex items-center gap-1">
                <Upload size={12} /> {t.importExcel}
              </button>
              <input
                type="file"
                accept=".xlsx, .xls"
                className="absolute inset-0 opacity-0 cursor-pointer"
                onChange={handleImportExcel}
              />
            </div>
            <button
              onClick={() => {
                if (!projectId) {
                  alert(language === 'zh' ? '请先选择项目' : 'Please select project');
                  return;
                }
                setShowDeviceBomImport(true);
              }}
              className="btn-secondary py-1.5 px-3 text-[10px] flex items-center gap-1"
              disabled={!projectId}
              title={language === 'zh' ? '从设备BOM中心的“采购BOM（列表）”导入到当前项目BOM（会覆盖当前BOM行）' : 'Import from device purchasing BOM'}
            >
              <Upload size={12} /> {language === 'zh' ? '从设备BOM导入' : 'Import from Device BOM'}
            </button>
          </div>

          {/* Row 1: 项目下拉（项目多时不占横向） */}
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-[10px] font-black text-slate-400 shrink-0">{language === 'zh' ? '项目' : 'Project'}</span>
            <select
              value={projectId || ''}
              onChange={e => onProjectChange(e.target.value)}
              className="input-field h-8 py-1 px-2 text-[10px] font-black min-w-[260px]"
            >
              <option value="">{language === 'zh' ? '请选择项目' : 'Select project'}</option>
              {projectOptions.map(p => (
                <option key={p.id} value={p.id}>{p.name}</option>
              ))}
            </select>
            {!!projectId && (
              <span className="text-[10px] font-black text-slate-400">
                {language === 'zh'
                  ? `该项目下共 ${procurementBomSummary.bomCount} 个BOM清单，合计 ${procurementBomSummary.rowCount} 行采购物料`
                  : `${procurementBomSummary.bomCount} BOM lists, ${procurementBomSummary.rowCount} rows`}
              </span>
            )}
          </div>

          {/* Row 2: 直接展示该项目下全部 BOM 清单 */}
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-[10px] font-black text-slate-400">BOM</span>
            <div className="flex items-center gap-1 flex-wrap">
              {bomCounts.map(b => (
                <button
                  key={`${b.type}::${b.id}`}
                  onClick={() => onBomChange(b.id, b.type)}
                  disabled={!projectId}
                  className={cn(
                    "px-3 py-1 rounded-xl text-xs font-black border flex items-center gap-2",
                    bomId === b.id && bomType === b.type ? "bg-slate-900 text-white border-slate-900" : "bg-white text-slate-600 border-slate-200",
                    !projectId ? "opacity-50" : ""
                  )}
                  title={language === 'zh' ? `该BOM下 ${b.count} 行` : `${b.count} rows`}
                >
                  <span className="flex flex-col items-start leading-tight max-w-[220px]">
                    <span className="truncate w-full">{b.name}</span>
                    {b.sourceLabel && (
                      <span className={cn(
                        "text-[10px] font-black truncate w-full",
                        bomId === b.id && bomType === b.type ? "text-white/70" : "text-slate-400"
                      )}>
                        {b.sourceLabel}
                      </span>
                    )}
                  </span>
                  <span className={cn(
                    "text-[10px] font-black px-2 py-0.5 rounded-full",
                    b.type === 'Mechanical' ? "bg-purple-100 text-purple-700" :
                    b.type === 'Electrical' ? "bg-blue-100 text-blue-700" :
                    b.type === 'Standard' ? "bg-amber-100 text-amber-700" :
                    "bg-rose-100 text-rose-700"
                  )}>
                    {getBomTypeLabel(b.type)}
                  </span>
                  <span className={cn("text-[10px] font-black px-2 py-0.5 rounded-full",
                    bomId === b.id && bomType === b.type ? "bg-white/20 text-white" : "bg-slate-100 text-slate-500")}>
                    {b.count}
                  </span>
                </button>
              ))}
            </div>
            {canCreateBom && (
              <button
                onClick={onCreateBom}
                className="btn-secondary py-1.5 px-3 text-[10px] flex items-center gap-1"
                disabled={!projectId}
              >
                <Plus size={12} /> {language === 'zh' ? '新增BOM' : 'New BOM'}
              </button>
            )}
          </div>
        </div>
      </div>

      {/* 从设备BOM导入：下拉选择弹窗 */}
      <Portal>
        <AnimatePresence>
          {showDeviceBomImport && (
            <div className="fixed inset-0 z-[1000] flex items-center justify-center p-4">
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                onClick={() => setShowDeviceBomImport(false)}
                className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
              />
              <motion.div
                initial={{ opacity: 0, scale: 0.96, y: 10 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.96, y: 10 }}
                className="relative bg-white w-full max-w-2xl rounded-[32px] shadow-2xl overflow-hidden"
              >
              <div className="p-6 border-b border-slate-100 bg-slate-50/60 flex items-center justify-between">
                <div>
                  <div className="text-[10px] font-black text-slate-400 uppercase tracking-widest">从设备BOM导入到采购BOM</div>
                  <div className="text-lg font-black text-slate-900 mt-1">
                    {language === 'zh' ? '选择设备型号 / 采购BOM版本 / 台数' : 'Select model/template/count'}
                  </div>
                </div>
                <button onClick={() => setShowDeviceBomImport(false)} className="text-slate-400 hover:text-slate-700">
                  <X size={20} />
                </button>
              </div>

              <div className="p-6 space-y-4">
                <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">设备型号</label>
                    <select
                      value={importDeviceModelId}
                      onChange={e => setImportDeviceModelId(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-4 py-3 font-bold text-sm"
                    >
                      {deviceModelOptions.map(dm => (
                        <option key={dm.id} value={dm.id}>{formatDeviceModelLabel(dm, deviceModelOptions)}</option>
                      ))}
                    </select>
                  </div>

                  <div className="space-y-2 md:col-span-2">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">采购BOM版本（列表）</label>
                    <select
                      value={importTemplateId}
                      onChange={e => setImportTemplateId(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-4 py-3 font-bold text-sm"
                      disabled={devicePurchasingTplOptions.length === 0}
                    >
                      {devicePurchasingTplOptions.length === 0 && (
                        <option value="">{language === 'zh' ? '该设备暂无采购BOM版本（请先在设备BOM中心导入采购BOM）' : 'No template'}</option>
                      )}
                      {devicePurchasingTplOptions.map(tpl => (
                        <option key={tpl.id} value={tpl.id}>
                          {(() => {
                            const importedAt = (tpl as any).importedAt || tpl.createdAt;
                            const changed = importedAt && tpl.updatedAt && tpl.updatedAt > importedAt;
                            const flag = changed ? ' · 已改动' : '';
                            const hist = (tpl as any).sourceHistoryLastDate ? ` · 历史:${(tpl as any).sourceHistoryLastDate}` : '';
                            return `${tpl.version} · ${tpl.name || ''}${hist}${flag}`.trim();
                          })()}
                        </option>
                      ))}
                    </select>
                    {(() => {
                      const tpl = devicePurchasingTplOptions.find(t => t.id === importTemplateId) as any;
                      if (!tpl) return null;
                      const importedAt = tpl.importedAt || tpl.createdAt;
                      const changed = importedAt && tpl.updatedAt && tpl.updatedAt > importedAt;
                      const histParts = [
                        tpl.sourceHistoryLastVersion ? `版本:${tpl.sourceHistoryLastVersion}` : '',
                        tpl.sourceHistoryLastDate ? `日期:${tpl.sourceHistoryLastDate}` : '',
                        tpl.sourceHistoryLastBy ? `更改人:${tpl.sourceHistoryLastBy}` : '',
                        Number.isFinite(tpl.sourceHistoryCount) ? `条数:${tpl.sourceHistoryCount}` : ''
                      ].filter(Boolean);
                      return (
                        <div className="mt-2 text-[11px] text-slate-500 font-bold space-y-1">
                          <div>名称：{tpl.name || '-'}</div>
                          <div>来源文件：{tpl.sourceFileName || '-'}</div>
                          <div>来源历史：{histParts.length ? histParts.join('；') : '-'}</div>
                          <div>系统后续改动：{changed ? `有（最后更新 ${String(tpl.updatedAt).slice(0, 10)}）` : '无'}</div>
                        </div>
                      );
                    })()}
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">计划台数</label>
                    <input
                      type="number"
                      min={1}
                      value={importDeviceCount}
                      onChange={e => setImportDeviceCount(Math.max(1, Math.floor(Number(e.target.value) || 1)))}
                      className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-4 py-3 font-bold text-sm"
                    />
                  </div>
                  <div className="md:col-span-2 rounded-2xl border border-amber-100 bg-amber-50 p-4 text-sm text-amber-900">
                    <div className="font-black text-xs">注意</div>
                    <div className="text-xs font-bold mt-1">
                      将覆盖当前项目「{projectId}」下当前BOM「{bomCounts.find(b => b.id === bomId)?.name || bomId}」的所有采购行。
                    </div>
                  </div>
                </div>
              </div>

              <div className="p-6 border-t border-slate-100 bg-white flex items-center gap-3">
                <button
                  onClick={() => setShowDeviceBomImport(false)}
                  className="flex-1 py-3 rounded-2xl font-black text-sm text-slate-500 hover:bg-slate-50 transition-all border border-slate-200"
                >
                  取消
                </button>
                <button
                  onClick={() => void doImportFromDevicePurchasingBom()}
                  disabled={!importDeviceModelId || !importTemplateId}
                  className={cn(
                    "flex-1 py-3 rounded-2xl font-black text-sm transition-all",
                    (!importDeviceModelId || !importTemplateId)
                      ? "bg-slate-100 text-slate-400"
                      : "bg-slate-900 text-white hover:bg-slate-800"
                  )}
                >
                  开始导入
                </button>
              </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>
      </Portal>

      {isAdmin && showTemplateEditor && (
        <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/60 space-y-3">
          <div className="text-[11px] font-bold text-indigo-600">
            {language === 'zh' ? '全局模板：对所有项目、所有 BOM 类型生效' : 'Global template: applies to all projects and BOM types'}
          </div>
          <div className="flex items-center gap-2">
            <input
              value={newHeaderName}
              onChange={e => setNewHeaderName(e.target.value)}
              placeholder={language === 'zh' ? '新增表头名称' : 'New Header Name'}
              className="input-field h-10 py-2 px-3 text-xs"
            />
            <select value={newHeaderType} onChange={e => setNewHeaderType(e.target.value as MaterialColumnDataType)} className="input-field h-10 py-2 px-3 text-xs">
              <option value="text">Text</option>
              <option value="numeric">Numeric</option>
              <option value="date">Date</option>
            </select>
            <select value={newHeaderGroup} onChange={e => setNewHeaderGroup(e.target.value as MaterialColumnGroup)} className="input-field h-10 py-2 px-3 text-xs">
              <option value="engineering">Engineering</option>
              <option value="purchasing">Purchasing</option>
              <option value="warehouse">Warehouse</option>
              <option value="quality">Quality</option>
              <option value="custom">Custom</option>
            </select>
            <button onClick={addCustomHeader} className="btn-primary h-10 py-2 px-3 text-xs">
              <Plus size={12} /> {language === 'zh' ? '新增表头' : 'Add Header'}
            </button>
          </div>
          <div className="space-y-2 max-h-[220px] overflow-y-auto">
            {fullTemplate.map(col => {
              const assigneeNames = employees
                .filter(emp => col.assigneeIds.includes(emp.id))
                .map(emp => getEmpName(emp, language))
                .join('、');
              return (
                <div key={col.id} className="bg-white border border-slate-200 rounded-xl p-2 flex items-center gap-2">
                  <input
                    value={getColumnLabel(col)}
                    onChange={e => {
                      const next = fullTemplate.map(x => x.id === col.id ? { ...x, titleZh: e.target.value, titleEn: e.target.value } : x);
                      setTemplateAndSave(next);
                    }}
                    className="input-field h-9 py-1 px-2 text-xs min-w-[160px]"
                  />
                  <label className="text-[10px] font-bold text-slate-500 flex items-center gap-1">
                    <input
                      type="checkbox"
                      checked={col.enabled}
                      onChange={e => {
                        const next = fullTemplate.map(x => x.id === col.id ? { ...x, enabled: e.target.checked } : x);
                        setTemplateAndSave(next);
                      }}
                    />
                    {language === 'zh' ? '启用' : 'Enable'}
                  </label>
                  <div className="flex items-center gap-1 flex-wrap">
                    {employees.map(emp => {
                      const selected = col.assigneeIds.includes(emp.id);
                      return (
                        <button
                          key={emp.id}
                          onClick={() => {
                            const nextAssignees = selected ? col.assigneeIds.filter(id => id !== emp.id) : [...col.assigneeIds, emp.id];
                            const next = fullTemplate.map(x => x.id === col.id ? { ...x, assigneeIds: nextAssignees } : x);
                            setTemplateAndSave(next);
                          }}
                          className={cn(
                            "px-2 py-0.5 rounded-md text-[10px] font-bold border",
                            selected ? "bg-blue-600 text-white border-blue-600" : "bg-white text-slate-500 border-slate-200"
                          )}
                        >
                          {getEmpName(emp, language)}
                        </button>
                      );
                    })}
                  </div>
                  <span className="text-[10px] text-slate-400 font-bold min-w-[160px] truncate">
                    {language === 'zh' ? `已分配: ${assigneeNames || '未分配'}` : `Assigned: ${assigneeNames || 'None'}`}
                  </span>
                  <button
                    onClick={() => setTemplateAndSave(fullTemplate.filter(x => x.id !== col.id))}
                    className="ml-auto text-rose-600 hover:text-rose-700 p-1"
                  >
                    <Trash2 size={14} />
                  </button>
                </div>
              );
            })}
          </div>
        </div>
      )}

      <div className="flex-1 overflow-auto p-2 relative bg-white rounded-2xl border border-slate-100 shadow-sm mt-4">
        <HotTable
          ref={hotRef}
          className="custom-hot"
          data={hotData}
          minRows={30}
          minSpareRows={5}
          colHeaders={[...activeTemplate.map(col => getColumnLabel(col)), 'ID']}
          columns={[
            ...activeTemplate.map(col => ({
              type: col.key === 'prStatus'
                ? 'dropdown'
                : col.dataType === 'numeric'
                  ? 'numeric'
                  : col.dataType === 'date'
                    ? 'date'
                    : 'text',
              source: col.key === 'prStatus' ? statusOptions : undefined,
              dateFormat: col.dataType === 'date' ? 'YYYY-MM-DD' : undefined,
              numericFormat: col.dataType === 'numeric'
                ? (quantityLikeKeys.has(col.key) ? { pattern: '0' } : { pattern: '0.00' })
                : undefined,
              // 具体单元格是否可编辑由 cells/beforeChange 控制（支持前置条件+列归属）
              readOnly: col.key === 'totalPrice' || col.key === 'reportedAt',
              width: col.width || 100
            })),
            { type: 'text', readOnly: true, width: 0 }
          ]}
          cells={(row, col) => {
            const props: any = {};
            if (col === idColumnIndex) {
              props.readOnly = true;
              return props;
            }
            if (col < 0 || col >= activeTemplate.length) return props;
            const tCol = activeTemplate[col];
            if (tCol.key === 'totalPrice' || tCol.key === 'reportedAt') {
              props.readOnly = true;
              return props;
            }
            // CS Order：默认只读（仅管理员或采购相关人员可改；一般由项目 csOrder 自动继承）
            if (tCol.key === 'csOrder' && !isAdmin) {
              // 采购人员：由“列分配”或 group 权限控制；其它人一律只读
              // canEditColumn 已做 assignee/group 判断，这里额外兜底：非采购组也禁止
              const isPurchasing = canEditByGroup('purchasing');
              if (!isPurchasing) {
                props.readOnly = true;
                return props;
              }
            }
            // 管理员：可强制编辑任意列（除了 totalPrice / ID）
            if (isAdmin) {
              props.readOnly = false;
              return props;
            }

            // 先判断列归属/角色是否允许编辑
            if (!canEditColumn(tCol)) {
              props.readOnly = true;
              return props;
            }

            // 需求前置条件：除需求字段本身外，其它字段都需要需求字段齐全
            if (!prereqKeySet.has(tCol.key)) {
              const missing = getMissingPrereqsForRow(row);
              if (missing.length > 0) {
                props.readOnly = true;
                return props;
              }
            }
            props.readOnly = false;
            return props;
          }}
          hiddenColumns={{ columns: hiddenColumnIndexes, indicators: false, copyPasteEnabled: false }}
          stretchH="none"
          // 不限制高度：随容器自适应（容器本身可滚动）
          height="auto"
          manualColumnResize={true}
          rowHeights={7}
          rowHeaders={true}
          fixedColumnsLeft={0}
          contextMenu={hasSelectedProject}
          copyPaste={true}
          outsideClickDeselects={false}
          autoColumnSize={false}
          dropdownMenu={true}
          filters={true}
          columnSorting={true}
          afterColumnResize={(currentColumn, newSize) => {
            const hotInstance = hotRef.current?.hotInstance;
            const physicalColumn = hotInstance?.toPhysicalColumn
              ? hotInstance.toPhysicalColumn(currentColumn)
              : currentColumn;
            if (physicalColumn < 0 || physicalColumn >= activeTemplate.length) return;
            const targetCol = activeTemplate[physicalColumn];
            if (!targetCol) return;
            const nextTemplate = fullTemplate.map(col => col.id === targetCol.id ? { ...col, width: newSize } : col);
            setTemplateAndSave(nextTemplate);
          }}
          beforeChange={(changes, source) => {
            if ((source as any) === 'autoCalculate' || (source as any) === 'autoStatus') return;
            if (!changes) return;
            const isBigPaste = (source === 'CopyPaste.paste' || source === 'Autofill.fill') && changes.length > 200;
            const showHint = (msg: string) => {
              // 避免连续弹窗刷屏
              if (isBigPaste) return; // 大批量粘贴时不弹窗，避免浏览器卡死
              const w = window as any;
              const now = Date.now();
              if (w.__tms_last_block_hint && now - w.__tms_last_block_hint < 1200) return;
              w.__tms_last_block_hint = now;
              alert(msg);
            };
            let blockedByAssignee = false;
            let blockedByPrereq = false;
            changes.forEach(change => {
              const colIndex = change[1] as number;
              if (colIndex < 0 || colIndex >= activeTemplate.length) return;
              const editable = editableColumnMap[colIndex];
              const col = activeTemplate[colIndex];
              if (!editable || col.key === 'totalPrice') {
                change[3] = change[2];
                blockedByAssignee = true;
                if (source === 'edit' || source === 'CopyPaste.paste' || source === 'Autofill.fill') {
                  showHint(language === 'zh'
                    ? `该列不可编辑：${getColumnLabel(col)}（未分配给当前用户）`
                    : `Column is read-only: ${getColumnLabel(col)}`);
                }
                return;
              }

              // 前置条件检查：后续列必须先填需求字段
              if (!isAdmin && !prereqKeySet.has(col.key)) {
                const rowIndex = change[0] as number;
                const missing = getMissingPrereqsForRow(rowIndex);
                if (missing.length > 0) {
                  change[3] = change[2];
                  blockedByPrereq = true;
                  if (source === 'edit' || source === 'CopyPaste.paste' || source === 'Autofill.fill') {
                    const missingLabels = missing.map(getKeyLabelByKey).join('、');
                    showHint(language === 'zh'
                      ? `请先填写需求字段：${missingLabels}`
                      : `Please fill required fields first: ${missing.map(getKeyLabelByKey).join(', ')}`);
                  }
                }
              }
            });

            // 大批量粘贴时，用一次性提示代替弹窗刷屏
            if (isBigPaste && (blockedByAssignee || blockedByPrereq)) {
              setTimeout(() => {
                const reasons: string[] = [];
                if (blockedByAssignee) reasons.push('含不可编辑列（未分配给当前用户）');
                if (blockedByPrereq) reasons.push('部分列因“需求字段未填”被拦截');
                alert(`已拦截部分粘贴内容：${reasons.join('；')}`);
              }, 0);
            }
          }}
          afterGetColHeader={(col, TH) => {
            if (!TH || col < 0 || col >= activeTemplate.length) return;
            const current = activeTemplate[col];
            if (current.group === 'engineering') {
              TH.style.backgroundColor = '#6ee7b7';
              TH.style.color = '#064e3b';
            } else if (current.group === 'purchasing') {
              TH.style.backgroundColor = '#93c5fd';
              TH.style.color = '#1e3a8a';
            } else if (current.group === 'warehouse') {
              TH.style.backgroundColor = '#f472b6';
              TH.style.color = '#831843';
            } else if (current.group === 'quality') {
              TH.style.backgroundColor = '#d8b4fe';
              TH.style.color = '#4c1d95';
            }
            TH.style.fontWeight = 'bold';
          }}
          afterChange={(changes, source) => {
            // 性能策略：
            // - 大批量粘贴（1000+行）时，实时计算 totalPrice / prStatus 会导致浏览器“无响应”
            // - 这些派生字段已在“保存”流程（handleSave）里统一计算
            // 因此这里不再做实时计算，避免粘贴高峰卡死
            if (!changes) return;
            if (source === 'loadData') return;
            if ((source as any) === 'autoCalculate' || (source as any) === 'autoStatus') return;
            return;
          }}
          licenseKey="non-commercial-and-evaluation"
        />
      </div>
    </div>
  );
}

function BiddingManagement() {
  const { state, setState, language } = useContext(AppContext)!;
  const t = translations[language];

  return (
    <div className="space-y-6">
      <div className="bg-indigo-900 rounded-[40px] p-10 text-white relative overflow-hidden">
        <div className="relative z-10 max-w-2xl">
          <h3 className="text-3xl font-black tracking-tighter mb-4">{t.bidding}</h3>
          <p className="text-indigo-200 font-bold mb-8">
            Manage vendor bidding process, compare prices and select the best offer for your project materials.
          </p>
          <div className="flex items-center gap-4">
             <div className="bg-white/10 backdrop-blur-md rounded-2xl p-4 flex-1">
               <p className="text-[10px] font-black uppercase tracking-widest text-indigo-300 mb-1">Active Bids</p>
               <p className="text-2xl font-black">{state.biddingTasks.filter(t => t.status === 'Bidding').length}</p>
             </div>
             <div className="bg-white/10 backdrop-blur-md rounded-2xl p-4 flex-1">
               <p className="text-[10px] font-black uppercase tracking-widest text-indigo-300 mb-1">Completed</p>
               <p className="text-2xl font-black">{state.biddingTasks.filter(t => t.status === 'Complete').length}</p>
             </div>
          </div>
        </div>
        <Gavel className="absolute right-10 top-1/2 -translate-y-1/2 text-white/5" size={240} />
      </div>

      <div className="grid grid-cols-1 gap-6">
        {state.biddingTasks.length === 0 ? (
          <div className="p-20 text-center bg-white rounded-[40px] border border-dashed border-slate-200">
            <div className="w-20 h-20 bg-slate-50 rounded-full flex items-center justify-center mx-auto mb-6">
              <Gavel size={40} className="text-slate-200" />
            </div>
            <h4 className="text-xl font-black text-slate-900 mb-2">No Bidding Tasks Found</h4>
            <p className="text-slate-400 font-bold max-w-sm mx-auto">Start a new bidding process by grouping pending requirements and inviting suppliers.</p>
          </div>
        ) : (
          state.biddingTasks.map(task => (
             <div key={task.id} className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
                <div className="flex items-center justify-between mb-6">
                   <div>
                     <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Task ID: {task.id}</span>
                     <h4 className="text-lg font-black text-slate-900">Bidding for {task.requirementIds.length} items</h4>
                   </div>
                   <span className={cn(
                     "px-3 py-1 rounded-full text-xs font-black uppercase",
                     task.status === 'Complete' ? "bg-emerald-50 text-emerald-600" : "bg-blue-50 text-blue-600"
                   )}>
                     {task.status}
                   </span>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 mb-6">
                  {task.suppliers.map(s => (
                    <div key={s.id} className={cn(
                      "p-4 rounded-2xl border transition-all",
                      s.isSelected ? "border-blue-600 bg-blue-50" : "border-slate-100 bg-slate-50"
                    )}>
                      <div className="flex items-start justify-between mb-2">
                        <span className="font-black text-slate-900">{s.name}</span>
                        {s.isSelected && <CheckCircle2 className="text-blue-600" size={16} />}
                      </div>
                      <p className="text-xl font-black text-slate-900">¥ {s.price.toLocaleString()}</p>
                      {s.comment && <p className="text-[10px] text-slate-400 font-bold mt-2 italic">"{s.comment}"</p>}
                    </div>
                  ))}
                </div>
                <div className="flex items-center justify-between pt-6 border-t border-slate-50">
                   <div className="flex items-center gap-2 text-slate-400">
                     <Calendar size={14} />
                     <span className="text-[10px] font-black uppercase tracking-widest">Created: {format(parseISO(task.createdAt), 'yyyy-MM-dd')}</span>
                   </div>
                   <div className="flex gap-2">
                     <button className="btn-secondary py-2 h-10 px-4 text-xs font-black">{t.details}</button>
                     {task.status !== 'Complete' && <button className="btn-primary py-2 h-10 px-4 text-xs font-black">{t.createPR}</button>}
                   </div>
                </div>
             </div>
          ))
        )}
      </div>
    </div>
  );
}

function PurchasingManagement() {
  const { state, language } = useContext(AppContext)!;
  const t = translations[language];

  return (
    <div className="space-y-8">
      {/* PR Section */}
      <section className="space-y-4">
        <h3 className="text-xl font-black text-slate-900 flex items-center gap-2">
          <ClipboardCheck size={24} className="text-blue-600" />
          {t.purchaseRequest}
        </h3>
        <div className="grid grid-cols-1 gap-4">
          {state.purchaseRequests.length === 0 ? (
            <div className="p-12 text-center bg-slate-50/50 rounded-3xl border border-dashed border-slate-200 text-slate-400 font-bold italic text-sm">
              No Purchase Requests yet.
            </div>
          ) : (
            state.purchaseRequests.map(pr => (
              <div key={pr.id} className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm flex items-center justify-between">
                <div>
                   <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">PR ID: {pr.id}</span>
                   <h4 className="text-lg font-black text-slate-900">¥ {pr.totalPrice.toLocaleString()}</h4>
                   <p className="text-xs font-bold text-slate-400">{pr.requirementIds.length} materials from {pr.supplierId}</p>
                </div>
                <div className="flex items-center gap-8">
                  <div className="flex items-center gap-3">
                    {pr.approvalFlow.map((step, i) => (
                      <div key={i} className="flex items-center gap-1">
                        <div className={cn(
                          "w-8 h-8 rounded-lg flex items-center justify-center text-[10px] font-black",
                          step.status === 'Approved' ? "bg-emerald-100 text-emerald-600" :
                          step.status === 'Rejected' ? "bg-rose-100 text-rose-600" : "bg-slate-100 text-slate-400"
                        )}>
                          {step.role === 'pm' ? 'PM' : step.role === 'approver' ? 'LDR' : 'FIN'}
                        </div>
                        {i < pr.approvalFlow.length - 1 && <ChevronRight size={12} className="text-slate-300" />}
                      </div>
                    ))}
                  </div>
                  <button className={cn(
                    "px-4 py-2 rounded-xl text-xs font-black transition-all",
                    pr.status === 'Approved' ? "bg-emerald-500 text-white" : "bg-slate-100 text-slate-400"
                  )}>
                    {pr.status === 'Approved' ? t.createPO : t.pending}
                  </button>
                </div>
              </div>
            ))
          )}
        </div>
      </section>

      {/* PO Section */}
      <section className="space-y-4">
        <h3 className="text-xl font-black text-slate-900 flex items-center gap-2">
          <FileCheck size={24} className="text-blue-600" />
          {t.purchaseOrder}
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {state.purchaseOrders.length === 0 ? (
            <div className="md:col-span-2 p-12 text-center bg-slate-50/50 rounded-3xl border border-dashed border-slate-200 text-slate-400 font-bold italic text-sm">
              No Purchase Orders yet.
            </div>
          ) : (
            state.purchaseOrders.map(po => (
              <div key={po.id} className="bg-white border border-slate-100 rounded-3xl p-6 shadow-sm">
                <div className="flex justify-between items-start mb-4">
                   <div>
                     <p className="text-[10px] font-black text-blue-500 uppercase tracking-widest">{po.poNumber}</p>
                     <p className="text-xs text-slate-400 font-bold">{po.supplierId}</p>
                   </div>
                   <span className="bg-blue-50 text-blue-600 px-3 py-1 rounded-full text-[10px] font-black uppercase">{po.status}</span>
                </div>
                <div className="space-y-4">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-slate-50 rounded-xl flex items-center justify-center">
                      <Truck className="text-slate-400" size={18} />
                    </div>
                    <div className="flex-1">
                      <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                        <div className="h-full bg-blue-600 w-1/2" />
                      </div>
                      <div className="flex justify-between mt-1 text-[9px] font-black text-slate-400 uppercase tracking-tighter">
                        <span>Issued</span>
                        <span>Delivered</span>
                      </div>
                    </div>
                  </div>
                  <div className="flex items-center justify-between text-xs font-bold text-slate-500">
                    <span className="flex items-center gap-1"><Clock size={12} /> Exp: {po.estimatedDelivery || '-'}</span>
                    <button className="text-blue-600 hover:underline">Track</button>
                  </div>
                </div>
              </div>
            ))
          )}
        </div>
      </section>
    </div>
  );
}

function ProcurementAnalytics({ requirements }: { requirements: MaterialRequirement[] }) {
  const { language } = useContext(AppContext)!;
  const t = translations[language];

  const statusData = useMemo(() => {
    const statuses = ['Pending Review', 'Approved', 'Bidding', 'PO Issued', 'Delivered'];
    return statuses.map(s => ({
      name: s,
      value: requirements.filter(r => r.status === s).length
    }));
  }, [requirements]);

  const overdueRequirements = useMemo(() => {
    return requirements.filter(r => r.status !== 'Delivered' && new Date(r.expectedArrival) < new Date());
  }, [requirements]);

  const COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444'];

  return (
    <div className="space-y-8">
      {/* Overview Cards */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
          <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">{t.materialRequirements}</p>
          <p className="text-3xl font-black text-slate-900">{requirements.length}</p>
        </div>
        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm border-l-4 border-l-rose-500">
          <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">{t.overdueAlert}</p>
          <p className="text-3xl font-black text-rose-600">{overdueRequirements.length}</p>
        </div>
        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm border-l-4 border-l-blue-500">
          <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">{t.onTime}</p>
          <p className="text-3xl font-black text-blue-600">{requirements.length - overdueRequirements.length}</p>
        </div>
        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm border-l-4 border-l-emerald-500">
          <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">{language === 'zh' ? '入库达成' : 'Success Rate'}</p>
          <p className="text-3xl font-black text-emerald-600">{requirements.length > 0 ? Math.round((requirements.filter(r => r.status === 'Delivered').length / requirements.length) * 100) : 0}%</p>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
      <div className="bg-white p-8 rounded-[40px] border border-slate-100 shadow-sm">
        <h3 className="text-lg font-black text-slate-900 mb-6 flex items-center gap-2">
          <LucidePieChart size={20} className="text-blue-600" />
          {language === 'zh' ? '当前进度占比' : 'Status Distribution'}
        </h3>
        <div className="h-[300px] w-full">
          <ResponsiveContainer width="100%" height="100%">
            <RePieChart>
              <Pie
                data={statusData}
                cx="50%"
                cy="50%"
                innerRadius={60}
                outerRadius={100}
                paddingAngle={5}
                dataKey="value"
              >
                {statusData.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                ))}
              </Pie>
              <RechartsTooltip />
              <Legend />
            </RePieChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="bg-white p-8 rounded-[40px] border border-slate-100 shadow-sm">
         <h3 className="text-lg font-black text-slate-900 mb-6 flex items-center gap-2">
          <TrendingUp size={20} className="text-blue-600" />
          {language === 'zh' ? '紧急程度统计' : 'Urgency Breakdown'}
        </h3>
        <div className="space-y-6">
          {['Low', 'Medium', 'High', 'Urgent'].map((u, i) => {
            const count = requirements.filter(r => r.urgency === u).length;
            const percentage = requirements.length > 0 ? (count / requirements.length) * 100 : 0;
            return (
              <div key={u} className="space-y-2">
                <div className="flex justify-between items-center text-xs font-black">
                  <span className={cn(
                    "uppercase tracking-widest",
                    u === 'Urgent' ? "text-rose-600" : "text-slate-400"
                  )}>{u}</span>
                  <span className="text-slate-900">{count}</span>
                </div>
                <div className="h-2 bg-slate-50 rounded-full overflow-hidden">
                  <motion.div 
                    initial={{ width: 0 }}
                    animate={{ width: `${percentage}%` }}
                    className={cn(
                      "h-full rounded-full",
                      u === 'Urgent' ? "bg-rose-500 shadow-lg shadow-rose-200" :
                      u === 'High' ? "bg-orange-500" : "bg-blue-500"
                    )}
                  />
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  </div>
);
}


export function EntryView() {
  const context = useContext(AppContext);
  if (!context) return null;
  const { state, setState, language, handleSaveToDatabase, addLog } = context;

  const [viewingEmployeeId, setViewingEmployeeId] = useState(state.currentUser?.id || '');
  const [formData, setFormData] = useState({
    employeeId: state.currentUser?.id || '',
    projectId: '',
    category2Id: '',
    category3Id: '',
    date: format(new Date(), 'yyyy-MM-dd'),
    hours: '',
    comment: '',
    allocationId: '' // 新增：当前项目分工
  });
  const [adjData, setAdjData] = useState({
    date: format(new Date(), 'yyyy-MM-dd'),
    overtime: '',
    leave: ''
  });

  // 联动日期
  useEffect(() => {
    setAdjData(prev => ({ ...prev, date: formData.date }));
  }, [formData.date]);

  const [calendarRange, setCalendarRange] = useState({
    start: format(startOfMonth(new Date()), 'yyyy-MM-dd'),
    end: format(endOfMonth(new Date()), 'yyyy-MM-dd')
  });
  const formRef = useRef<HTMLFormElement>(null);

  const [editingCommentId, setEditingCommentId] = useState<string | null>(null);
  const [tempComment, setTempComment] = useState('');

  const [selectedDayDetails, setSelectedDayDetails] = useState<string | null>(null);
  const [editingEntryId, setEditingEntryId] = useState<string | null>(null);

  // 一人多角色映射管理状态
  const [newMapping, setNewMapping] = useState<{
    userId: string,
    category2Ids: string[],
    category3Ids: string[],
    allocationId: string
  }>({ userId: '', category2Ids: [], category3Ids: [], allocationId: '' });
  const [editingMappingId, setEditingMappingId] = useState<string | null>(null);

  const startEditMapping = (m: any) => {
    setEditingMappingId(m.id);
    setNewMapping({
      userId: m.userId || '',
      category2Ids: m.category2Ids || [],
      category3Ids: m.category3Ids || [],
      allocationId: m.allocationId || ''
    });
  };

  const addMapping = async () => {
    if (newMapping.category2Ids.length === 0 || newMapping.category3Ids.length === 0 || !newMapping.allocationId) return;
    
    // Validation: Check for conflicts
    const conflicts: string[] = [];
    const currentMappings = state.userMappings || [];
    
    newMapping.category2Ids.forEach(c2Id => {
      newMapping.category3Ids.forEach(c3Id => {
        const existing = currentMappings.find(m => 
          m.id !== editingMappingId && 
          ((m.userId === newMapping.userId) || 
           ((!m.userId || m.userId === '') && (!newMapping.userId || newMapping.userId === ''))) &&
          m.category2Ids?.includes(c2Id) &&
          m.category3Ids?.includes(c3Id)
        );
        
        if (existing && existing.allocationId !== newMapping.allocationId) {
          const c2 = state.categories2.find(c => c.id === c2Id);
          const c3 = state.categories3.find(c => c.id === c3Id);
          conflicts.push(`${getCatName(c2, language)} @ ${getCatName(c3, language)}`);
        }
      });
    });

    if (conflicts.length > 0) {
      alert((language === 'zh' ? '冲突：以下分类组合已分配给其他项目分工：\n' : 'Conflict: The following category combinations are already assigned to other allocations:\n') + conflicts.join('\n'));
      return;
    }

    let newState;
    if (editingMappingId) {
      newState = {
        ...state,
        userMappings: state.userMappings.map(m => m.id === editingMappingId ? { ...m, ...newMapping } : m)
      };
    } else {
      const id = 'um-' + Date.now();
      const mapping = { ...newMapping, id };
      newState = { ...state, userMappings: [...state.userMappings, mapping] };
    }
    
    setState(newState);
    handleSaveToDatabase(newState);
    setNewMapping({ userId: '', category2Ids: [], category3Ids: [], allocationId: '' });
    setEditingMappingId(null);
  };

  const isSuperAdmin = state.currentUser?.role === 'admin';
  const visibleEmployees = getVisibleEmployees(state.currentUser, state.employees);
  const editableEmployees = getEditableEmployees(state.currentUser, state.employees);
  const editableEmployeeIdSet = useMemo(() => new Set(editableEmployees.map(e => e.id)), [editableEmployees]);
  const canEditViewingEmployee = useMemo(() => editableEmployeeIdSet.has(viewingEmployeeId), [editableEmployeeIdSet, viewingEmployeeId]);

  // Sync viewingEmployeeId when currentUser changes
  useEffect(() => {
    if (state.currentUser) {
      setViewingEmployeeId(state.currentUser.id);
    }
  }, [state.currentUser]);

  // Sync formData.employeeId when viewingEmployeeId changes (Fix for Ma Yijun filling for others)
  useEffect(() => {
    setFormData(prev => ({ ...prev, employeeId: viewingEmployeeId }));
  }, [viewingEmployeeId]);

  // Auto-set allocationId based on user mappings
  useEffect(() => {
    if (formData.category2Id && formData.category3Id) {
      // 1. Try user-specific mapping
      let mapping = state.userMappings.find(m => 
        m.userId === formData.employeeId && 
        m.category2Ids?.includes(formData.category2Id) && 
        m.category3Ids?.includes(formData.category3Id)
      );
      
      // 2. Try global mapping if no user-specific one
      if (!mapping) {
        mapping = state.userMappings.find(m => 
          (!m.userId || m.userId === '') && 
          m.category2Ids?.includes(formData.category2Id) && 
          m.category3Ids?.includes(formData.category3Id)
        );
      }

      if (mapping) {
        setFormData(prev => ({ ...prev, allocationId: mapping.allocationId }));
      } else {
        // Only clear if it was previously set by a mapping or if we want strict enforcement
        // For now, let's clear it to ensure the user knows no mapping was found
        setFormData(prev => ({ ...prev, allocationId: '' }));
      }
    } else {
      // If categories are cleared, clear allocation too
      setFormData(prev => ({ ...prev, allocationId: '' }));
    }
  }, [formData.category2Id, formData.category3Id, formData.employeeId]); // Removed state.userMappings from deps to avoid re-triggering on every state change unless categories change

  const hasMapping = useMemo(() => {
    if (!formData.category2Id || !formData.category3Id) return false;
    return state.userMappings.some(m => 
      (m.userId === formData.employeeId || !m.userId || m.userId === '') &&
      m.category2Ids?.includes(formData.category2Id) &&
      m.category3Ids?.includes(formData.category3Id)
    );
  }, [formData.category2Id, formData.category3Id, formData.employeeId, state.userMappings]);

  const myEntries = state.entries.filter(e => e.employeeId === viewingEmployeeId);
  const myAdjustments = state.dailyAdjustments.filter(a => a.employeeId === viewingEmployeeId);
  const selectedEmployee = state.employees.find(e => e.id === viewingEmployeeId);

  const handleUpdateComment = async (id: string) => {
    if (!canEditViewingEmployee) {
      alert(language === 'zh' ? '你只有查看权限，不能修改他人工时。' : 'Read-only: you cannot edit other people\'s entries.');
      return;
    }
    const entry = state.entries.find(e => e.id === id);
    if (!entry) return;
    const newState = {
      ...state,
      entries: state.entries.map(e => e.id === id ? { ...e, comment: tempComment } : e)
    };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '修改备注' : 'Update Comment', `Entry: ${id}`);
    setEditingCommentId(null);
  };

  const handleDeleteEntry = async (id: string) => {
    if (!canEditViewingEmployee) {
      alert(language === 'zh' ? '你只有查看权限，不能删除他人工时。' : 'Read-only: you cannot delete other people\'s entries.');
      return;
    }
    const entryToDelete = state.entries.find(e => e.id === id);
    if (!entryToDelete) return;

    const newState = {
      ...state,
      entries: state.entries.filter(e => e.id !== id),
      deletedEntries: [entryToDelete, ...state.deletedEntries].slice(0, 50)
    };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '删除工时记录' : 'Remove Entry', `${entryToDelete.date} - ${entryToDelete.hours}h`);
  };

  const handleDeleteAdjustment = async (id: string) => {
    if (!canEditViewingEmployee) {
      alert(language === 'zh' ? '你只有查看权限，不能修改他人的加班/请假调整。' : 'Read-only: you cannot edit other people\'s adjustments.');
      return;
    }
    const adjToDelete = state.dailyAdjustments.find(a => a.id === id);
    if (!adjToDelete) return;

    const newState = {
      ...state,
      dailyAdjustments: state.dailyAdjustments.filter(a => a.id !== id),
      deletedAdjustments: [adjToDelete, ...state.deletedAdjustments].slice(0, 50)
    };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '删除调整' : 'Remove Adjustment', `${adjToDelete.date}`);
  };

  const handleUndoDelete = () => {
    if (state.deletedEntries.length === 0 && state.deletedAdjustments.length === 0) return;
    
    let newState = { ...state };
    
    if (state.deletedEntries.length > 0) {
      const [entryToRestore, ...remainingEntries] = state.deletedEntries;
      newState = {
        ...newState,
        entries: [entryToRestore, ...newState.entries],
        deletedEntries: remainingEntries
      };
      addLog(language === 'zh' ? '撤销删除' : 'Undo Delete', `Entry: ${entryToRestore.date}`);
    } else if (state.deletedAdjustments.length > 0) {
      const [adjToRestore, ...remainingAdjs] = state.deletedAdjustments;
      newState = {
        ...newState,
        dailyAdjustments: [adjToRestore, ...newState.dailyAdjustments],
        deletedAdjustments: remainingAdjs
      };
      addLog(language === 'zh' ? '撤销删除' : 'Undo Delete', `Adjustment: ${adjToRestore.date}`);
    }
    
    setState(newState);
    handleSaveToDatabase(newState);
  };

  const calculateEfficiency = (periodEntries: TimesheetEntry[], adjustments: DailyAdjustment[]) => {
    const dailyData: { [date: string]: { actual: number, leave: number, general: number, project: number } } = {};
    
    periodEntries.forEach(e => {
      if (!dailyData[e.date]) dailyData[e.date] = { actual: 0, leave: 0, general: 0, project: 0 };
      if (e.projectId === 'p-leave') {
        dailyData[e.date].leave += e.hours;
      } else if (e.projectId === 'p-general') {
        dailyData[e.date].general += e.hours;
        dailyData[e.date].actual += e.hours;
      } else {
        dailyData[e.date].project += e.hours;
        dailyData[e.date].actual += e.hours;
      }
    });

    adjustments.forEach(adj => {
      if (!dailyData[adj.date]) dailyData[adj.date] = { actual: 0, leave: 0, general: 0, project: 0 };
      dailyData[adj.date].leave += adj.leave;
      dailyData[adj.date].actual += adj.overtime;
    });

    let totalActual = 0;
    let totalExpected = 0;
    let totalProject = 0;
    let totalGeneral = 0;

    Object.keys(dailyData).forEach(date => {
      const day = dailyData[date];
      // User formula: Daily Total = General + Project
      const dailyTotal = day.general + day.project;
      // Available Time = Daily Total - Leave + Overtime
      // day.actual includes overtime? 
      // Let's look at how day.actual was calculated:
      // adjustments.forEach(adj => { ... dailyData[adj.date].actual += adj.overtime; })
      // So Overtime = day.actual - (day.general + day.project)
      const overtime = day.actual - dailyTotal;
      const available = dailyTotal - day.leave + overtime;
      
      totalExpected += available > 0 ? available : 8; 
      totalProject += day.project;
      totalGeneral += day.general;
    });

    const efficiency = totalExpected > 0 ? ((totalGeneral + totalProject) / totalExpected) : 0;
    const projectRatio = (totalProject + totalGeneral) > 0 ? (totalProject / (totalProject + totalGeneral)) : 0;

    return { efficiency, projectRatio };
  };

  const getAllocOptions = () => {
    return state.allocations;
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!canEditViewingEmployee) {
      alert(language === 'zh' ? '你只有查看权限，不能代填/修改他人工时。' : 'Read-only: you cannot create/update other people\'s entries.');
      return;
    }
    if (!formData.employeeId || !formData.projectId || !formData.category2Id || !formData.category3Id || !formData.hours) {
      alert(language === 'zh' ? '请填写完整信息（包括分类2和分类3）' : 'Please fill all info (including Category 2 & 3)');
      return;
    }

    const hours = parseFloat(formData.hours);
    const date = formData.date;
    
    // Auto-set allocation based on mapping
    let allocId = formData.allocationId;
    const mapping = state.userMappings.find(m => 
      (m.userId === formData.employeeId || !m.userId || m.userId === '') && 
      m.category2Ids?.includes(formData.category2Id) && 
      m.category3Ids?.includes(formData.category3Id)
    );
    if (mapping) {
      allocId = mapping.allocationId;
    }

    const selectedProj = state.projects.find(p => p.id === formData.projectId);
    const hasBudget = selectedProj?.budgets && allocId && selectedProj.budgets[allocId] > 0;
    
    if (!hasBudget && formData.projectId !== 'p-general' && formData.projectId !== 'p-leave') {
      const confirmMsg = language === 'zh' 
        ? `提示：该项目没有预设您当前分工(${state.allocations.find(a => a.id === allocId)?.nameZh || '-'})的工时，是否继续填报？` 
        : `Note: This project has no pre-allocated hours for your current allocation (${state.allocations.find(a => a.id === allocId)?.nameEn || '-'}). Continue?`;
      if (!window.confirm(confirmMsg)) return;
    }

    const dayAdj = myAdjustments.find(a => a.date === date);
    const overtime = dayAdj?.overtime || 0;
    const leave = dayAdj?.leave || 0;
    const maxHours = 8 + overtime - leave;
    
    const existingHours = myEntries.filter(e => e.date === date && e.id !== editingEntryId).reduce((acc, e) => acc + e.hours, 0);
    
    if (existingHours + hours > maxHours + 0.001) {
      alert(language === 'zh' 
        ? `超出当日工时上限！当日上限为 ${maxHours}h (8 + 加班 ${overtime} - 请假 ${leave})，已填报 ${existingHours}h。` 
        : `Exceeds daily limit! Limit is ${maxHours}h (8 + OT ${overtime} - Leave ${leave}), already logged ${existingHours}h.`);
      return;
    }

    const entryId = editingEntryId || 'e-' + Date.now();
    const entryData = {
      id: entryId,
      employeeId: formData.employeeId,
      projectId: formData.projectId,
      category2Id: formData.category2Id,
      category3Id: formData.category3Id,
      allocationId: allocId,
      date: formData.date,
      hours: hours,
      comment: formData.comment
    };

    const newState = {
      ...state,
      entries: editingEntryId 
        ? state.entries.map(e => e.id === editingEntryId ? entryData : e)
        : [entryData, ...state.entries]
    };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? (editingEntryId ? '修改工时' : '新增工时') : (editingEntryId ? 'Edit Entry' : 'Add Entry'), `${formData.date}: ${hours}h`);
    setEditingEntryId(null);
    setFormData({ ...formData, hours: '', comment: '', allocationId: '' });
  };

  const handleEditEntry = (entry: TimesheetEntry) => {
    if (!canEditViewingEmployee) {
      alert(language === 'zh' ? '你只有查看权限，不能编辑他人工时。' : 'Read-only: you cannot edit other people\'s entries.');
      return;
    }
    setFormData({
      employeeId: entry.employeeId,
      projectId: entry.projectId,
      category2Id: entry.category2Id || '',
      category3Id: entry.category3Id || '',
      date: entry.date,
      hours: entry.hours.toString(),
      comment: entry.comment || '',
      allocationId: entry.allocationId || ''
    });
    setEditingEntryId(entry.id);
    formRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  };

  const handleAdjSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!canEditViewingEmployee) {
      alert(language === 'zh' ? '你只有查看权限，不能代填/修改他人的加班/请假调整。' : 'Read-only: you cannot edit other people\'s adjustments.');
      return;
    }
    const overtimeInput = parseFloat(adjData.overtime) || 0;
    const leaveInput = parseFloat(adjData.leave) || 0;
    
    const existing = state.dailyAdjustments.find(a => a.date === adjData.date && a.employeeId === viewingEmployeeId);
    const adjId = existing?.id || 'adj-' + Date.now();
    
    // User request: adjustments are cumulative
    const newOvertime = (existing?.overtime || 0) + overtimeInput;
    const newLeave = (existing?.leave || 0) + leaveInput;

    const newAdj = {
      id: adjId,
      employeeId: viewingEmployeeId,
      date: adjData.date,
      overtime: newOvertime,
      leave: newLeave
    };

    const newState = {
      ...state,
      dailyAdjustments: existing 
        ? state.dailyAdjustments.map(a => a.id === adjId ? newAdj : a)
        : [newAdj, ...state.dailyAdjustments]
    };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '保存调整' : 'Save Adjustment', `${adjData.date}: OT +${overtimeInput} (Total: ${newOvertime}), Leave +${leaveInput} (Total: ${newLeave}) (For: ${viewingEmployeeId})`);
    alert(language === 'zh' ? '调整已保存' : 'Adjustments saved');
    setAdjData({ ...adjData, overtime: '', leave: '' });
  };

  const groupedEntries = myEntries.reduce((acc, e) => {
    if (!acc[e.date]) acc[e.date] = [];
    acc[e.date].push(e);
    return acc;
  }, {} as { [date: string]: TimesheetEntry[] });

  const sortedDates = Object.keys(groupedEntries).sort((a, b) => b.localeCompare(a));

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="w-full space-y-10 pb-20">
        <WeeklyGrid language={language} state={state} setState={setState} handleSaveToDatabase={handleSaveToDatabase} initialRange={calendarRange} />

        {/* Detailed History & Mapping with Fixed Width - Left Aligned */}
        <div className="max-w-6xl w-full space-y-12">
          {/* Detailed History */}
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
            <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-3">
              <div className="text-sm font-bold text-slate-900">
                {language === 'zh' ? '历史填报明细' : 'Detailed History'}
              </div>
              {(state.deletedEntries.length > 0 || state.deletedAdjustments.length > 0) && (
                <button
                  onClick={handleUndoDelete}
                  className="btn-secondary h-8 px-3 text-xs"
                >
                  {language === 'zh'
                    ? `撤销删除 (${state.deletedEntries.length + state.deletedAdjustments.length})`
                    : `Undo Delete (${state.deletedEntries.length + state.deletedAdjustments.length})`}
                </button>
              )}
            </div>

            <div className="max-h-[600px] overflow-y-auto custom-scrollbar">
              <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-slate-50 z-10">
                  <tr className="border-b border-slate-100">
                      <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '日期' : 'Date'}</th>
                      <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '项目' : 'Project'}</th>
                      <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Cat 2</th>
                      <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Cat 3</th>
                      <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '工时' : 'Hours'}</th>
                      <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '备注' : 'Comment'}</th>
                      <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">{language === 'zh' ? '操作' : 'Action'}</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                    {(() => {
                      const groupedEntries = myEntries.reduce((acc, e) => {
                        if (!acc[e.date]) acc[e.date] = [];
                        acc[e.date].push(e);
                        return acc;
                      }, {} as { [date: string]: TimesheetEntry[] });

                      // Add dates from adjustments that might not have entries
                      myAdjustments.forEach(adj => {
                        if (!groupedEntries[adj.date]) {
                          groupedEntries[adj.date] = [];
                        }
                      });

                      const sortedDates = Object.keys(groupedEntries).sort((a, b) => b.localeCompare(a));

                      return sortedDates.map(date => {
                        const dayAdj = myAdjustments.find(a => a.date === date);
                        const dayEntries = groupedEntries[date];
                        
                        // User request: Daily Total = Actual hours (General + Project), excluding Leave entries
                        const dailyTotal = dayEntries.filter(e => e.projectId !== 'p-leave').reduce((acc, e) => acc + e.hours, 0);
                        
                        // Available Time = 8 (Standard) - Leave Adjustment + Overtime Adjustment
                        const totalOvertime = dayAdj?.overtime || 0;
                        const totalLeave = dayAdj?.leave || 0;
                        const availableTime = 8 - totalLeave + totalOvertime;

                        return (
                          <React.Fragment key={date}>
                            <tr className="bg-slate-50/50">
                              <td colSpan={7} className="px-6 py-3">
                                <div className="flex items-center justify-between">
                                  <div className="flex items-center gap-6">
                                    <div className="flex items-center gap-2">
                                      <Calendar size={14} className="text-blue-600" />
                                      <span className="text-xs font-black text-slate-900">{date}</span>
                                    </div>
                                    <div className="flex items-center gap-4">
                                      <div className="flex flex-col">
                                        <span className="text-[9px] font-black text-slate-400 uppercase tracking-tighter">{language === 'zh' ? '当日总计' : 'Daily Total'}</span>
                                        <span className="text-xs font-black text-slate-900">{dailyTotal.toFixed(1)}h</span>
                                      </div>
                                      <div className="flex flex-col">
                                        <span className="text-[9px] font-black text-slate-400 uppercase tracking-tighter">{language === 'zh' ? '可用工时' : 'Available Time'}</span>
                                        <span className="text-xs font-black text-blue-600">{availableTime.toFixed(1)}h</span>
                                      </div>
                                    </div>
                                    {dayAdj && (
                                      <div className="flex items-center gap-2 border-l border-slate-200 pl-4">
                                        {dayAdj.overtime > 0 && (
                                          <div className="flex flex-col">
                                            <span className="text-[9px] font-black text-green-600 uppercase tracking-tighter">{language === 'zh' ? '加班' : 'OT'}</span>
                                            <span className="text-xs font-black text-green-600">+{dayAdj.overtime}h</span>
                                          </div>
                                        )}
                                        {dayAdj.leave > 0 && (
                                          <div className="flex flex-col">
                                            <span className="text-[9px] font-black text-rose-600 uppercase tracking-tighter">{language === 'zh' ? '请假' : 'LV'}</span>
                                            <span className="text-xs font-black text-rose-600">-{dayAdj.leave}h</span>
                                          </div>
                                        )}
                                        <button 
                                          onClick={() => handleDeleteAdjustment(dayAdj.id)}
                                          className="p-1.5 hover:bg-rose-50 text-slate-300 hover:text-rose-600 rounded-lg transition-all ml-2"
                                          title={language === 'zh' ? '删除调整' : 'Delete Adjustment'}
                                        >
                                          <X size={14} />
                                        </button>
                                      </div>
                                    )}
                                  </div>
                                </div>
                              </td>
                            </tr>
                            {dayEntries.map(entry => (
                              <tr key={entry.id} className="hover:bg-slate-50/30 transition-colors group">
                                <td className="px-6 py-4 text-xs font-bold text-slate-400">{entry.date}</td>
                                <td className="px-6 py-4 font-bold text-slate-900 text-sm">
                                  <div className="flex flex-col gap-1">
                                    <div className="flex items-center gap-2">
                                      <div className="w-1.5 h-1.5 rounded-full bg-blue-500"></div>
                                      <span className="text-xs font-black text-slate-900">{state.projects.find(p => p.id === entry.projectId)?.name}</span>
                                    </div>
                                    <div className="flex items-center gap-2">
                                      <Briefcase size={10} className="text-slate-400" />
                                      <span className="text-[10px] text-slate-500 font-bold">
                                        {(() => {
                                          const alloc = state.allocations.find(a => a.id === entry.allocationId);
                                          return language === 'zh' ? (alloc?.nameZh || '-') : (alloc?.nameEn || alloc?.nameZh || '-');
                                        })()}
                                      </span>
                                    </div>
                                  </div>
                                </td>
                                <td className="px-6 py-4">
                                  <span className="text-[10px] bg-blue-50 text-blue-500 px-1.5 py-0.5 rounded font-bold">
                                    {getCatName(state.categories2.find(c => c.id === entry.category2Id), language)}
                                  </span>
                                </td>
                                <td className="px-6 py-4">
                                  <span className="text-[10px] bg-indigo-50 text-indigo-500 px-1.5 py-0.5 rounded font-bold">
                                    {getCatName(state.categories3.find(c => c.id === entry.category3Id), language)}
                                  </span>
                                </td>
                                <td className="px-6 py-4 font-black text-blue-600 text-sm">{entry.hours}h</td>
                                <td className="px-6 py-4">
                                  {editingCommentId === entry.id ? (
                                    <div className="flex items-center gap-2">
                                      <input 
                                        type="text" 
                                        value={tempComment} 
                                        onChange={e => setTempComment(e.target.value)}
                                        className="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1 text-xs font-bold outline-none focus:ring-2 focus:ring-blue-500/20 w-full"
                                      />
                                      <button onClick={() => handleUpdateComment(entry.id)} className="text-blue-600 hover:text-blue-500">
                                        <CheckCircle2 size={16} />
                                      </button>
                                      <button onClick={() => setEditingCommentId(null)} className="text-slate-300 hover:text-slate-400">
                                        <X size={16} />
                                      </button>
                                    </div>
                                  ) : (
                                    <p className="text-xs text-slate-400 italic truncate max-w-[200px]">{entry.comment || '-'}</p>
                                  )}
                                </td>
                                <td className="px-6 py-4 text-right">
                                  <div className="flex items-center justify-end gap-2 transition-opacity">
                                    {editingCommentId !== entry.id && (
                                      <button 
                                        onClick={() => { setEditingCommentId(entry.id); setTempComment(entry.comment || ''); }}
                                        className="p-2 bg-slate-50 text-slate-400 hover:bg-slate-100 hover:text-slate-600 rounded-xl transition-all"
                                        title={language === 'zh' ? '修改备注' : 'Edit Comment'}
                                      >
                                        <MessageSquare size={14} />
                                      </button>
                                    )}
                                    <button 
                                      onClick={() => handleDeleteEntry(entry.id)}
                                      className="p-2 bg-rose-50 text-rose-400 hover:bg-rose-100 hover:text-rose-600 rounded-xl transition-all"
                                      title={language === 'zh' ? '删除' : 'Delete'}
                                    >
                                      <Trash2 size={14} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            ))}
                          </React.Fragment>
                        );
                      });
                    })()}
                    {sortedDates.length === 0 && (
                      <tr>
                        <td colSpan={5} className="px-6 py-12 text-center text-slate-300 font-bold italic">
                          {language === 'zh' ? '暂无记录' : 'No entries found'}
                        </td>
                      </tr>
                    )}
                </tbody>
              </table>
            </div>
          </div>

          {/* Mapping Section at the bottom of Entry page */}
          <div className="mt-12 pt-12 border-t border-slate-100 space-y-6">
            <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
              <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-3">
                <div className="text-sm font-bold text-slate-900">
                  {language === 'zh' ? '一人多角色填报映射' : 'One Person Multi-Role Mapping'}
                </div>
                <span className="text-xs text-slate-500">
                  {language === 'zh'
                    ? '说明：定义“人员+分类”对应的“项目分工”，实现填报时身份的自动切换。'
                    : 'Note: Map "User + Category" to "Project Allocation" for automatic identity switching.'}
                </span>
              </div>

              <div className="p-4 space-y-6">
              {/* Mapping Form */}
              <div className="grid grid-cols-1 md:grid-cols-5 gap-4 items-start bg-slate-50 p-6 rounded-2xl border border-slate-100">
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">{language === 'zh' ? '选择人员 (可选)' : 'Select Employee (Optional)'}</label>
                  <select 
                    value={newMapping.userId || ''}
                    onChange={e => setNewMapping({...newMapping, userId: e.target.value})}
                    className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs font-bold outline-none"
                  >
                    <option value="">{language === 'zh' ? '所有人员 (全局映射)' : 'All Employees (Global)'}</option>
                    {visibleEmployees.map(emp => <option key={emp.id} value={emp.id}>{getEmpName(emp, language)}</option>)}
                  </select>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Category 2 (多选)</label>
                  <div className="bg-white border border-slate-200 rounded-xl p-2 max-h-[120px] overflow-y-auto space-y-1">
                    {state.categories2.map(c => (
                      <label key={c.id} className="flex items-center gap-2 text-[10px] font-bold cursor-pointer hover:bg-slate-50 p-1 rounded">
                        <input 
                          type="checkbox" 
                          checked={newMapping.category2Ids.includes(c.id)}
                          onChange={e => {
                            const ids = e.target.checked 
                              ? [...newMapping.category2Ids, c.id]
                              : newMapping.category2Ids.filter(id => id !== c.id);
                            setNewMapping({ ...newMapping, category2Ids: ids, category3Ids: [] });
                          }}
                        />
                        {getCatName(c, language)}
                      </label>
                    ))}
                  </div>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Category 3 (多选)</label>
                  <div className="bg-white border border-slate-200 rounded-xl p-2 max-h-[120px] overflow-y-auto space-y-1">
                    {state.categories3.filter(c => newMapping.category2Ids.includes(c.parentId)).map(c => (
                      <label key={c.id} className="flex items-center gap-2 text-[10px] font-bold cursor-pointer hover:bg-slate-50 p-1 rounded">
                        <input 
                          type="checkbox" 
                          checked={newMapping.category3Ids.includes(c.id)}
                          onChange={e => {
                            const ids = e.target.checked 
                              ? [...newMapping.category3Ids, c.id]
                              : newMapping.category3Ids.filter(id => id !== c.id);
                            setNewMapping({ ...newMapping, category3Ids: ids });
                          }}
                        />
                        {getCatName(c, language)}
                      </label>
                    ))}
                  </div>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">{language === 'zh' ? '对应项目分工' : 'Project Allocation'}</label>
                  <select 
                    value={newMapping.allocationId}
                    onChange={e => setNewMapping({...newMapping, allocationId: e.target.value})}
                    className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs font-bold outline-none"
                  >
                    <option value="">Select Allocation</option>
                    {state.allocations.map(a => <option key={a.id} value={a.id}>{language === 'zh' ? a.nameZh : a.nameEn}</option>)}
                  </select>
                </div>
                <div className="pt-6">
                  <button 
                    onClick={addMapping}
                    className={cn(
                      "w-full px-4 py-2 rounded-xl font-black text-[10px] uppercase tracking-widest shadow-lg transition-all",
                      editingMappingId ? "bg-indigo-600 hover:bg-indigo-500 text-white" : "bg-blue-600 hover:bg-blue-500 text-white"
                    )}
                  >
                    {editingMappingId ? (language === 'zh' ? '保存修改' : 'Save Changes') : (language === 'zh' ? '添加对应' : 'Add Mapping')}
                  </button>
                  {editingMappingId && (
                    <button 
                      onClick={() => {
                        setEditingMappingId(null);
                        setNewMapping({ userId: '', category2Ids: [], category3Ids: [], allocationId: '' });
                      }}
                      className="w-full mt-2 text-[9px] font-bold text-slate-400 hover:text-slate-600"
                    >
                      {language === 'zh' ? '取消' : 'Cancel'}
                    </button>
                  )}
                </div>
              </div>

              <div className="overflow-x-auto rounded-2xl border border-slate-100">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-100">
                      <th className="px-6 py-3 text-[9px] font-black text-slate-400 uppercase tracking-widest">Scope</th>
                      <th className="px-6 py-3 text-[9px] font-black text-slate-400 uppercase tracking-widest">Category 2</th>
                      <th className="px-6 py-3 text-[9px] font-black text-slate-400 uppercase tracking-widest">Category 3</th>
                      <th className="px-6 py-3 text-[9px] font-black text-slate-400 uppercase tracking-widest">Auto Allocation</th>
                      <th className="px-6 py-3 text-[9px] font-black text-slate-400 uppercase tracking-widest text-right">Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {state.userMappings
                      .filter(m => {
                        const user = state.currentUser;
                        if (!user) return false;
                        // 易红 (e-2) 看所有人
                        if (user.id === 'admin-1' || user.id === 'e-2') return true;
                        // 麻义俊 (e-10) 看所有蓝领 (BC)
                        if (user.id === 'e-10') {
                          const mappingUser = state.employees.find(e => e.id === m.userId);
                          return !m.userId || mappingUser?.type === 'BC';
                        }
                        // 其他人只看自己或全局
                        return m.userId === user.id || !m.userId || m.userId === '';
                      })
                      .map(m => (
                      <tr key={m.id} className="hover:bg-slate-50/30 transition-colors group">
                        <td className="px-6 py-3 text-[10px] font-bold text-slate-600">
                          {m.userId ? getEmpName(state.employees.find(e => e.id === m.userId), language) : 'Global (All)'}
                        </td>
                        <td className="px-6 py-3 text-[9px] font-bold text-slate-500 max-w-[200px]">
                          <div className="flex flex-wrap gap-1">
                            {m.category2Ids?.map(id => (
                              <span key={id} className="bg-slate-100 px-1.5 py-0.5 rounded">{getCatName(state.categories2.find(c => c.id === id), language)}</span>
                            ))}
                          </div>
                        </td>
                        <td className="px-6 py-3 text-[9px] font-bold text-slate-500 max-w-[300px]">
                          <div className="flex flex-wrap gap-1">
                            {m.category3Ids?.map(id => (
                              <span key={id} className="bg-slate-100 px-1.5 py-0.5 rounded">{getCatName(state.categories3.find(c => c.id === id), language)}</span>
                            ))}
                          </div>
                        </td>
                        <td className="px-6 py-3 text-[10px] font-bold text-blue-600">
                          {state.allocations.find(a => a.id === m.allocationId)?.nameEn || state.allocations.find(a => a.id === m.allocationId)?.nameZh}
                        </td>
                        <td className="px-6 py-3 text-right">
                          <div className="flex items-center justify-end gap-2">
                            <button 
                              onClick={() => startEditMapping(m)}
                              className="p-1.5 text-slate-300 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                            >
                              <Pencil size={12} />
                            </button>
                            <button 
                              onClick={() => {
                                if (window.confirm(language === 'zh' ? '确定删除此映射吗？' : 'Delete this mapping?')) {
                                  const newState = { ...state, userMappings: state.userMappings.filter(um => um.id !== m.id) };
                                  setState(newState);
                                  handleSaveToDatabase(newState);
                                }
                              }}
                              className="p-1.5 text-slate-300 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                            >
                              <Trash2 size={12} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              </div>
            </div>
          </div>
        </div>
    </motion.div>
  );
}

// --- 2. Project Dashboard (Real-time Analytics) ---
export function ProjectDashboard() {
  const context = useContext(AppContext);
  const [dateRange, setDateRange] = useState<{ start: string; end: string }>({
    start: format(startOfMonth(new Date()), 'yyyy-MM-dd'),
    end: format(endOfMonth(new Date()), 'yyyy-MM-dd')
  });
  const [fullCycleProjects, setFullCycleProjects] = useState<Record<string, boolean>>({});
  const [rateDraft, setRateDraft] = useState({ BC: '95', WC: '130' });
  const [accountingProjectId, setAccountingProjectId] = useState<string>('');
  const [selectedWeekKey, setSelectedWeekKey] = useState<string>('');

  if (!context) return null;
  const { state, setState, language, handleBatchImportEntries, handleSaveToDatabase } = context;
  const isSuperAdmin = state.currentUser?.role === 'admin';
  const globalCostRates = {
    BC: Number(state.userSettings?.costRatesGlobal?.BC) || 95,
    WC: Number(state.userSettings?.costRatesGlobal?.WC) || 130
  };

  useEffect(() => {
    setRateDraft({ BC: String(globalCostRates.BC), WC: String(globalCostRates.WC) });
  }, [globalCostRates.BC, globalCostRates.WC]);

  // 用当前输入框的草稿值实时计算人工成本（不必点“保存”也能看到变化）
  const previewCostRates = useMemo(() => ({
    BC: Math.max(1, Number(rateDraft.BC) || globalCostRates.BC),
    WC: Math.max(1, Number(rateDraft.WC) || globalCostRates.WC),
  }), [rateDraft.BC, rateDraft.WC, globalCostRates.BC, globalCostRates.WC]);

  const saveGlobalCostRates = () => {
    const bc = Math.max(1, Number(rateDraft.BC) || globalCostRates.BC);
    const wc = Math.max(1, Number(rateDraft.WC) || globalCostRates.WC);
    setState(prev => {
      const newState = {
        ...prev,
        userSettings: {
          ...(prev.userSettings || {}),
          costRatesGlobal: { BC: bc, WC: wc }
        }
      };
      handleSaveToDatabase(newState);
      return newState;
    });
  };

  const genDemoTimesheetData = () => {
    if (!isSuperAdmin) return;
    if (!window.confirm('将生成一些“模拟工时数据”（仅用于测试导出/看板）。\n\n继续吗？')) return;

    const start = parseISO(dateRange.start);
    const end = parseISO(dateRange.end);
    const days = eachDayOfInterval({ start, end }).filter(d => !isWeekend(d)).slice(0, 12); // 控制数量，避免太多

    const employees = (state.employees || []).filter(e => e.id !== 'admin-1').slice(0, 6);
    const projects = (state.projects || []).filter(p => p.id !== 'p-general' && p.id !== 'p-leave').slice(0, 4);
    const cat2s = (state.categories2 || []).slice(0, 3);
    const cat3s = (state.categories3 || []).slice(0, 6);
    const allocs = (state.allocations || []).slice(0, 4);

    if (!employees.length || !projects.length || !cat2s.length || !cat3s.length || !allocs.length) {
      alert('缺少基础数据：请先确保有员工/项目/分类2/分类3/分工。');
      return;
    }

    const nowIso = new Date().toISOString();
    const mkId = () => `ts-demo-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

    const newEntries: TimesheetEntry[] = [];
    days.forEach((d, di) => {
      const date = format(d, 'yyyy-MM-dd');
      employees.forEach((emp, ei) => {
        // 每天每人 1~2 条，0.5 步进，合计尽量 <= 8
        const n = (ei + di) % 2 === 0 ? 2 : 1;
        let remaining = 8;
        for (let k = 0; k < n; k++) {
          const proj = projects[(ei + di + k) % projects.length];
          const cat2 = cat2s[(ei + k) % cat2s.length];
          const candidates = cat3s.filter(c => c.parentId === cat2.id);
          const cat3 = (candidates.length ? candidates : cat3s)[(di + k) % (candidates.length ? candidates.length : cat3s.length)];
          const alloc = allocs[(ei + di) % allocs.length];
          const hours = Math.max(0.5, Math.min(remaining, (((ei + di + k) % 7) + 1) * 0.5)); // 0.5~4.0
          remaining = Math.max(0, remaining - hours);
          newEntries.push({
            id: mkId(),
            employeeId: emp.id,
            projectId: proj.id,
            category2Id: cat2.id,
            category3Id: cat3.id,
            allocationId: alloc.id,
            date,
            hours,
            comment: `模拟数据：${proj.name} / ${cat2.nameZh} / ${cat3.nameZh}`,
            submittedAt: nowIso
          });
          if (remaining <= 0.5) break;
        }
      });
    });

    const newState = { ...state, entries: [...(state.entries || []), ...newEntries] };
    setState(newState);
    handleSaveToDatabase(newState);
    alert(`已生成模拟工时数据：${newEntries.length} 条（可用于测试导出/统计）。`);
  };

  const clearDemoTimesheetData = () => {
    if (!isSuperAdmin) return;
    const count = (state.entries || []).filter(e => String(e.id || '').startsWith('ts-demo-')).length;
    if (count === 0) return alert('当前没有模拟数据（id 以 ts-demo- 开头）。');
    if (!window.confirm(`将删除 ${count} 条模拟工时数据（id 以 ts-demo- 开头）。\n\n继续吗？`)) return;
    const newState = { ...state, entries: (state.entries || []).filter(e => !String(e.id || '').startsWith('ts-demo-')) };
    setState(newState);
    handleSaveToDatabase(newState);
    alert('已清除模拟工时数据。');
  };

  const setRange = (type: 'thisMonth' | 'lastMonth' | 'thisWeek' | 'lastWeek') => {
    const today = new Date();
    let start, end;
    if (type === 'thisMonth') {
      start = startOfMonth(today);
      end = endOfMonth(today);
    } else if (type === 'lastMonth') {
      const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      start = startOfMonth(lastMonth);
      end = endOfMonth(lastMonth);
    } else if (type === 'thisWeek') {
      start = startOfWeek(today, { weekStartsOn: 1 });
      end = endOfWeek(today, { weekStartsOn: 1 });
    } else if (type === 'lastWeek') {
      const lastWeek = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
      start = startOfWeek(lastWeek, { weekStartsOn: 1 });
      end = endOfWeek(lastWeek, { weekStartsOn: 1 });
    }
    if (start && end) {
      setDateRange({ start: format(start, 'yyyy-MM-dd'), end: format(end, 'yyyy-MM-dd') });
    }
  };

  // 0. Filter projects by manager if not super admin
  const managedProjects = state.projects.filter(p => p.id !== 'p-leave');
  const accountingProjects = managedProjects.filter(p => p.id !== 'p-general' && p.id !== 'p-leave');
  useEffect(() => {
    if (!accountingProjectId && accountingProjects.length > 0) {
      setAccountingProjectId(accountingProjects[0].id);
    }
  }, [accountingProjectId, accountingProjects]);

  const managedProjectIds = managedProjects.map(p => p.id);

  const getRateByType = (type?: Employee['type']) => (type === 'WC' ? previewCostRates.WC : previewCostRates.BC);

    // 1. Filter entries by date range and managed projects
    const filteredEntries = state.entries.filter(e => 
      e.date >= dateRange.start && e.date <= dateRange.end && 
      managedProjectIds.includes(e.projectId)
    );
    
    const start = parseISO(dateRange.start);
    const end = parseISO(dateRange.end);
    const workingDays = eachDayOfInterval({ start, end }).filter(d => !isWeekend(d)).length;
    const theoreticalPerPerson = workingDays * 8;

    const totalActual = filteredEntries.filter(e => e.projectId !== 'p-leave').reduce((acc, e) => acc + e.hours, 0);
    const totalLeave = filteredEntries.filter(e => e.projectId === 'p-leave').reduce((acc, e) => acc + e.hours, 0);
    const uniqueEmployeesInMonth = new Set(filteredEntries.map(e => e.employeeId)).size || 1;
    
    // Global Ratio: Sum(Actual) / Sum(Theoretical - Leave)
    const totalTheoretical = (theoreticalPerPerson * uniqueEmployeesInMonth) - totalLeave;
    const globalRatio = totalTheoretical > 0 ? (totalActual / totalTheoretical) : 0;

    const employeeMetrics = state.employees.map(emp => {
      const empEntries = filteredEntries.filter(e => e.employeeId === emp.id);
      const empAdjs = state.dailyAdjustments.filter(a => a.employeeId === emp.id && a.date >= dateRange.start && a.date <= dateRange.end);
      
      const generalHours = empEntries.filter(e => e.projectId === 'p-general').reduce((acc, e) => acc + e.hours, 0);
      const projectHours = empEntries.filter(e => e.projectId !== 'p-general' && e.projectId !== 'p-leave').reduce((acc, e) => acc + e.hours, 0);
      
      const totalOvertimeHours = empAdjs.reduce((acc, a) => acc + a.overtime, 0);
      const totalLeaveHours = empAdjs.reduce((acc, a) => acc + a.leave, 0);
      
      // Attendance Days = Number of days with actual entries (excluding leave entries)
      const attendanceDaysCount = new Set(empEntries.filter(e => e.projectId !== 'p-leave').map(e => e.date)).size;
      
      // Total Attendance Hours = Attendance Days * 8 + Total Overtime - Total Leave
      const totalAttendanceHours = (attendanceDaysCount * 8) + totalOvertimeHours - totalLeaveHours;
      
      // Time Investment = General Time + Project Time
      const timeInvestment = generalHours + projectHours;
      
      const gongTou = totalAttendanceHours > 0 ? (timeInvestment / totalAttendanceHours) * 100 : 0;
      const xiangTou = timeInvestment > 0 ? (projectHours / timeInvestment) * 100 : 0;

      return {
        id: emp.id,
        name: getEmpName(emp, language),
        totalActualHours: timeInvestment,
        generalHours,
        projectHours,
        attendanceDays: attendanceDaysCount,
        totalAttendanceHours,
        totalOvertimeHours,
        totalLeaveHours,
        gongTou,
        xiangTou
      };
    }).filter(e => e.totalActualHours > 0 || e.projectHours > 0).sort((a, b) => b.totalActualHours - a.totalActualHours);

    // Summary Row: Calculate based on total sums
    const totalAttendanceHoursSum = employeeMetrics.reduce((acc, m) => acc + m.totalAttendanceHours, 0);
    const totalActualHoursSum = employeeMetrics.reduce((acc, m) => acc + m.totalActualHours, 0);
    const totalProjectHoursSum = employeeMetrics.reduce((acc, m) => acc + m.projectHours, 0);

    const totalMetrics = {
      totalActualHours: totalActualHoursSum,
      generalHours: employeeMetrics.reduce((acc, m) => acc + m.generalHours, 0),
      projectHours: totalProjectHoursSum,
      attendanceDays: employeeMetrics.reduce((acc, m) => acc + m.attendanceDays, 0),
      totalAttendanceHours: totalAttendanceHoursSum,
      gongTou: totalAttendanceHoursSum > 0 ? (totalActualHoursSum / totalAttendanceHoursSum) * 100 : 0,
      xiangTou: totalActualHoursSum > 0 ? (totalProjectHoursSum / totalActualHoursSum) * 100 : 0,
    };

  const employeeData = employeeMetrics.map(m => ({ name: m.name, hours: m.totalActualHours }));

  const COLORS = ['#3B82F6', '#8B5CF6', '#EC4899', '#F59E0B', '#10B981', '#6366F1'];
  const pastelColors = [
    'E3F2FD', 'F1F8E9', 'FFFDE7', 'F3E5F5', 'FBE9E7', 'E0F2F1', 'FFF3E0', 'FCE4EC',
    'E8EAF6', 'EFEBE9', 'FAFAFA', 'ECEFF1'
  ];

  // Burn Rate Calculation (Average hours per working day so far)
  const today = new Date();
  const isCurrentRange = format(new Date(), 'yyyy-MM-dd') >= dateRange.start && format(new Date(), 'yyyy-MM-dd') <= dateRange.end;
  const daysPassed = isCurrentRange 
    ? eachDayOfInterval({ start, end: new Date() }).filter(d => !isWeekend(d)).length 
    : workingDays;
  const burnRate = daysPassed > 0 ? (totalActual / daysPassed).toFixed(1) : '0';

  const accountingSummary = useMemo(() => {
    const laborEntries = filteredEntries.filter(e => e.projectId === accountingProjectId);
    const laborHours = laborEntries.reduce((s, e) => s + (Number(e.hours) || 0), 0);
    const laborCost = laborEntries.reduce((s, e) => {
      const emp = state.employees.find(x => x.id === e.employeeId);
      return s + (Number(e.hours) || 0) * getRateByType(emp?.type);
    }, 0);
    const hardwareRows = state.materialRequirements.filter(r => r.projectId === accountingProjectId);
    const hardwareCost = hardwareRows.reduce((s, r) => s + (((Number(r.totalPrice) || 0) > 0 ? Number(r.totalPrice) : (Number(r.unitPrice) || 0) * (Number(r.quantity) || 0))), 0);
    const totalCost = laborCost + hardwareCost;
    return { laborHours, laborCost, hardwareCost, totalCost };
  }, [filteredEntries, accountingProjectId, state.employees, state.materialRequirements, previewCostRates.BC, previewCostRates.WC]);
  const accountingFullCycleSummary = useMemo(() => {
    const laborEntries = state.entries.filter(e => e.projectId === accountingProjectId);
    const laborHours = laborEntries.reduce((s, e) => s + (Number(e.hours) || 0), 0);
    const laborCost = laborEntries.reduce((s, e) => {
      const emp = state.employees.find(x => x.id === e.employeeId);
      return s + (Number(e.hours) || 0) * getRateByType(emp?.type);
    }, 0);
    const hardwareRows = state.materialRequirements.filter(r => r.projectId === accountingProjectId);
    const hardwareCost = hardwareRows.reduce((s, r) => s + (((Number(r.totalPrice) || 0) > 0 ? Number(r.totalPrice) : (Number(r.unitPrice) || 0) * (Number(r.quantity) || 0))), 0);
    const totalCost = laborCost + hardwareCost;
    return { laborHours, laborCost, hardwareCost, totalCost };
  }, [state.entries, accountingProjectId, state.employees, state.materialRequirements, previewCostRates.BC, previewCostRates.WC]);
  const accountingProject = accountingProjects.find(p => p.id === accountingProjectId) || null;
  const accountingExecutionSummary = useMemo(() => {
    const rows = state.productionForecasts.filter(r => r.projectId === accountingProjectId);
    const totalNodes = rows.length;
    const doneNodes = rows.filter(r => r.assemblyStatus === 'Done').length;
    const assemblingNodes = rows.filter(r => r.assemblyStatus === 'Assembling').length;
    const readyNodes = rows.filter(r => r.assemblyStatus === 'ReadyToAssemble').length;
    const blockedNodes = rows.filter(r => r.assemblyStatus === 'Blocked').length;
    const notReadyNodes = rows.filter(r => r.assemblyStatus === 'NotReady').length;
    const highRiskNodes = rows.filter(r => r.risk === 'High').length;
    const nextReadyDate = rows
      .filter(r => r.earliestStart)
      .sort((a, b) => a.earliestStart.localeCompare(b.earliestStart))[0]?.earliestStart || '';
    return {
      totalNodes,
      doneNodes,
      assemblingNodes,
      readyNodes,
      blockedNodes,
      notReadyNodes,
      highRiskNodes,
      nextReadyDate,
      completionRate: totalNodes > 0 ? (doneNodes / totalNodes) * 100 : 0,
      blockedRatio: totalNodes > 0 ? ((blockedNodes + notReadyNodes) / totalNodes) * 100 : 0
    };
  }, [state.productionForecasts, accountingProjectId]);

  const accountingWeeklySeries = useMemo(() => {
    const weekStart = startOfWeek(parseISO(dateRange.start), { weekStartsOn: 1 });
    const weekEnd = startOfWeek(parseISO(dateRange.end), { weekStartsOn: 1 });
    const weeks: { week: string; labor: number; hardware: number; total: number; key: string }[] = [];
    let cursor = weekStart;
    while (cursor.getTime() <= weekEnd.getTime()) {
      weeks.push({ key: format(cursor, 'yyyy-MM-dd'), week: format(cursor, 'MM/dd'), labor: 0, hardware: 0, total: 0 });
      cursor = addWeeks(cursor, 1);
    }
    const map = new Map<string, typeof weeks[number]>();
    weeks.forEach(w => map.set(w.key, w));

    filteredEntries
      .filter(e => e.projectId === accountingProjectId)
      .forEach(e => {
        const emp = state.employees.find(x => x.id === e.employeeId);
        const k = format(startOfWeek(parseISO(e.date), { weekStartsOn: 1 }), 'yyyy-MM-dd');
        const b = map.get(k);
        if (b) b.labor += (Number(e.hours) || 0) * getRateByType(emp?.type);
      });

    state.materialRequirements
      .filter(r => r.projectId === accountingProjectId)
      .forEach(r => {
        const d = r.deliveryTime || r.expectedArrival || r.updatedAt || r.createdAt;
        if (!d) return;
        const k = format(startOfWeek(parseISO(d), { weekStartsOn: 1 }), 'yyyy-MM-dd');
        const b = map.get(k);
        if (b) b.hardware += ((Number(r.totalPrice) || 0) > 0 ? Number(r.totalPrice) : (Number(r.unitPrice) || 0) * (Number(r.quantity) || 0));
      });

    return weeks.map(w => ({ ...w, labor: Math.round(w.labor), hardware: Math.round(w.hardware), total: Math.round(w.labor + w.hardware) }));
  }, [filteredEntries, state.employees, state.materialRequirements, accountingProjectId, dateRange.start, dateRange.end, previewCostRates.BC, previewCostRates.WC]);
  const accountingRiskStats = useMemo(() => {
    const startTs = parseISO(dateRange.start).getTime();
    const endDate = parseISO(dateRange.end);
    endDate.setHours(23, 59, 59, 999);
    const endTs = endDate.getTime();
    const map = new Map<string, number>();
    (state.workflowEvents || [])
      .filter(e => e.projectId === accountingProjectId)
      .filter(e => e.action === 'assembly_status_change' && e.toStatus === 'Blocked')
      .filter(e => {
        const ts = parseISO(e.createdAt).getTime();
        return ts >= startTs && ts <= endTs;
      })
      .forEach(e => {
        const category = (e.message || '').split(':')[0].trim() || '未分类';
        map.set(category, (map.get(category) || 0) + 1);
      });
    return Array.from(map.entries())
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count);
  }, [state.workflowEvents, accountingProjectId, dateRange.start, dateRange.end]);
  useEffect(() => {
    if (accountingWeeklySeries.length > 0) {
      setSelectedWeekKey(accountingWeeklySeries[accountingWeeklySeries.length - 1].key);
    }
  }, [accountingProjectId, dateRange.start, dateRange.end, accountingWeeklySeries.length]);
  const selectedWeekRange = useMemo(() => {
    if (!selectedWeekKey) return null;
    const s = parseISO(selectedWeekKey);
    const e = endOfWeek(s, { weekStartsOn: 1 });
    return { start: s, end: e };
  }, [selectedWeekKey]);
  const weekLaborDetails = useMemo(() => {
    if (!selectedWeekRange) return [] as { employeeName: string; hours: number; amount: number }[];
    return filteredEntries
      .filter(e => e.projectId === accountingProjectId)
      .filter(e => {
        const d = parseISO(e.date);
        return d.getTime() >= selectedWeekRange.start.getTime() && d.getTime() <= selectedWeekRange.end.getTime();
      })
      .map(e => {
        const emp = state.employees.find(x => x.id === e.employeeId);
        const rate = getRateByType(emp?.type);
        return {
          employeeName: getEmpName(emp, 'zh'),
          hours: Number(e.hours) || 0,
          amount: (Number(e.hours) || 0) * rate
        };
      })
      .sort((a, b) => b.amount - a.amount);
  }, [filteredEntries, accountingProjectId, selectedWeekRange, state.employees, previewCostRates.BC, previewCostRates.WC]);
  const weekHardwareDetails = useMemo(() => {
    if (!selectedWeekRange) return [] as { name: string; qty: number; amount: number }[];
    return state.materialRequirements
      .filter(r => r.projectId === accountingProjectId)
      .filter(r => {
        const dt = r.deliveryTime || r.expectedArrival || r.updatedAt || r.createdAt;
        if (!dt) return false;
        const d = parseISO(dt);
        return d.getTime() >= selectedWeekRange.start.getTime() && d.getTime() <= selectedWeekRange.end.getTime();
      })
      .map(r => ({
        name: r.model || r.name || r.code,
        qty: Number(r.quantity) || 0,
        amount: ((Number(r.totalPrice) || 0) > 0 ? Number(r.totalPrice) : (Number(r.unitPrice) || 0) * (Number(r.quantity) || 0))
      }))
      .sort((a, b) => b.amount - a.amount);
  }, [state.materialRequirements, accountingProjectId, selectedWeekRange]);

  const handleExportExcel = () => {
    try {
      const wb = XLSX.utils.book_new();
      
      // ---------------------------
      // Export scope & permissions
      // ---------------------------
      const currentUser = state.currentUser;
      const role = currentUser?.role;
      const isAdminExport = role === 'admin';
      const isPmExport = role === 'pm';

      const dateFilteredAll = state.entries.filter(e =>
        e.date >= dateRange.start && e.date <= dateRange.end
      );

      // PM can export all entries of projects they manage; others only self (admin exports all)
      const managedByMe = new Set(
        (state.projects || [])
          .filter(p => (p as any).managerId && (p as any).managerId === currentUser?.id)
          .map(p => p.id)
      );

      const exportEntries = dateFilteredAll.filter(e => {
        if (isAdminExport) return true;
        if (isPmExport) return managedByMe.has(e.projectId) || e.employeeId === currentUser?.id;
        return e.employeeId === currentUser?.id;
      });

      const headerStyle = {
        fill: { fgColor: { rgb: "334155" } },
        font: { color: { rgb: "FFFFFF" }, bold: true },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
          top: { style: "medium", color: { rgb: "000000" } },
          bottom: { style: "medium", color: { rgb: "000000" } },
          left: { style: "medium", color: { rgb: "000000" } },
          right: { style: "medium", color: { rgb: "000000" } }
        }
      };

      // 仅用于展示“工作工时”，不计入请假项目
      const exportEntriesNoLeave = exportEntries.filter(e => e.projectId !== 'p-leave');

      const exportTotalHours = exportEntriesNoLeave.reduce((s, e) => s + (Number(e.hours) || 0), 0);

      // ---------------------------
      // Sheet 1: 工时填报表（矩阵，明细展开）
      // 纵轴：人名（同一人多行展开，每条工时记录独立一行，不做合计）
      // 横轴：日期 + 星期（两行表头）
      // 行内必须显示：项目、分类2、分类3
      // 每个人的背景色区分；备注作为批注挂在对应日期单元格上
      // ---------------------------
      const dates = eachDayOfInterval({ start: parseISO(dateRange.start), end: parseISO(dateRange.end) }).map(d => format(d, 'yyyy-MM-dd'));
      const weekZh = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];

      const employeesUsed = state.employees
        .filter(e => e.id !== 'admin-1')
        .filter(emp => exportEntriesNoLeave.some(e => e.employeeId === emp.id))
        .sort((a, b) => getEmpName(a, 'zh').localeCompare(getEmpName(b, 'zh')));

      const computeMetricsForEmp = (empId: string) => {
        const empEntries = exportEntriesNoLeave.filter(e => e.employeeId === empId);
        const empAdjs = state.dailyAdjustments.filter(a => a.employeeId === empId && a.date >= dateRange.start && a.date <= dateRange.end);
        const generalHours = empEntries.filter(e => e.projectId === 'p-general').reduce((acc, e) => acc + (Number(e.hours) || 0), 0);
        const projectHours = empEntries.filter(e => e.projectId !== 'p-general').reduce((acc, e) => acc + (Number(e.hours) || 0), 0);
        const totalOvertimeHours = empAdjs.reduce((acc, a) => acc + (Number(a.overtime) || 0), 0);
        const totalLeaveHours = empAdjs.reduce((acc, a) => acc + (Number(a.leave) || 0), 0);
        const attendanceDaysCount = new Set(empEntries.map(e => e.date)).size;
        const totalAttendanceHours = (attendanceDaysCount * 8) + totalOvertimeHours - totalLeaveHours;
        const timeInvestment = generalHours + projectHours;
        const gongTou = totalAttendanceHours > 0 ? (timeInvestment / totalAttendanceHours) : 0;
        const xiangTou = timeInvestment > 0 ? (projectHours / timeInvestment) : 0;
        return { timeInvestment, gongTou, xiangTou };
      };

      const matrixData: any[] = [];
      const headerRow1 = [
        '姓名',
        '项目',
        '分类2',
        '分类3',
        ...dates.map(d => format(parseISO(d), 'MM/dd'))
      ];
      const headerRow2 = [
        '',
        '',
        '',
        '',
        ...dates.map(d => weekZh[parseISO(d).getDay()])
      ];
      matrixData.push(headerRow1);
      matrixData.push(headerRow2);

      const wsMatrix = XLSX.utils.aoa_to_sheet(matrixData);
      headerRow1.forEach((_, c) => {
        const cell1 = XLSX.utils.encode_cell({ r: 0, c });
        const cell2 = XLSX.utils.encode_cell({ r: 1, c });
        if (wsMatrix[cell1]) wsMatrix[cell1].s = headerStyle;
        if (wsMatrix[cell2]) wsMatrix[cell2].s = headerStyle;
      });

      const pastelColors = [
        'E3F2FD', 'F1F8E9', 'FFFDE7', 'F3E5F5', 'FBE9E7', 'E0F2F1', 'FFF3E0', 'FCE4EC',
        'E8EAF6', 'EFEBE9', 'FAFAFA', 'ECEFF1'
      ];

      // Prepare sorted entry rows
      const entryRows = exportEntriesNoLeave
        .slice()
        .sort((a, b) => {
          const ea = state.employees.find(x => x.id === a.employeeId);
          const eb = state.employees.find(x => x.id === b.employeeId);
          const na = getEmpName(ea, 'zh');
          const nb = getEmpName(eb, 'zh');
          if (na !== nb) return na.localeCompare(nb);
          if (a.date !== b.date) return a.date.localeCompare(b.date);
          // stable sort by project/category
          const pa = (state.projects.find(p => p.id === a.projectId)?.name || a.projectId);
          const pb = (state.projects.find(p => p.id === b.projectId)?.name || b.projectId);
          if (pa !== pb) return pa.localeCompare(pb);
          const c2a = getCatName(state.categories2.find(c => c.id === a.category2Id), 'zh');
          const c2b = getCatName(state.categories2.find(c => c.id === b.category2Id), 'zh');
          if (c2a !== c2b) return c2a.localeCompare(c2b);
          const c3a = getCatName(state.categories3.find(c => c.id === a.category3Id), 'zh');
          const c3b = getCatName(state.categories3.find(c => c.id === b.category3Id), 'zh');
          if (c3a !== c3b) return c3a.localeCompare(c3b);
          return String(a.id).localeCompare(String(b.id));
        });

      const empColorById = new Map<string, string>();
      employeesUsed.forEach((e, idx) => empColorById.set(e.id, pastelColors[idx % pastelColors.length]));

      const merges: any[] = [];
      let currentRow = 2;
      let groupStartRow = 2;
      let currentEmpId: string | null = null;

      const writeRowStyle = (rowIndex: number, baseColor: string) => {
        // Apply background color and borders across the entire row width
        const colCount = 4 + dates.length;
        for (let cIdx = 0; cIdx < colCount; cIdx++) {
          const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c: cIdx });
          if (!wsMatrix[cellRef]) wsMatrix[cellRef] = { v: '', t: 's' };
          wsMatrix[cellRef].s = {
            fill: { fgColor: { rgb: baseColor } },
            font: { sz: 10, bold: cIdx === 0 },
            alignment: { vertical: "center", horizontal: cIdx <= 3 ? "left" : "center", wrapText: true },
            border: {
              top: { style: "thin", color: { rgb: "CBD5E1" } },
              bottom: { style: "thin", color: { rgb: "CBD5E1" } },
              left: { style: "thin", color: { rgb: "CBD5E1" } },
              right: { style: "thin", color: { rgb: "CBD5E1" } }
            },
            numFmt: cIdx >= 4 ? "0.0" : undefined
          };
        }
      };

      entryRows.forEach((e, idx) => {
        const emp = state.employees.find(x => x.id === e.employeeId);
        const empName = getEmpName(emp, 'zh');
        const projName = state.projects.find(p => p.id === e.projectId)?.name || e.projectId;
        const cat2Name = getCatName(state.categories2.find(c => c.id === e.category2Id), 'zh');
        const cat3Name = getCatName(state.categories3.find(c => c.id === e.category3Id), 'zh');

        // Handle merge groups for the "姓名" column
        if (currentEmpId === null) {
          currentEmpId = e.employeeId;
          groupStartRow = currentRow;
        } else if (currentEmpId !== e.employeeId) {
          if (currentRow - groupStartRow > 1) {
            merges.push({ s: { r: groupStartRow, c: 0 }, e: { r: currentRow - 1, c: 0 } });
          }
          currentEmpId = e.employeeId;
          groupStartRow = currentRow;
        }

        const row: any[] = new Array(4 + dates.length).fill('');
        row[0] = empName;
        row[1] = projName;
        row[2] = cat2Name;
        row[3] = cat3Name;
        const dateIdx = dates.indexOf(e.date);
        if (dateIdx >= 0) row[4 + dateIdx] = parseFloat((Number(e.hours) || 0).toFixed(1));

        XLSX.utils.sheet_add_aoa(wsMatrix, [row], { origin: currentRow });

        const baseColor = empColorById.get(e.employeeId) || 'FAFAFA';
        writeRowStyle(currentRow, baseColor);

        // Add comment to the specific date cell
        if (dateIdx >= 0) {
          const cellRef = XLSX.utils.encode_cell({ r: currentRow, c: 4 + dateIdx });
          const alloc = state.allocations.find(a => a.id === e.allocationId);
          const noteParts: string[] = [];
          if (alloc) noteParts.push(`分工：${alloc.nameZh || alloc.nameEn}`);
          const comment = (e.comment || '').trim();
          if (comment) noteParts.push(`备注：${comment}`);
          const note = noteParts.join('；');
          if (note) {
            wsMatrix[cellRef].c = [{ t: note, a: "TMS" }];
          }
        }

        currentRow++;
      });

      // finalize last merge group
      if (currentEmpId !== null && currentRow - groupStartRow > 1) {
        merges.push({ s: { r: groupStartRow, c: 0 }, e: { r: currentRow - 1, c: 0 } });
      }
      if (merges.length > 0) {
        (wsMatrix as any)['!merges'] = merges;
      }

      wsMatrix['!cols'] = [
        { wch: 12 }, // 姓名
        { wch: 28 }, // 项目
        { wch: 18 }, // 分类2
        { wch: 20 }, // 分类3
        ...dates.map(() => ({ wch: 8 })) // 日期
      ];
      XLSX.utils.book_append_sheet(wb, wsMatrix, "工时填报表");

      // ---------------------------
      // Sheet 2: 汇总（人员+项目+合计）
      // ---------------------------
      const summaryData: any[] = [];
      summaryData.push(['导出范围', `${dateRange.start} ~ ${dateRange.end}`]);
      summaryData.push(['导出人', currentUser ? getEmpName(state.employees.find(e => e.id === currentUser.id), 'zh') : '' ]);
      summaryData.push([]);
      summaryData.push(['人员汇总']);
      summaryData.push(['姓名', '总工时', 'GT(工投)', 'XT(项投)']);

      const empSummary = employeesUsed.map(emp => {
        const m = computeMetricsForEmp(emp.id);
        return {
          name: getEmpName(emp, 'zh'),
          hours: parseFloat(m.timeInvestment.toFixed(1)),
          gt: m.gongTou,
          xt: m.xiangTou
        };
      }).sort((a, b) => b.hours - a.hours);

      empSummary.forEach(r => summaryData.push([r.name, r.hours, r.gt, r.xt]));
      summaryData.push(['合计', parseFloat(exportTotalHours.toFixed(1)), '', '']);
      summaryData.push([]);
      summaryData.push(['项目汇总（当月所有项目，按权限可见范围）']);
      summaryData.push(['项目', '总工时', '占比']);

      const projectTotals = (state.projects || [])
        .map(p => {
          const total = exportEntriesNoLeave.filter(e => e.projectId === p.id).reduce((sum, e) => sum + (Number(e.hours) || 0), 0);
          return { name: p.name, total };
        })
        .filter(p => p.total > 0)
        .sort((a, b) => b.total - a.total);
      projectTotals.forEach(p => summaryData.push([p.name, parseFloat(p.total.toFixed(1)), exportTotalHours > 0 ? (p.total / exportTotalHours) : 0]));

      const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
      // Basic styling
      summaryData.forEach((row, rIdx) => {
        row.forEach((_, cIdx) => {
          const cellRef = XLSX.utils.encode_cell({ r: rIdx, c: cIdx });
          if (!wsSummary[cellRef]) return;
          const v = row[0];
          const isSection = typeof v === 'string' && (v.includes('汇总') || v === '人员汇总');
          const isHeader = Array.isArray(row) && rIdx > 0 && (row[0] === '姓名' || row[0] === '项目');
          if (isHeader) wsSummary[cellRef].s = headerStyle;
          if (isSection) {
            wsSummary[cellRef].s = {
              font: { bold: true, sz: 12, color: { rgb: "0F172A" } },
              alignment: { horizontal: "left", vertical: "center" }
            };
          }
          if (cIdx >= 2 && (row[0] === '姓名' || (typeof row[0] === 'string' && empSummary.some(e => e.name === row[0])))) {
            // percent columns in employee summary
            wsSummary[cellRef].s = { ...(wsSummary[cellRef].s || {}), numFmt: (cIdx >= 2 ? "0.0%" : "0.0") };
          }
          if (row[0] === '项目' && cIdx === 2) {
            wsSummary[cellRef].s = { ...(wsSummary[cellRef].s || {}), numFmt: "0.0%" };
          }
        });
      });
      wsSummary['!cols'] = [{ wch: 18 }, { wch: 14 }, { wch: 12 }, { wch: 12 }];
      XLSX.utils.book_append_sheet(wb, wsSummary, "汇总");

      XLSX.writeFile(wb, `TMS_工时导出_${dateRange.start}_to_${dateRange.end}.xlsx`);
    } catch (err) {
      console.error('Export failed:', err);
      alert(language === 'zh' ? '导出失败，请检查数据。' : 'Export failed. Please check data.');
    }
  };

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="w-full space-y-12">
      <div className="dashboard-card p-4">
        <div className="flex flex-col xl:flex-row xl:items-center xl:justify-between gap-4">
          <div className="flex flex-wrap items-center gap-3">
            <div className="flex bg-white p-1 rounded-2xl border border-slate-200 shadow-sm">
              {[
                { id: 'thisWeek', label: language === 'zh' ? '本周' : 'This Week' },
                { id: 'lastWeek', label: language === 'zh' ? '上周' : 'Last Week' },
                { id: 'thisMonth', label: language === 'zh' ? '本月' : 'This Month' },
                { id: 'lastMonth', label: language === 'zh' ? '上月' : 'Last Month' }
              ].map(range => (
                <button
                  key={range.id}
                  onClick={() => setRange(range.id as any)}
                  className={cn(
                    "px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all",
                    dateRange.start === format(range.id === 'thisMonth' ? startOfMonth(new Date()) : range.id === 'lastMonth' ? startOfMonth(new Date(new Date().getFullYear(), new Date().getMonth() - 1, 1)) : range.id === 'thisWeek' ? startOfWeek(new Date(), { weekStartsOn: 1 }) : startOfWeek(new Date(new Date().getTime() - 7 * 24 * 60 * 60 * 1000), { weekStartsOn: 1 }), 'yyyy-MM-dd')
                      ? "bg-slate-900 text-white shadow-sm"
                      : "text-slate-500 hover:bg-slate-50"
                  )}
                >
                  {range.label}
                </button>
              ))}
            </div>
            <div className="bg-white px-4 py-2 rounded-2xl border border-slate-200 flex items-center gap-3 shadow-sm">
              <Calendar className="text-blue-500" size={16} />
              <input 
                type="date" 
                value={dateRange.start}
                onChange={e => setDateRange({ ...dateRange, start: e.target.value })}
                className="font-black text-slate-700 outline-none bg-transparent text-[10px]"
              />
              <span className="text-slate-300">→</span>
              <input 
                type="date" 
                value={dateRange.end}
                onChange={e => setDateRange({ ...dateRange, end: e.target.value })}
                className="font-black text-slate-700 outline-none bg-transparent text-[10px]"
              />
            </div>
            <button 
              onClick={handleExportExcel}
              className="btn-primary py-2.5 px-4 text-xs"
            >
              <Download size={16} />
              导出
            </button>
            {isSuperAdmin && (
              <>
                <button onClick={genDemoTimesheetData} className="btn-secondary py-2.5 px-4 text-xs">
                  生成模拟工时
                </button>
                <button onClick={clearDemoTimesheetData} className="btn-secondary py-2.5 px-4 text-xs text-rose-600">
                  清除模拟工时
                </button>
              </>
            )}
            <label className="btn-secondary py-2.5 px-4 text-xs cursor-pointer">
              <Upload size={16} />
              汇总
              <input type="file" multiple className="hidden" onChange={e => e.target.files && handleBatchImportEntries(e.target.files)} />
            </label>
          </div>
          <div className="flex flex-wrap items-center gap-2 rounded-2xl border border-slate-200 bg-slate-50/80 px-3 py-2 shadow-sm">
            <span className="text-[10px] font-black text-slate-400 uppercase">项目核算单价</span>
            <div className="h-4 w-px bg-slate-200" />
              <span className="text-[10px] font-black text-slate-500">BC</span>
              <input value={rateDraft.BC} onChange={e => setRateDraft(prev => ({ ...prev, BC: e.target.value }))} className="input-field h-7 py-0.5 px-1.5 text-[10px] w-16" />
              <span className="text-[10px] font-black text-slate-500">WC</span>
              <input value={rateDraft.WC} onChange={e => setRateDraft(prev => ({ ...prev, WC: e.target.value }))} className="input-field h-7 py-0.5 px-1.5 text-[10px] w-16" />
            <button onClick={saveGlobalCostRates} className="btn-secondary py-1 px-2 text-[10px] font-black">
              保存
            </button>
          </div>
        </div>
      </div>

      {/* Employee Metrics Table */}
      <section className="space-y-6">
        <div className="dashboard-card overflow-hidden">
          <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-3">
            <div className="text-sm font-bold text-slate-900">
              团队效率对比
              <span className="text-xs font-medium text-slate-500 ml-2">({dateRange.start} ~ {dateRange.end})</span>
            </div>
          </div>
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="dashboard-table-head">
                <th className="dashboard-table-head-cell">姓名</th>
                <th className="dashboard-table-head-cell">出勤天数</th>
                <th className="dashboard-table-head-cell">总出勤 (h)</th>
                <th className="dashboard-table-head-cell">时间投入 (h)</th>
                <th className="dashboard-table-head-cell">GT (工投)</th>
                <th className="dashboard-table-head-cell">XT (项投)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {employeeMetrics.map(emp => (
                <tr key={emp.id} className="hover:bg-slate-50/50 transition-colors group">
                  <td className="dashboard-table-cell">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 bg-slate-100 rounded-xl flex items-center justify-center text-slate-400 font-black text-sm group-hover:bg-blue-600 group-hover:text-white transition-colors">
                        {emp.name[0]}
                      </div>
                      <div>
                        <p className="font-black text-slate-900">{emp.name}</p>
                        <p className="text-[8px] font-black text-slate-400 uppercase tracking-widest">{state.employees.find(e => e.id === emp.id)?.type || 'WC'}</p>
                      </div>
                    </div>
                  </td>
                  <td className="dashboard-table-cell font-bold text-slate-600">{emp.attendanceDays}d</td>
                  <td className="dashboard-table-cell font-bold text-slate-900">{emp.totalAttendanceHours.toFixed(1)}h</td>
                  <td className="dashboard-table-cell font-bold text-blue-600">{emp.totalActualHours.toFixed(1)}h</td>
                  <td className="dashboard-table-cell">
                    <span 
                      className={cn("px-3 py-1 rounded-full border text-xs font-black inline-block", getGTXTColor(emp.gongTou))}
                      title={`计算: 时间投入(${emp.totalActualHours.toFixed(1)}) / 总出勤工时(${emp.totalAttendanceHours.toFixed(1)})`}
                    >
                      {emp.gongTou.toFixed(1)}%
                    </span>
                  </td>
                  <td className="dashboard-table-cell">
                    <span 
                      className={cn("px-3 py-1 rounded-full border text-xs font-black inline-block", getGTXTColor(emp.xiangTou))}
                      title={`计算: 项目工时(${emp.projectHours.toFixed(1)}) / 时间投入(${emp.totalActualHours.toFixed(1)})`}
                    >
                      {emp.xiangTou.toFixed(1)}%
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot className="dashboard-table-foot">
              <tr>
                <td className="dashboard-table-cell">总计</td>
                <td className="dashboard-table-cell">{totalMetrics.attendanceDays}d</td>
                <td className="dashboard-table-cell">{totalMetrics.totalAttendanceHours.toFixed(1)}h</td>
                <td className="dashboard-table-cell text-blue-600">{totalMetrics.totalActualHours.toFixed(1)}h</td>
                <td className="dashboard-table-cell">
                  <span className={cn("px-3 py-1 rounded-full border text-xs font-black inline-block", getGTXTColor(totalMetrics.gongTou))}>
                    {totalMetrics.gongTou.toFixed(1)}%
                  </span>
                </td>
                <td className="dashboard-table-cell">
                  <span className={cn("px-3 py-1 rounded-full border text-xs font-black inline-block", getGTXTColor(totalMetrics.xiangTou))}>
                    {totalMetrics.xiangTou.toFixed(1)}%
                  </span>
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </section>

      <section className="space-y-4">
        <div className="dashboard-card p-6 space-y-6">
          <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
            <div>
              <p className="dashboard-kicker mb-2">项目总览</p>
              <h3 className="dashboard-title">{accountingProject?.name || '项目总览'}</h3>
              <p className="dashboard-subtitle">上半部分区分当前范围与全周期累计，下半部分保留项目执行与当前范围风险。</p>
            </div>
            <div className="flex flex-wrap items-center gap-3">
              <select value={accountingProjectId} onChange={e => setAccountingProjectId(e.target.value)} className="input-field h-10 py-1 px-3 text-xs font-bold min-w-[240px]">
                {accountingProjects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
              <div className="dashboard-subcard px-4 py-2">
                <p className="dashboard-kicker">项目经理</p>
                <p className="text-sm font-black text-slate-700">{getEmpName(state.employees.find(e => e.id === accountingProject?.managerId) || null, language)}</p>
              </div>
            </div>
          </div>
          <div className="space-y-3">
            <div>
              <p className="dashboard-kicker mb-2">当前范围</p>
              <div className="grid grid-cols-2 xl:grid-cols-4 gap-3">
                <div className="rounded-[1.75rem] border border-blue-100 p-4 bg-blue-50/70"><p className="text-[10px] font-black text-blue-400 uppercase tracking-[0.18em]">人工工时</p><p className="mt-2 text-2xl font-black text-blue-700">{Math.round(accountingSummary.laborHours)}h</p><p className="mt-1 text-xs text-blue-500">{dateRange.start} ~ {dateRange.end}</p></div>
                <div className="rounded-[1.75rem] border border-sky-100 p-4 bg-sky-50/70"><p className="text-[10px] font-black text-sky-400 uppercase tracking-[0.18em]">人工成本</p><p className="mt-2 text-2xl font-black text-sky-700">¥{Math.round(accountingSummary.laborCost).toLocaleString()}</p><p className="mt-1 text-xs text-sky-500">按 BC/WC 单价计算</p></div>
                <div className="rounded-[1.75rem] border border-indigo-100 p-4 bg-indigo-50/70"><p className="text-[10px] font-black text-indigo-400 uppercase tracking-[0.18em]">硬件成本</p><p className="mt-2 text-2xl font-black text-indigo-700">¥{Math.round(accountingSummary.hardwareCost).toLocaleString()}</p><p className="mt-1 text-xs text-indigo-500">当前范围已计入物料</p></div>
                <div className="rounded-[1.75rem] border border-violet-100 p-4 bg-violet-50/70"><p className="text-[10px] font-black text-violet-400 uppercase tracking-[0.18em]">总成本</p><p className="mt-2 text-2xl font-black text-violet-700">¥{Math.round(accountingSummary.totalCost).toLocaleString()}</p><p className="mt-1 text-xs text-violet-500">人工 + 硬件</p></div>
              </div>
            </div>

            <div>
              <p className="dashboard-kicker mb-2">全周期累计</p>
              <div className="grid grid-cols-2 xl:grid-cols-4 gap-3">
                <div className="rounded-[1.75rem] border border-slate-200 p-4 bg-slate-50/80"><p className="text-[10px] font-black text-slate-400 uppercase tracking-[0.18em]">累计人工工时</p><p className="mt-2 text-2xl font-black text-slate-900">{Math.round(accountingFullCycleSummary.laborHours)}h</p><p className="mt-1 text-xs text-slate-500">项目至今全部提报</p></div>
                <div className="rounded-[1.75rem] border border-slate-200 p-4 bg-slate-50/80"><p className="text-[10px] font-black text-slate-400 uppercase tracking-[0.18em]">累计人工成本</p><p className="mt-2 text-2xl font-black text-slate-900">¥{Math.round(accountingFullCycleSummary.laborCost).toLocaleString()}</p><p className="mt-1 text-xs text-slate-500">项目全周期人工累计</p></div>
                <div className="rounded-[1.75rem] border border-slate-200 p-4 bg-slate-50/80"><p className="text-[10px] font-black text-slate-400 uppercase tracking-[0.18em]">累计硬件成本</p><p className="mt-2 text-2xl font-black text-slate-900">¥{Math.round(accountingFullCycleSummary.hardwareCost).toLocaleString()}</p><p className="mt-1 text-xs text-slate-500">项目全周期物料累计</p></div>
                <div className="rounded-[1.75rem] border border-slate-200 p-4 bg-slate-50/80"><p className="text-[10px] font-black text-slate-400 uppercase tracking-[0.18em]">累计总成本</p><p className="mt-2 text-2xl font-black text-slate-900">¥{Math.round(accountingFullCycleSummary.totalCost).toLocaleString()}</p><p className="mt-1 text-xs text-slate-500">项目到当前的总投入</p></div>
              </div>
            </div>

            <div>
              <p className="dashboard-kicker mb-2">执行闭环（全项目）</p>
              <div className="grid grid-cols-2 xl:grid-cols-4 gap-3">
                <div className="rounded-[1.75rem] border border-emerald-100 p-4 bg-emerald-50/70"><p className="text-[10px] font-black text-emerald-400 uppercase tracking-[0.18em]">装配完成率</p><p className="mt-2 text-2xl font-black text-emerald-700">{accountingExecutionSummary.completionRate.toFixed(1)}%</p><p className="mt-1 text-xs text-emerald-500">完工 {accountingExecutionSummary.doneNodes} / {accountingExecutionSummary.totalNodes || 0}</p></div>
                <div className="rounded-[1.75rem] border border-amber-100 p-4 bg-amber-50/70"><p className="text-[10px] font-black text-amber-400 uppercase tracking-[0.18em]">可开工节点</p><p className="mt-2 text-2xl font-black text-amber-700">{accountingExecutionSummary.readyNodes}</p><p className="mt-1 text-xs text-amber-500">{accountingExecutionSummary.nextReadyDate ? `最早 ${accountingExecutionSummary.nextReadyDate}` : '暂无可开工日期'}</p></div>
                <div className="rounded-[1.75rem] border border-cyan-100 p-4 bg-cyan-50/70"><p className="text-[10px] font-black text-cyan-400 uppercase tracking-[0.18em]">装配中节点</p><p className="mt-2 text-2xl font-black text-cyan-700">{accountingExecutionSummary.assemblingNodes}</p><p className="mt-1 text-xs text-cyan-500">正在执行的装配任务</p></div>
                <div className="rounded-[1.75rem] border border-rose-100 p-4 bg-rose-50/70"><p className="text-[10px] font-black text-rose-400 uppercase tracking-[0.18em]">阻塞占比</p><p className="mt-2 text-2xl font-black text-rose-700">{accountingExecutionSummary.blockedRatio.toFixed(1)}%</p><p className="mt-1 text-xs text-rose-500">阻塞/未就绪 {accountingExecutionSummary.blockedNodes + accountingExecutionSummary.notReadyNodes}</p></div>
              </div>
            </div>
          </div>

          <div className="grid grid-cols-1 xl:grid-cols-[minmax(0,1.25fr)_minmax(320px,0.75fr)] gap-4">
            <div className="dashboard-card bg-slate-50/40 p-4">
              <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
                <div>
                  <p className="dashboard-kicker">成本趋势</p>
                  <h4 className="text-lg font-black text-slate-900">按周成本</h4>
                </div>
                <div className="text-right">
                  <p className="dashboard-kicker">当前周期总成本</p>
                  <p className="text-lg font-black text-slate-900">¥{Math.round(accountingSummary.totalCost).toLocaleString()}</p>
                </div>
              </div>
              <div className="h-[220px]">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={accountingWeeklySeries}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#eef2ff" />
                    <XAxis dataKey="week" />
                    <YAxis />
                    <Tooltip />
                    <Bar dataKey="labor" fill="#3b82f6" name="人工成本" radius={[6, 6, 0, 0]} />
                    <Bar dataKey="hardware" fill="#8b5cf6" name="硬件成本" radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="mt-3 flex flex-wrap gap-2">
                {accountingWeeklySeries.map(w => (
                  <button
                    key={w.key}
                    onClick={() => setSelectedWeekKey(w.key)}
                    className={cn(
                      "px-2 py-1 rounded-lg text-[10px] font-black border",
                      selectedWeekKey === w.key ? "bg-slate-900 text-white border-slate-900" : "bg-white text-slate-600 border-slate-200"
                    )}
                  >
                    {w.week}
                  </button>
                ))}
              </div>
            </div>

            <div className="space-y-4">
              <div className="dashboard-card p-4">
                <p className="dashboard-kicker">执行状态</p>
                <div className="mt-3 space-y-3">
                  {[
                    { label: '完工', count: accountingExecutionSummary.doneNodes, ratio: accountingExecutionSummary.completionRate, tone: 'emerald' },
                    { label: '装配中', count: accountingExecutionSummary.assemblingNodes, ratio: accountingExecutionSummary.totalNodes > 0 ? (accountingExecutionSummary.assemblingNodes / accountingExecutionSummary.totalNodes) * 100 : 0, tone: 'cyan' },
                    { label: '可开工', count: accountingExecutionSummary.readyNodes, ratio: accountingExecutionSummary.totalNodes > 0 ? (accountingExecutionSummary.readyNodes / accountingExecutionSummary.totalNodes) * 100 : 0, tone: 'amber' },
                    { label: '阻塞', count: accountingExecutionSummary.blockedNodes + accountingExecutionSummary.notReadyNodes, ratio: accountingExecutionSummary.blockedRatio, tone: 'rose' }
                  ].map(item => (
                    <div key={item.label} className="space-y-1">
                      <div className="flex items-center justify-between text-xs">
                        <span className="font-black text-slate-700">{item.label}</span>
                        <span className="font-bold text-slate-500">{item.count} / {accountingExecutionSummary.totalNodes || 0}</span>
                      </div>
                      <div className="h-2 rounded-full bg-slate-100 overflow-hidden">
                        <div
                          className={cn(
                            "h-full rounded-full transition-all",
                            item.tone === 'emerald' && "bg-emerald-500",
                            item.tone === 'cyan' && "bg-cyan-500",
                            item.tone === 'amber' && "bg-amber-500",
                            item.tone === 'rose' && "bg-rose-500"
                          )}
                          style={{ width: `${Math.min(item.ratio, 100)}%` }}
                        />
                      </div>
                    </div>
                  ))}
                  {accountingExecutionSummary.totalNodes === 0 && (
                    <p className="text-xs text-slate-400">当前项目还没有生成排班闭环数据。</p>
                  )}
                </div>
                <div className="mt-4 grid grid-cols-2 gap-3">
                  <div className="dashboard-subcard p-3">
                    <p className="dashboard-kicker">节点总数</p>
                    <p className="mt-1 text-xl font-black text-slate-900">{accountingExecutionSummary.totalNodes}</p>
                  </div>
                  <div className="dashboard-subcard p-3">
                    <p className="dashboard-kicker">高风险节点</p>
                    <p className="mt-1 text-xl font-black text-rose-600">{accountingExecutionSummary.highRiskNodes}</p>
                  </div>
                </div>
              </div>

              <div className="dashboard-card border-rose-100 bg-rose-50/50 p-4">
                <p className="dashboard-kicker text-rose-400 mb-2">风险分类（当前范围阻塞事件）</p>
                <div className="flex flex-wrap gap-2">
                  {accountingRiskStats.length === 0 && <span className="text-xs text-slate-400">当前周期无阻塞事件</span>}
                  {accountingRiskStats.map(item => (
                    <span key={item.name} className="px-2 py-1 rounded-lg text-[10px] font-black border border-rose-200 bg-white text-rose-700">
                      {item.name} {item.count}
                    </span>
                  ))}
                </div>
              </div>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div className="dashboard-card bg-slate-50/70 p-3">
              <p className="dashboard-kicker mb-2">周人工明细</p>
              <div className="space-y-1 max-h-[180px] overflow-y-auto">
                {weekLaborDetails.length === 0 && <p className="text-xs text-slate-400">该周无人工成本</p>}
                {weekLaborDetails.map((r, idx) => (
                  <div key={`wl-${idx}`} className="flex items-center justify-between text-xs">
                    <span className="font-bold text-slate-700">{r.employeeName}</span>
                    <span className="text-slate-500">{r.hours.toFixed(1)}h / ¥{Math.round(r.amount).toLocaleString()}</span>
                  </div>
                ))}
              </div>
            </div>
            <div className="dashboard-card bg-slate-50/70 p-3">
              <p className="dashboard-kicker mb-2">周硬件明细</p>
              <div className="space-y-1 max-h-[180px] overflow-y-auto">
                {weekHardwareDetails.length === 0 && <p className="text-xs text-slate-400">该周无硬件成本</p>}
                {weekHardwareDetails.map((r, idx) => (
                  <div key={`wh-${idx}`} className="flex items-center justify-between text-xs">
                    <span className="font-bold text-slate-700">{r.name}</span>
                    <span className="text-slate-500">{r.qty} / ¥{Math.round(r.amount).toLocaleString()}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* Project Budget Execution Details */}
      <section className="space-y-8">
        <h3 className="dashboard-title flex items-center gap-3">
          <Target className="text-blue-600" />
          项目预算执行汇总
          <span className="text-sm font-bold text-slate-400 ml-2">({dateRange.start} ~ {dateRange.end})</span>
        </h3>
        <div className="grid grid-cols-1 gap-8">
          {managedProjects
            .filter(p => p.id !== 'p-general' && p.id !== 'p-leave')
            .filter(p => p.visible !== false)
            .sort((a, b) => (b.sortOrder || 0) - (a.sortOrder || 0))
            .map(proj => {
              const projEntries = filteredEntries.filter(e => e.projectId === proj.id);
              const totalHours = projEntries.reduce((acc, e) => acc + e.hours, 0);
              
              // 全生命周期数据
              const allTimeEntries = state.entries.filter(e => e.projectId === proj.id);
              const allTimeHours = allTimeEntries.reduce((acc, e) => acc + e.hours, 0);
              const totalBudget = Object.values(proj.budgets).reduce((acc, b) => acc + b, 0);
              const totalSpanPercent = totalBudget > 0 ? (allTimeHours / totalBudget) * 100 : 0;

              // BC/WC 拆分
              const allTimeBC = allTimeEntries.filter(e => state.employees.find(emp => emp.id === e.employeeId)?.type === 'BC').reduce((acc, e) => acc + e.hours, 0);
              const allTimeWC = allTimeEntries.filter(e => state.employees.find(emp => emp.id === e.employeeId)?.type === 'WC').reduce((acc, e) => acc + e.hours, 0);
              
              const budgetBC = Object.entries(proj.budgets).reduce((acc, [allocId, b]) => {
                const alloc = state.allocations.find(a => a.id === allocId);
                return alloc?.type === 'BC' ? acc + b : acc;
              }, 0);
              const budgetWC = Object.entries(proj.budgets).reduce((acc, [allocId, b]) => {
                const alloc = state.allocations.find(a => a.id === allocId);
                return alloc?.type === 'WC' ? acc + b : acc;
              }, 0);

              const xbBC = budgetBC > 0 ? (allTimeBC / budgetBC) * 100 : 0;
              const xbWC = budgetWC > 0 ? (allTimeWC / budgetWC) * 100 : 0;

              const isFullCycle = fullCycleProjects[proj.id];
              const displayEntries = isFullCycle ? allTimeEntries : projEntries;
              const displayTotalHours = displayEntries.reduce((acc, e) => acc + e.hours, 0);

              // 人员区分
              const personBreakdown = state.employees.map(emp => {
                const empProjEntries = displayEntries.filter(e => e.employeeId === emp.id);
                const empProjHours = empProjEntries.reduce((acc, e) => acc + e.hours, 0);
                
                // Calculate budget for this employee on this project based on their allocations
                const empAllocIds = Array.from(new Set(empProjEntries.map(e => e.allocationId)));
                const empBudget = empAllocIds.reduce((acc, aid) => acc + (proj.budgets[aid] || 0), 0);
                
                const ratio = empBudget > 0 ? (empProjHours / empBudget) * 100 : 0;
                return { name: getEmpName(emp, language), hours: empProjHours, ratio, posBudget: empBudget };
              }).filter(p => p.hours > 0).sort((a, b) => b.hours - a.hours);

              return (
                <div key={proj.id} className="dashboard-card p-6">
                  <div className="flex flex-col xl:flex-row xl:justify-between xl:items-start gap-6 mb-6">
                    <div>
                      <div className="flex flex-wrap items-center gap-3 mb-2">
                        <h4 className="text-2xl font-black text-slate-900">{proj.name}</h4>
                        <span className="dashboard-chip">
                          {isFullCycle ? '全周期' : `${dateRange.start} ~ ${dateRange.end}`}
                        </span>
                        {(isSuperAdmin || proj.managerId === state.currentUser?.id) && (
                          <div className="flex items-center gap-2">
                            <button 
                              onClick={() => {
                                const newState = { ...state, projects: state.projects.map(p => p.id === proj.id ? { ...p, visible: false } : p) };
                                setState(newState);
                                handleSaveToDatabase(newState);
                              }}
                              className="px-3 py-1 bg-slate-100 hover:bg-rose-50 text-slate-400 hover:text-rose-600 rounded-xl text-[10px] font-black uppercase tracking-widest transition-colors"
                            >
                              关闭项目
                            </button>
                            <input 
                              type="number" 
                              value={proj.sortOrder || 0}
                              onChange={e => {
                                const newState = { ...state, projects: state.projects.map(p => p.id === proj.id ? { ...p, sortOrder: parseInt(e.target.value) } : p) };
                                setState(newState);
                                handleSaveToDatabase(newState);
                              }}
                              className="w-12 bg-slate-50 border border-slate-200 rounded-xl text-[10px] font-black text-center"
                              title="排序值"
                            />
                          </div>
                        )}
                        <button
                          onClick={() => setFullCycleProjects(prev => ({ ...prev, [proj.id]: !prev[proj.id] }))}
                          className={cn(
                            "px-3 py-1 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all",
                            isFullCycle ? "bg-blue-600 text-white shadow-md shadow-blue-100" : "bg-blue-50 text-blue-600 hover:bg-blue-100"
                          )}
                        >
                          {isFullCycle ? '全周期模式' : '当前周期'}
                        </button>
                      </div>
                      <p className="text-slate-400 font-bold uppercase tracking-widest text-[10px]">
                        PM: {getEmpName(state.employees.find(e => e.id === proj.managerId) || null, language)}
                      </p>
                    </div>
                    <div className="flex flex-wrap gap-4 xl:gap-8">
                      <div className="text-right">
                        <p className="text-3xl font-black text-blue-600">{displayTotalHours.toFixed(1)}h</p>
                        <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                          {isFullCycle ? '全周期总计' : '区间总计'}
                        </p>
                      </div>
                      
                      <div className="flex gap-6 border-l border-slate-200 pl-6">
                        <div className="text-right">
                          <p className={cn("text-xl font-black", getXBColor(xbWC))}>
                            {allTimeWC.toFixed(0)} / {budgetWC.toFixed(0)}h
                          </p>
                          <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest">
                            WC消耗占比 ({xbWC.toFixed(1)}%)
                          </p>
                        </div>
                        <div className="text-right border-l border-slate-200 pl-6">
                          <p className={cn("text-xl font-black", getXBColor(xbBC))}>
                            {allTimeBC.toFixed(0)} / {budgetBC.toFixed(0)}h
                          </p>
                          <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest">
                            BC消耗占比 ({xbBC.toFixed(1)}%)
                          </p>
                        </div>
                      </div>

                      <div className="text-right border-l border-slate-200 pl-6">
                        <p className={cn("text-3xl font-black", getXBColor(totalSpanPercent))}>
                          {allTimeHours.toFixed(0)} / {totalBudget.toFixed(0)}h
                        </p>
                        <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                          总体跨度占比 ({totalSpanPercent.toFixed(1)}%)
                        </p>
                      </div>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    {/* Allocation Breakdown */}
                    <div className="space-y-6">
                      <h5 className="dashboard-kicker border-b border-slate-200 pb-2">
                        按分工
                      </h5>
                      <div className="space-y-4">
                        {state.allocations.map(alloc => {
                          const allocActual = displayEntries
                            .filter(e => e.allocationId === alloc.id)
                            .reduce((acc, e) => acc + e.hours, 0);
                          const budget = proj.budgets[alloc.id] || 0;
                          const percent = budget > 0 ? (allocActual / budget) * 100 : 0;
                          
                          if (allocActual === 0 && budget === 0) return null;

                          return (
                            <div key={alloc.id} className={cn("p-3 rounded-2xl border transition-all", budget === 0 && allocActual > 0 ? "bg-rose-50 border-rose-200" : "bg-slate-50/70 border-slate-200 space-y-2")}>
                              <div className="flex justify-between items-center">
                                <span className="text-sm font-bold text-slate-700">
                                  {language === 'zh' ? (alloc.nameZh || alloc.nameEn) : (alloc.nameEn || alloc.nameZh)}
                                  {budget === 0 && allocActual > 0 && (
                                    <span className="ml-2 text-[8px] font-black text-rose-600 bg-rose-100 px-1.5 py-0.5 rounded-full uppercase">
                                      无预算
                                    </span>
                                  )}
                                </span>
                                <div className="flex items-center gap-2">
                                  <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">XB:</span>
                                  <span className={cn("text-sm font-black", getXBColor(percent))}>
                                    {percent.toFixed(1)}%
                                  </span>
                                </div>
                              </div>
                              <div className="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden mt-2">
                                <div 
                                  className={cn("h-full transition-all duration-500", getXBColor(percent).replace('text-', 'bg-'))} 
                                  style={{ width: `${Math.min(percent, 100)}%` }} 
                                />
                              </div>
                              <p className="text-[9px] font-bold text-slate-400 text-right mt-1">
                                {allocActual.toFixed(1)} / {budget.toFixed(0)}h = {percent.toFixed(1)}%
                              </p>
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    {/* Person Breakdown */}
                    <div className="space-y-6">
                      <h5 className="dashboard-kicker border-b border-slate-200 pb-2">
                        按人员
                      </h5>
                      <div className="space-y-4 max-h-[300px] overflow-y-auto pr-2 custom-scrollbar">
                        {personBreakdown.map(person => (
                          <div key={person.name} className="flex justify-between items-center bg-slate-50/70 border border-slate-200 p-3 rounded-2xl">
                            <div>
                              <p className="text-sm font-black text-slate-900">{person.name}</p>
                              <p className="text-[10px] font-bold text-slate-400">{person.hours.toFixed(1)}h</p>
                            </div>
                            <div className="text-right">
                              <p className="text-sm font-black text-indigo-600">{person.ratio.toFixed(1)}%</p>
                              <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">
                                {person.hours.toFixed(1)} / {person.posBudget.toFixed(0)}h = {person.ratio.toFixed(1)}%
                              </p>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Time Trend */}
                    <div className="space-y-6">
                      <h5 className="dashboard-kicker border-b border-slate-200 pb-2">
                        投入趋势
                      </h5>
                      <div className="h-[200px]">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={(() => {
                            const chartStart = isFullCycle 
                              ? (displayEntries.length > 0 ? parseISO(displayEntries.reduce((min, e) => e.date < min ? e.date : min, displayEntries[0].date)) : start)
                              : start;
                            const chartEnd = isFullCycle 
                              ? (displayEntries.length > 0 ? parseISO(displayEntries.reduce((max, e) => e.date > max ? e.date : max, displayEntries[0].date)) : end)
                              : end;
                            
                            // Limit to last 30 days if full cycle is too long to avoid clutter
                            const allDays = eachDayOfInterval({ start: chartStart, end: chartEnd }).filter(d => !isWeekend(d));
                            const displayDays = allDays.length > 30 ? allDays.slice(-30) : allDays;

                            return displayDays.map(d => {
                              const dateStr = format(d, 'yyyy-MM-dd');
                              return {
                                date: format(d, 'MM/dd'),
                                hours: displayEntries.filter(e => e.date === dateStr).reduce((acc, e) => acc + e.hours, 0)
                              };
                            });
                          })()}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                            <XAxis dataKey="date" axisLine={false} tickLine={false} tick={{ fontSize: 10, fontWeight: 700, fill: '#94a3b8' }} />
                            <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fontWeight: 700, fill: '#94a3b8' }} />
                            <Tooltip 
                              contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)', fontWeight: 900 }}
                              cursor={{ fill: '#f8fafc' }}
                            />
                            <Bar dataKey="hours" fill="#3B82F6" radius={[4, 4, 0, 0]} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  </div>
                </div>
              );
            })}
        </div>
        {isSuperAdmin && state.projects.some(p => p.visible === false) && (
          <div className="mt-8 p-6 bg-slate-50 rounded-[2rem] border border-dashed border-slate-200 text-center">
            <p className="text-sm font-bold text-slate-400 mb-4">已关闭的项目</p>
            <div className="flex flex-wrap justify-center gap-2">
              {state.projects.filter(p => p.visible === false).map(p => (
                <button 
                  key={p.id}
                  onClick={() => {
                    const newState = { ...state, projects: state.projects.map(proj => proj.id === p.id ? { ...proj, visible: true } : proj) };
                    setState(newState);
                    handleSaveToDatabase(newState);
                  }}
                  className="px-4 py-2 bg-white border border-slate-200 rounded-xl text-xs font-black text-slate-600 hover:border-blue-500 hover:text-blue-600 transition-all"
                >
                  {p.name}
                </button>
              ))}
            </div>
          </div>
        )}
      </section>
    </motion.div>
  );
}

// --- Device BOM Center (Master BOM) ---
function DeviceBomCenterView() {
  const context = useContext(AppContext);
  if (!context) return null;
  const { state, setState, language, handleSaveToDatabase, addLog, setView } = context as any;
  const t = translations[language];

  const [selectedDeviceId, setSelectedDeviceId] = useState<string>(state.deviceModels?.[0]?.id || '');
  const [selectedBomType, setSelectedBomType] = useState<MaterialType>('Mechanical');
  const [selectedBomKind, setSelectedBomKind] = useState<'assembly' | 'purchasing'>('purchasing');
  const [selectedTemplateId, setSelectedTemplateId] = useState<string>('');
  const hotRef = useRef<any>(null);
  const importInputRef = useRef<HTMLInputElement | null>(null);
  const assemblyImportRef = useRef<HTMLInputElement | null>(null);
  const purchasingImportRef = useRef<HTMLInputElement | null>(null);
  const [lastRecon, setLastRecon] = useState<{ ok: boolean; diffs: any[]; summary: any } | null>(null);

  // 设备型号下拉：默认只显示“名称”；仅当名称重复才显示 (ID)
  const deviceOptions = useMemo(() => {
    const all = (state.deviceModels || []) as DeviceModel[];
    return all.map((d: DeviceModel) => ({ id: d.id, name: formatDeviceModelLabel(d, all) }));
  }, [state.deviceModels]);

  const renameDeviceModel = async () => {
    if (!selectedDeviceId) {
      alert(language === 'zh' ? '请先选择设备型号' : 'Please select a device model');
      return;
    }
    const dm = (state.deviceModels || []).find((d: DeviceModel) => d.id === selectedDeviceId);
    if (!dm) return;
    const nextName = (window.prompt(language === 'zh' ? '请输入新的设备型号名称' : 'Enter new device model name', dm.name) || '').trim();
    if (!nextName || nextName === dm.name) return;
    const now = new Date().toISOString();
    const nextModels = (state.deviceModels || []).map((d: DeviceModel) =>
      d.id === selectedDeviceId ? { ...d, name: nextName, updatedAt: now } : d
    );
    const newState = { ...state, deviceModels: nextModels };
    setState(newState);
    await handleSaveToDatabase(newState);
    addLog?.('device_model', `重命名设备型号：${selectedDeviceId} → ${nextName}`);
    alert(language === 'zh' ? '设备型号名称已更新' : 'Device model renamed');
  };

  const renameCurrentTemplate = async () => {
    if (!selectedTemplateId) {
      alert(language === 'zh' ? '请先选择一个BOM版本' : 'Please select a BOM version');
      return;
    }
    const tpl = (state.deviceBomTemplates || []).find((x: DeviceBOMTemplate) => x.id === selectedTemplateId);
    if (!tpl) return;
    const nextName = (window.prompt(
      language === 'zh' ? '请输入新的BOM名称（会全局生效）' : 'Enter new BOM name (global)',
      tpl.name || ''
    ) || '').trim();
    if (nextName === String(tpl.name || '').trim()) return;
    const now = new Date().toISOString();
    const nextTemplates = (state.deviceBomTemplates || []).map((x: DeviceBOMTemplate) =>
      x.id === selectedTemplateId ? { ...x, name: nextName, updatedAt: now } : x
    );
    const newState = { ...state, deviceBomTemplates: nextTemplates };
    setState(newState);
    await handleSaveToDatabase(newState);
    addLog?.('device_bom', `重命名BOM版本：${tpl.deviceModelId} ${tpl.version} → ${nextName}`);
    alert(language === 'zh' ? 'BOM名称已更新（全局生效）' : 'BOM name updated (global)');
  };

  const templateOptions = useMemo(() => {
    if (!selectedDeviceId) return [];
    return (state.deviceBomTemplates || [])
      .filter((x: DeviceBOMTemplate) =>
        x.deviceModelId === selectedDeviceId
        && x.bomType === selectedBomType
        && (x.templateKind || 'purchasing') === selectedBomKind
      )
      .sort((a: DeviceBOMTemplate, b: DeviceBOMTemplate) => String(a.version).localeCompare(String(b.version)))
      .map((x: DeviceBOMTemplate) => {
        const base = `${x.version} · ${x.name || ''}`.trim();
        const importedAt = x.importedAt || x.createdAt;
        const changed = importedAt && x.updatedAt && x.updatedAt > importedAt;
        const hist = x.sourceHistoryLastDate ? ` · 历史:${x.sourceHistoryLastDate}` : '';
        const flag = changed ? ' · 已改动' : '';
        return { id: x.id, label: `${base}${hist}${flag}` };
      });
  }, [state.deviceBomTemplates, selectedDeviceId, selectedBomType, selectedBomKind]);

  useEffect(() => {
    if (!selectedDeviceId) {
      setSelectedTemplateId('');
      return;
    }
    if (templateOptions.length === 0) {
      setSelectedTemplateId('');
      return;
    }
    if (!templateOptions.some(o => o.id === selectedTemplateId)) {
      setSelectedTemplateId(templateOptions[0].id);
    }
  }, [selectedDeviceId, selectedBomType, templateOptions, selectedTemplateId]);

  const currentTemplate = useMemo(
    () => (state.deviceBomTemplates || []).find((x: DeviceBOMTemplate) => x.id === selectedTemplateId),
    [state.deviceBomTemplates, selectedTemplateId]
  );
  const currentLines = useMemo(
    () => (state.deviceBomLines || []).filter((l: DeviceBOMLine) => l.templateId === selectedTemplateId),
    [state.deviceBomLines, selectedTemplateId]
  );

  // --- Assembly tree view helpers ---
  const [expandedNodes, setExpandedNodes] = useState<Record<string, boolean>>({});

  const levelPathParts = (p: string) => p.split('.').map(x => Number(x));
  const compareLevelPath = (a: string, b: string) => {
    const pa = levelPathParts(a);
    const pb = levelPathParts(b);
    const n = Math.max(pa.length, pb.length);
    for (let i = 0; i < n; i++) {
      const va = pa[i] ?? -1;
      const vb = pb[i] ?? -1;
      if (va !== vb) return va - vb;
    }
    return 0;
  };

  const assemblyNodes = useMemo(() => {
    if (selectedBomKind !== 'assembly') return [] as DeviceBOMLine[];
    return currentLines
      .filter(l => !!l.levelPath)
      .slice()
      .sort((a, b) => compareLevelPath(String(a.levelPath), String(b.levelPath)));
  }, [currentLines, selectedBomKind]);

  const childCountMap = useMemo(() => {
    const map = new Map<string, number>();
    assemblyNodes.forEach(n => {
      const p = String(n.parentLevelPath || '');
      if (!p) return;
      map.set(p, (map.get(p) || 0) + 1);
    });
    return map;
  }, [assemblyNodes]);

  // default expand to level 2: expand all level-1 nodes so level-2 is visible
  useEffect(() => {
    if (selectedBomKind !== 'assembly') return;
    const next: Record<string, boolean> = {};
    assemblyNodes.forEach(n => {
      const path = String(n.levelPath || '');
      if (!path) return;
      const depth = path.split('.').length;
      if (depth === 1) next[path] = true;
    });
    setExpandedNodes(next);
  }, [selectedTemplateId, selectedBomKind]); // eslint-disable-line react-hooks/exhaustive-deps

  const visibleAssemblyNodes = useMemo(() => {
    const byPath = new Map<string, DeviceBOMLine>();
    assemblyNodes.forEach(n => byPath.set(String(n.levelPath), n));
    const visible: DeviceBOMLine[] = [];
    const isVisible = (n: DeviceBOMLine) => {
      const parent = String(n.parentLevelPath || '');
      if (!parent) return true;
      if (!expandedNodes[parent]) return false;
      const pNode = byPath.get(parent);
      return pNode ? isVisible(pNode) : true;
    };
    assemblyNodes.forEach(n => { if (isVisible(n)) visible.push(n); });
    return visible;
  }, [assemblyNodes, expandedNodes]);

  const tableData = useMemo(() => {
    return currentLines.map(l => ([
      l.partCodeModel,
      l.revision,
      l.name,
      l.specs,
      l.quantityPerDevice,
      l.unit,
      l.brand,
      l.supplier,
      l.remark
    ]));
  }, [currentLines]);

  const addDeviceModel = () => {
    const code = (window.prompt('请输入设备型号编码（如 DM-100）') || '').trim();
    if (!code) return;
    const exists = (state.deviceModels || []).some((d: DeviceModel) => d.id === code);
    if (exists) {
      alert('该设备型号编码已存在，请换一个编码');
      setSelectedDeviceId(code);
      return;
    }
    const name = (window.prompt('请输入设备型号名称/描述', code) || '').trim();
    const now = new Date().toISOString();
    const dm: DeviceModel = { id: code, name: name || code, createdAt: now, updatedAt: now };
    const newState = { ...state, deviceModels: [dm, ...(state.deviceModels || [])] };
    setState(newState);
    handleSaveToDatabase(newState);
    setSelectedDeviceId(dm.id);
  };

  const ensureDevice = () => {
    if ((state.deviceModels || []).length > 0) return;
    addDeviceModel();
  };

  const createTemplate = (copyFromId?: string) => {
    if (!selectedDeviceId) {
      ensureDevice();
      return;
    }
    const defaultVer = 'V1.0';
    const version = (window.prompt('请输入BOM版本号（如 V1.0）', defaultVer) || '').trim();
    if (!version) return;
    const name = (window.prompt('请输入BOM名称（可空）', `${selectedBomType === 'Mechanical' ? '机械' : selectedBomType === 'Electrical' ? '电气' : 'BOM'} ${version}`) || '').trim();
    const now = new Date().toISOString();
    const tpl: DeviceBOMTemplate = {
      id: `dbt-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      deviceModelId: selectedDeviceId,
      bomType: selectedBomType,
      templateKind: selectedBomKind,
      version,
      name,
      createdAt: now,
      updatedAt: now
    };

    let newLines: DeviceBOMLine[] = [];
    if (copyFromId) {
      const src = (state.deviceBomLines || []).filter((l: DeviceBOMLine) => l.templateId === copyFromId);
      newLines = src.map((l: DeviceBOMLine) => ({
        ...l,
        id: `dbl-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        templateId: tpl.id
      }));
    }

    const newState = {
      ...state,
      deviceBomTemplates: [tpl, ...(state.deviceBomTemplates || [])],
      deviceBomLines: [...newLines, ...(state.deviceBomLines || [])],
      changeRecords: [
        {
          id: `chg-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
          templateId: tpl.id,
          deviceModelId: selectedDeviceId,
          bomType: selectedBomType,
          templateKind: selectedBomKind,
          oldVersion: copyFromId ? ((state.deviceBomTemplates || []).find((x: DeviceBOMTemplate) => x.id === copyFromId)?.version || '') : '',
          newVersion: version,
          diffType: 'new_version' as const,
          details: copyFromId ? `由现有版本复制生成：${version}` : `新建版本：${version}`,
          operatorId: state.currentUser?.id || '',
          createdAt: now
        },
        ...(state.changeRecords || [])
      ]
    };
    setState(newState);
    handleSaveToDatabase(newState);
    setSelectedTemplateId(tpl.id);
  };

  const guessFromFileName = (name: string) => {
    const base = name.replace(/\.[^.]+$/, '');
    const lower = base.toLowerCase();
    const bomType: MaterialType =
      (lower.includes('electrical') || base.includes('电气'))
        ? 'Electrical'
        : 'Mechanical';
    const versionMatch = base.match(/V(\d+(\.\d+)*)/i) || base.match(/BOMV(\d+(\.\d+)*)/i);
    const version = versionMatch ? `V${versionMatch[1]}` : 'V1.0';
    const displayName = base;
    return { bomType, version, displayName };
  };

  const parseQtyWithUnit = (raw: any, fallbackUnit: string) => {
    const s = String(raw ?? '').trim();
    if (!s) return { qty: 0, unit: fallbackUnit || '' };
    // e.g. "5950.000 mm" / "1" / "3.062 kg"
    const m = s.match(/^\s*([+-]?\d+(?:\.\d+)?)\s*([a-zA-Z]+)?\s*$/);
    if (m) {
      const qty = Number(m[1]);
      const unit = (m[2] || fallbackUnit || '').trim();
      return { qty: Number.isFinite(qty) ? qty : 0, unit };
    }
    const n = Number(s.replace(/,/g, ''));
    return { qty: Number.isFinite(n) ? n : 0, unit: fallbackUnit || '' };
  };

  const parseLevelPath = (v: any) => String(v ?? '').trim();

  const computeParentPath = (path: string) => {
    if (!path) return '';
    const parts = path.split('.').filter(Boolean);
    if (parts.length <= 1) return '';
    return parts.slice(0, -1).join('.');
  };

  const buildRecon = (assemblyLines: DeviceBOMLine[], purchasingLines: DeviceBOMLine[]) => {
    const sumByPart = (lines: DeviceBOMLine[]) => {
      const map = new Map<string, number>();
      lines.forEach(l => {
        const k = String(l.partCodeModel || '').trim();
        if (!k) return;
        const q = Number(l.quantityPerDevice) || 0;
        if (q <= 0) return;
        map.set(k, (map.get(k) || 0) + q);
      });
      return map;
    };

    // Assembly is a tree (raw nodes). We must expand it to purchasable leaf parts and aggregate by part code.
    const expandAssemblyAgg = (nodes: DeviceBOMLine[]) => {
      const items = (nodes || []).filter(n => n.levelPath && n.partCodeModel && (Number(n.quantityPerDevice) || 0) > 0);
      const byPath = new Map<string, DeviceBOMLine>();
      items.forEach(n => byPath.set(String(n.levelPath), n));
      const hasChild = new Set<string>();
      items.forEach(n => { if (n.parentLevelPath) hasChild.add(String(n.parentLevelPath)); });
      const cum = new Map<string, number>();
      const getCum = (p: string): number => {
        if (cum.has(p)) return cum.get(p)!;
        const node = byPath.get(p);
        if (!node) return 1;
        const parent = String(node.parentLevelPath || '');
        const parentCum = parent ? getCum(parent) : 1;
        const local = Number(node.quantityPerDevice) || 0;
        const v = parentCum * local;
        cum.set(p, v);
        return v;
      };
      items.forEach(n => getCum(String(n.levelPath)));
      const isLeaf = (n: DeviceBOMLine) => {
        const fn = String((n.fileName || '')).toLowerCase();
        if (fn.endsWith('.ipt')) return true;
        return !hasChild.has(String(n.levelPath));
      };
      const agg = new Map<string, number>();
      items.forEach(n => {
        if (!isLeaf(n)) return;
        const part = String(n.partCodeModel || '').trim();
        if (!part) return;
        const q = cum.get(String(n.levelPath)) || 0;
        if (q <= 0) return;
        agg.set(part, (agg.get(part) || 0) + q);
      });
      return agg;
    };

    const asm = expandAssemblyAgg(assemblyLines);
    const pur = sumByPart(purchasingLines);
    const keys = new Set<string>([...asm.keys(), ...pur.keys()]);
    const diffs: any[] = [];
    keys.forEach(k => {
      const a = asm.get(k) || 0;
      const p = pur.get(k) || 0;
      const delta = a - p;
      if (Math.abs(delta) > 1e-9) {
        diffs.push({ part: k, assemblyQty: a, purchasingQty: p, delta });
      }
    });
    diffs.sort((x, y) => Math.abs(y.delta) - Math.abs(x.delta));
    return {
      ok: diffs.length === 0,
      diffs,
      summary: {
        assemblyItems: asm.size,
        purchasingItems: pur.size,
        diffCount: diffs.length
      }
    };
  };

  const exportReconToXlsx = (recon: any) => {
    try {
      const wb = XLSX.utils.book_new();
      const rows = [
        ['零件代号/型号', '装配展开汇总数量', '采购BOM数量', '差异(装配-采购)'],
        ...recon.diffs.map((d: any) => [d.part, d.assemblyQty, d.purchasingQty, d.delta])
      ];
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, '对账差异');
      XLSX.writeFile(wb, `BOM对账差异_${selectedDeviceId}_${new Date().toISOString().slice(0,10)}.xlsx`);
    } catch (e) {
      alert('导出对账报告失败');
    }
  };

  const importSpecificBom = (kind: 'assembly' | 'purchasing') => async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!selectedDeviceId) {
      alert('请先选择设备型号（device model id）');
      return;
    }
    const { bomType: guessedType, version: guessedVersion, displayName } = guessFromFileName(file.name);
    const version = (window.prompt('请输入BOM版本号（将用于本次导入）', guessedVersion) || '').trim();
    if (!version) return;

    const now = new Date().toISOString();
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });

        // Parse optional history sheet (e.g. "历史" / "History")
        const historySheetName =
          wb.SheetNames.find(n => String(n).includes('历史')) ||
          wb.SheetNames.find(n => String(n).toLowerCase().includes('history'));
        let historyMeta: Partial<DeviceBOMTemplate> = {};
        try {
          if (historySheetName) {
            const hws = wb.Sheets[historySheetName];
            const hrows = XLSX.utils.sheet_to_json(hws, { defval: '', raw: false }) as any[];
            const cleaned = (hrows || []).filter(r => Object.values(r || {}).some(v => String(v || '').trim() !== ''));
            const last = cleaned[cleaned.length - 1] || {};
            historyMeta = {
              sourceHistoryCount: cleaned.length || 0,
              sourceHistoryLastDate: String(last['日期'] || last['Date'] || '').trim(),
              sourceHistoryLastVersion: String(last['BOM版本'] || last['Version'] || last['版本'] || '').trim(),
              sourceHistoryLastBy: String(last['更改人'] || last['Changed by'] || last['By'] || '').trim(),
            };
          }
        } catch (e) {
          // ignore history parse errors
        }

        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false }) as any[];
        if (!data.length) return alert('Excel为空或无法识别');

        // Create template & overwrite existing of same kind (per your rule)
        const tplId = `dbt-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
        const tpl: DeviceBOMTemplate = {
          id: tplId,
          deviceModelId: selectedDeviceId,
          bomType: guessedType,
          templateKind: kind,
          version,
          name: displayName,
          createdAt: now,
          updatedAt: now,
          importedAt: now,
          sourceFileName: file.name,
          ...historyMeta
        };

        const norm = (v: any) => String(v ?? '').trim();
        const pick = (row: any, keys: string[]) => {
          for (const k of keys) {
            const v = row?.[k];
            if (v !== undefined && v !== null && String(v).trim() !== '') return v;
          }
          return '';
        };

        let lines: DeviceBOMLine[] = [];
        if (kind === 'assembly') {
          // Build hierarchy first.
          // Supports:
          // 1) 机械表：项目=1/1.1/1.1.1 (level path)
          // 2) 电气表：Level=1/2/5... (depth outline) + Component qty/unit + Model
          const parsed = data.map((row) => {
            const levelRaw = pick(row, ['项目', 'Level', '层级', 'BOM表结构']);
            const levelPath = parseLevelPath(levelRaw);
            const part = norm(pick(row, [
              '零件代号', '零件代号/型号',
              'Model', '型号',
              'Item internal #', 'Item internal#',
              'SAP_Drawingno', 'SAP_DrawingNo'
            ]));
            const qtyRaw = pick(row, ['数量', 'Qty', 'Quantity', 'Component qty']);
            const unitFallback = norm(pick(row, ['单位', 'Unit', 'Component unit']));
            let { qty, unit } = parseQtyWithUnit(qtyRaw, unitFallback);
            // assemblies in some electrical BOM omit qty; treat as 1 for structure nodes
            if (!String(qtyRaw ?? '').trim()) qty = 1;
            const fileName = norm(pick(row, ['文件名称', 'File name']));
            const nodeType = norm(pick(row, ['BOM 表结构', 'BOM表结构', 'Item description']));
            const revision = norm(pick(row, ['版本', 'SAP_Revision', 'Revision', 'Rev'])) || 'NA';
            const desc = norm(pick(row, ['描述', 'Descrption', 'SAP_Description', 'Description', 'Item description', 'Item description '])) || part;
            const material = norm(pick(row, ['材料', 'SAP_Material']));
            const mass = norm(pick(row, ['质量']));
            const supplier = norm(pick(row, ['供应商', 'Supplier']));
            const costCenter = norm(pick(row, ['成本中心', 'Cost Center']));
            const stockNo = norm(pick(row, ['库存编号', 'Stock No', 'Stock']));
            const webUrl = norm(pick(row, ['Web 链接', 'Web link', 'Link']));

            return {
              levelRaw,
              levelPath,
              parentLevelPath: computeParentPath(levelPath),
              nodeType,
              part,
              qty,
              unit,
              fileName,
              revision,
              desc,
              material,
              mass,
              supplier,
              costCenter,
              stockNo,
              webUrl
            };
          }).filter(n => n.part);

          // Detect depth-outline mode: levelPath is a pure number (no ".") for most rows, e.g. Level=1/2/5
          const depthMode = parsed.filter(n => String(n.levelPath || '').trim()).length > 0
            && parsed.filter(n => String(n.levelPath || '').includes('.')).length === 0
            && parsed.filter(n => /^\d+$/.test(String(n.levelPath || ''))).length >= Math.max(5, Math.floor(parsed.length * 0.5));

          const rawNodes = (() => {
            if (!depthMode) return parsed;
            const counters: number[] = [];
            return parsed.map(n => {
              const depth = Math.max(1, Math.min(12, Number(String(n.levelPath || '').trim()) || 1));
              while (counters.length < depth) counters.push(0);
              counters.length = depth;
              counters[depth - 1] = (counters[depth - 1] || 0) + 1;
              // reset deeper (already truncated)
              const lp = counters.join('.');
              return { ...n, levelPath: lp, parentLevelPath: counters.slice(0, -1).join('.') };
            });
          })();

          // Store raw nodes as-is (tree view). Cumulative expansion is computed on demand (for recon / forecasting).
          lines = rawNodes.map(n => ({
            id: `dbl-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
            templateId: tplId,
            partCodeModel: n.part,
            revision: n.revision || 'NA',
            name: n.desc || n.part,
            specs: [n.material ? `材料:${n.material}` : '', n.mass ? `质量:${n.mass}` : ''].filter(Boolean).join('；'),
            quantityPerDevice: Number(n.qty) || 0,
            unit: n.unit || '',
            brand: '',
            supplier: n.supplier || '',
            remark: [n.nodeType ? `结构:${n.nodeType}` : '', n.fileName ? `文件:${n.fileName}` : ''].filter(Boolean).join('；'),
            levelPath: n.levelPath,
            parentLevelPath: n.parentLevelPath,
            nodeType: n.nodeType,
            fileName: n.fileName,
            meta: { costCenter: n.costCenter, stockNo: n.stockNo, webUrl: n.webUrl, levelRaw: n.levelRaw }
          })).filter(l => String(l.partCodeModel || '').trim());
        } else {
          // purchasing list: treat as already aggregated per device, but still sum duplicate part codes
          const tmp: DeviceBOMLine[] = [];
          data.forEach((row) => {
            const part = norm(pick(row, [
              '零件代号', '零件代号/型号',
              'Model', '型号',
              'Item internal #', 'Item internal#',
              'SAP_Drawingno', 'SAP_DrawingNo'
            ]));
            const qtyRaw = pick(row, ['数量', 'Qty', 'Quantity', 'Component qty']);
            const unitFallback = norm(pick(row, ['单位', 'Unit', 'Component unit']));
            const { qty, unit } = parseQtyWithUnit(qtyRaw, unitFallback);
            if (!part || qty <= 0) return;
            const revision = norm(pick(row, ['版本', 'SAP_Revision', 'Revision', 'Rev'])) || 'NA';
            const desc = norm(pick(row, ['描述', 'Descrption', 'SAP_Description', 'Description', 'Item description', 'Item description '])) || part;
            const material = norm(pick(row, ['材料', 'SAP_Material']));
            const mass = norm(pick(row, ['质量']));
            const supplier = norm(pick(row, ['供应商', 'Supplier']));
            const fileName = norm(pick(row, ['文件名称', 'File name']));
            const nodeType = norm(pick(row, ['BOM 表结构', 'BOM表结构', 'Item description']));
            tmp.push({
              id: `dbl-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
              templateId: tplId,
              partCodeModel: part,
              revision,
              name: desc,
              specs: [material ? `材料:${material}` : '', mass ? `质量:${mass}` : ''].filter(Boolean).join('；'),
              quantityPerDevice: qty,
              unit: unit || '',
              brand: '',
              supplier,
              remark: [nodeType ? `结构:${nodeType}` : '', fileName ? `文件:${fileName}` : ''].filter(Boolean).join('；'),
            });
          });
          const agg = new Map<string, DeviceBOMLine>();
          tmp.forEach(l => {
            const k = l.partCodeModel;
            if (!agg.has(k)) agg.set(k, { ...l });
            else agg.get(k)!.quantityPerDevice = (Number(agg.get(k)!.quantityPerDevice) || 0) + (Number(l.quantityPerDevice) || 0);
          });
          lines = Array.from(agg.values());
        }

        if (!lines.length) return alert('未解析到有效BOM行（请检查列名：项目/数量/零件代号/描述）');

        // Overwrite: remove existing templates & lines for this model+type+kind
        const oldTplIds = new Set(
          (state.deviceBomTemplates || [])
            .filter((x: DeviceBOMTemplate) => x.deviceModelId === selectedDeviceId && x.bomType === guessedType && (x.templateKind || 'purchasing') === kind)
            .map((x: DeviceBOMTemplate) => x.id)
        );

        const nextTemplates = [
          tpl,
          ...(state.deviceBomTemplates || []).filter((x: DeviceBOMTemplate) => !oldTplIds.has(x.id))
        ];
        const nextLines = [
          ...(state.deviceBomLines || []).filter((l: DeviceBOMLine) => !oldTplIds.has(l.templateId)),
          ...lines
        ];

        const newState: AppState = { ...state, deviceBomTemplates: nextTemplates, deviceBomLines: nextLines };
        setState(newState);
        await handleSaveToDatabase(newState);

        setSelectedBomType(guessedType);
        setSelectedBomKind(kind);
        setSelectedTemplateId(tplId);

        // If both kinds exist, auto-reconcile
        const asmTpl = nextTemplates.find((x: DeviceBOMTemplate) => x.deviceModelId === selectedDeviceId && x.bomType === guessedType && (x.templateKind || 'purchasing') === 'assembly');
        const purTpl = nextTemplates.find((x: DeviceBOMTemplate) => x.deviceModelId === selectedDeviceId && x.bomType === guessedType && (x.templateKind || 'purchasing') === 'purchasing');
        if (asmTpl && purTpl) {
          const asmLines = nextLines.filter((l: DeviceBOMLine) => l.templateId === asmTpl.id);
          const purLines = nextLines.filter((l: DeviceBOMLine) => l.templateId === purTpl.id);
          const recon = buildRecon(asmLines, purLines);
          setLastRecon(recon);
          alert(recon.ok ? '导入完成：对账一致' : `导入完成：对账发现 ${recon.summary.diffCount} 处差异（可在页面查看/导出）`);
        } else {
          setLastRecon(null);
          alert('导入完成');
        }
      } catch (err) {
        console.error(err);
        alert(`导入失败：${(err as any)?.message || err}`);
      } finally {
        if (kind === 'assembly' && assemblyImportRef.current) assemblyImportRef.current.value = '';
        if (kind === 'purchasing' && purchasingImportRef.current) purchasingImportRef.current.value = '';
      }
    };
    reader.readAsBinaryString(file);
  };

  const importDeviceBomFromExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const { bomType: guessedType, version: guessedVersion, displayName } = guessFromFileName(file.name);

    const ensureTemplateForImport = async (): Promise<{ templateId: string; bomType: MaterialType }> => {
      // If already selected, import into current template
      if (selectedTemplateId) return { templateId: selectedTemplateId, bomType: selectedBomType };

      // Otherwise, create device model + template quickly based on filename
      const dmId = (window.prompt('未选择BOM版本。请输入设备型号编码（主键，如 DM-100）') || '').trim();
      if (!dmId) throw new Error('cancel');
      const dmName = (window.prompt('请输入设备型号名称/描述', displayName) || '').trim();
      const now = new Date().toISOString();
      const dm: DeviceModel = { id: dmId, name: dmName || dmId, createdAt: now, updatedAt: now };

      const tpl: DeviceBOMTemplate = {
        id: `dbt-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        deviceModelId: dmId,
        bomType: guessedType,
        version: guessedVersion,
        name: displayName,
        createdAt: now,
        updatedAt: now
      };

      const nextState: AppState = {
        ...state,
        deviceModels: [
          ...(state.deviceModels || []).filter((x: DeviceModel) => x.id !== dmId),
          dm,
          ...(state.deviceModels || []).filter((x: DeviceModel) => x.id === dmId)
        ],
        deviceBomTemplates: [tpl, ...(state.deviceBomTemplates || [])],
        deviceBomLines: [...(state.deviceBomLines || [])]
      };
      setState(nextState);
      await handleSaveToDatabase(nextState);
      setSelectedDeviceId(dmId);
      setSelectedBomType(guessedType);
      setSelectedTemplateId(tpl.id);
      return { templateId: tpl.id, bomType: guessedType };
    };

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const { templateId, bomType } = await ensureTemplateForImport();
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false }) as any[];
        if (!data.length) {
          alert('Excel为空或无法识别');
          return;
        }

        const hasKey = (k: string) => Object.prototype.hasOwnProperty.call(data[0], k);
        const mode: 'mech_sap' | 'electrical' = hasKey('零件代号') || hasKey('BOM 表结构') || hasKey('SAP_Drawingno')
          ? 'mech_sap'
          : 'electrical';

        const pick = (row: any, keys: string[]) => {
          for (const k of keys) {
            const v = row?.[k];
            if (v !== undefined && v !== null && String(v).trim() !== '') return v;
          }
          return '';
        };
        const norm = (v: any) => String(v ?? '').trim();
        const normRev = (v: any) => {
          const r = norm(v);
          if (!r || r === '-' || r === 'N/A') return 'NA';
          return r;
        };
        const toQty = (v: any) => {
          const n = Number(String(v ?? '').replace(/,/g, '').trim());
          return Number.isFinite(n) ? n : 0;
        };

        const parsedLines: DeviceBOMLine[] = [];
        data.forEach((row, idx) => {
          if (mode === 'mech_sap') {
            const part = norm(pick(row, ['零件代号', 'SAP_Drawingno', 'SAP_DrawingNo', '零件代号/型号', '零件代号']));
            const qty = toQty(pick(row, ['数量', 'Qty', 'Quantity']));
            if (!part || qty <= 0) return;
            const rev = normRev(pick(row, ['版本', 'SAP_Revision', 'Revision', 'Rev']));
            const name = norm(pick(row, ['描述', 'SAP_Description', 'Item description', 'Item description ']));
            const material = norm(pick(row, ['材料', 'SAP_Material']));
            const weight = norm(pick(row, ['质量']));
            const specs = [material ? `材料:${material}` : '', weight ? `质量:${weight}` : ''].filter(Boolean).join('；');
            const supplier = norm(pick(row, ['供应商', 'Supplier']));
            const fileName = norm(pick(row, ['文件名称']));
            const bomStruct = norm(pick(row, ['BOM 表结构']));
            const remark = [bomStruct ? `结构:${bomStruct}` : '', fileName ? `文件:${fileName}` : ''].filter(Boolean).join('；');
            parsedLines.push({
              id: `dbl-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
              templateId,
              partCodeModel: part,
              revision: rev,
              name,
              specs,
              quantityPerDevice: qty,
              unit: '',
              brand: '',
              supplier,
              remark
            });
            return;
          }

          // electrical
          const part = norm(pick(row, ['Model', '零件代号/型号', '零件代号', 'Item internal #', 'Item internal#']));
          const qty = toQty(pick(row, ['Component qty', '数量', 'Qty']));
          // Many header/section rows have qty empty; ignore them.
          if (!part || qty <= 0) return;
          const unit = norm(pick(row, ['Component unit', '单位', 'Unit']));
          const name = norm(pick(row, ['描述', 'Item description', 'Item description ', 'Item description']));
          const supplier = norm(pick(row, ['Supplier', '供应商']));
          const remark = norm(pick(row, ['备注']));
          parsedLines.push({
            id: `dbl-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
            templateId,
            partCodeModel: part,
            revision: 'NA',
            name,
            specs: '',
            quantityPerDevice: qty,
            unit,
            brand: '',
            supplier,
            remark
          });
        });

        if (!parsedLines.length) {
          alert('未解析到有效BOM行（请检查是否存在“数量/Component qty”等列）');
          return;
        }

        const overwrite = window.confirm(
          `已识别 ${parsedLines.length} 行BOM明细。\n\n确定覆盖当前版本的明细吗？\n- 确定：覆盖\n- 取消：追加`
        );

        const now = new Date().toISOString();
        const newState: AppState = {
          ...state,
          deviceBomLines: overwrite
            ? [
                ...(state.deviceBomLines || []).filter((l: DeviceBOMLine) => l.templateId !== templateId),
                ...parsedLines
              ]
            : [...parsedLines, ...(state.deviceBomLines || [])],
          changeRecords: [
            {
              id: `chg-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
              templateId,
              deviceModelId: selectedDeviceId,
              bomType,
              templateKind: 'assembly',
              newVersion: ((state.deviceBomTemplates || []).find((x: DeviceBOMTemplate) => x.id === templateId)?.version || ''),
              diffType: 'manual_edit',
              details: `导入文件 ${file.name}（${overwrite ? '覆盖' : '追加'} ${parsedLines.length} 行）`,
              operatorId: state.currentUser?.id || '',
              createdAt: now
            },
            ...(state.changeRecords || [])
          ],
          deviceBomTemplates: (state.deviceBomTemplates || []).map((x: DeviceBOMTemplate) =>
            x.id === templateId ? { ...x, bomType, updatedAt: now, sourceFileName: file.name, importedAt: x.importedAt || x.createdAt } : x
          )
        };
        setState(newState);
        await handleSaveToDatabase(newState);

        addLog?.('device_bom', `导入设备BOM：${file.name} → ${templateId} (${overwrite ? '覆盖' : '追加'})`);
        alert('导入完成（已写入设备BOM中心）');
      } catch (err) {
        if ((err as any)?.message !== 'cancel') console.error(err);
      } finally {
        if (importInputRef.current) importInputRef.current.value = '';
      }
    };
    reader.readAsBinaryString(file);
  };

  const saveLines = async () => {
    if (!selectedTemplateId) {
      alert('请先选择或创建一个BOM版本');
      return;
    }
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;
    const rows = hot.getData() as any[][];
    const now = new Date().toISOString();
    const nextLines: DeviceBOMLine[] = [];
    rows.forEach((r) => {
      const part = String(r?.[0] ?? '').trim();
      const qty = Number(r?.[4] ?? 0);
      if (!part) return;
      if (!Number.isFinite(qty) || qty <= 0) return;
      nextLines.push({
        id: `dbl-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        templateId: selectedTemplateId,
        partCodeModel: part,
        revision: String(r?.[1] ?? '').trim() || 'NA',
        name: String(r?.[2] ?? '').trim(),
        specs: String(r?.[3] ?? '').trim(),
        quantityPerDevice: qty,
        unit: String(r?.[5] ?? '').trim(),
        brand: String(r?.[6] ?? '').trim(),
        supplier: String(r?.[7] ?? '').trim(),
        remark: String(r?.[8] ?? '').trim(),
      });
    });

    const newState: AppState = {
      ...state,
      deviceBomLines: [
        ...(state.deviceBomLines || []).filter((l: DeviceBOMLine) => l.templateId !== selectedTemplateId),
        ...nextLines
      ],
      changeRecords: [
        {
          id: `chg-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
          templateId: selectedTemplateId,
          deviceModelId: selectedDeviceId,
          bomType: selectedBomType,
          templateKind: selectedBomKind,
          newVersion: currentTemplate?.version || '',
          diffType: 'manual_edit',
          details: `手工保存BOM明细 ${nextLines.length} 行`,
          operatorId: state.currentUser?.id || '',
          createdAt: now
        },
        ...(state.changeRecords || [])
      ],
      deviceBomTemplates: (state.deviceBomTemplates || []).map((x: DeviceBOMTemplate) =>
        x.id === selectedTemplateId ? { ...x, updatedAt: now } : x
      )
    };
    setState(newState);
    await handleSaveToDatabase(newState);
    addLog?.('device_bom', `保存母版BOM：${selectedDeviceId} ${currentTemplate?.version || ''} (${selectedBomType})`);
    alert(language === 'zh' ? '已保存' : 'Saved');
  };

  // --- 引用生成需求：选择目标项目，然后把“当前设备型号”的全部【采购BOM（列表）】复制到该项目下 ---
  const [showGenerateModal, setShowGenerateModal] = useState(false);
  const projectOptionsForGen = useMemo(
    () => (state.projects || []).filter(p => p.id !== 'p-general' && p.id !== 'p-leave'),
    [state.projects]
  );
  const [genProjectId, setGenProjectId] = useState('');
  const [genDeviceCount, setGenDeviceCount] = useState(1);

  useEffect(() => {
    if (!showGenerateModal) return;
    if (!genProjectId && projectOptionsForGen.length > 0) setGenProjectId(projectOptionsForGen[0].id);
  }, [showGenerateModal, genProjectId, projectOptionsForGen]);

  const generateAllPurchasingBomsToProject = async () => {
    if (!selectedDeviceId) {
      alert(language === 'zh' ? '请先选择设备型号' : 'Please select device model');
      return;
    }
    if (!genProjectId) {
      alert(language === 'zh' ? '请选择目标项目' : 'Please select target project');
      return;
    }
    const deviceCount = Math.max(1, Math.floor(Number(genDeviceCount) || 1));

    const purchasingTemplates = (state.deviceBomTemplates || []).filter((t: any) =>
      t.deviceModelId === selectedDeviceId && ((t.templateKind || 'purchasing') === 'purchasing')
    ) as DeviceBOMTemplate[];
    if (purchasingTemplates.length === 0) {
      alert(language === 'zh' ? '该设备型号暂无“采购BOM（列表）”版本。' : 'No purchasing BOM templates found.');
      return;
    }

    const ok = window.confirm(language === 'zh'
      ? `确认引用生成需求并复制BOM？\n\n- 目标项目：${genProjectId}\n- 设备型号：${selectedDeviceId}\n- 采购BOM版本数：${purchasingTemplates.length}\n- 台数：${deviceCount}\n\n说明：将为每个采购BOM版本创建/更新一个项目下BOM清单，并覆盖该清单下的采购行。`
      : `Generate & copy all purchasing BOMs to project?\nProject: ${genProjectId}\nDevice: ${selectedDeviceId}\nTemplates: ${purchasingTemplates.length}\nCount: ${deviceCount}\n\nThis will overwrite rows under each generated BOM list.`);
    if (!ok) return;

    const nowIso = new Date().toISOString();
    const today = format(new Date(), 'yyyy-MM-dd');

    let keptReqs = [...(state.materialRequirements || [])] as any[];
    const nextBoms = [...(state.materialBoms || [])] as any[];
    const allNewReqs: MaterialRequirement[] = [];

    const upsertBom = (bomId: string, type: MaterialType, tpl: DeviceBOMTemplate) => {
      const idx = nextBoms.findIndex((b: any) => b.id === bomId && b.projectId === genProjectId && b.type === type);
      const name = `${tpl.version || ''}${tpl.name ? ` · ${tpl.name}` : ''}`.trim() || (language === 'zh' ? '默认BOM' : 'Default BOM');
      const now = nowIso;
      const payload = {
        id: bomId,
        projectId: genProjectId,
        type,
        name,
        createdAt: (idx >= 0 ? (nextBoms[idx] as any).createdAt : now) || now,
        importedFromDeviceModelId: selectedDeviceId,
        importedFromTemplateId: tpl.id,
        importedFromTemplateVersion: tpl.version,
        importedFromTemplateName: tpl.name,
        importedAt: now,
      };
      if (idx >= 0) nextBoms[idx] = { ...(nextBoms[idx] || {}), ...payload };
      else nextBoms.unshift(payload);
    };

    purchasingTemplates.forEach((tpl) => {
      const bomId = `bom-dev-${tpl.id}`;
      const type = tpl.bomType as MaterialType;
      // overwrite rows under this project+type+bomId
      keptReqs = keptReqs.filter((r: any) => !(r.projectId === genProjectId && r.type === type && ((r.bomId || 'bom-default') === bomId)));
      upsertBom(bomId, type, tpl);

      const lines = (state.deviceBomLines || []).filter((l: any) => l.templateId === tpl.id) as DeviceBOMLine[];
      lines.forEach((l, idx) => {
        const reqQty = (Number(l.quantityPerDevice) || 0) * deviceCount;
        const tag = `From DeviceBOM ${tpl.deviceModelId} ${tpl.bomType} ${tpl.version} ×${deviceCount}`;
        const req: MaterialRequirement = {
          id: `req-dev-${tpl.id}-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 7)}`,
          projectId: genProjectId,
          bomId,
          creatorId: state.currentUser?.id || 'system',
          type,
          code: `BOM-${Math.floor(Math.random() * 100000)}`,
          investmentOrder: '',
          csOrder: '',
          costCenter: '',
          name: l.name || l.partCodeModel,
          model: l.partCodeModel,
          drawingNo: String((l as any).drawingNo || ''),
          revision: String((l as any).revision || 'NA'),
          brand: String((l as any).brand || ''),
          specs: String((l as any).specs || ''),
          quantity: reqQty,
          unit: String((l as any).unit || ''),
          unitPrice: 0,
          totalPrice: 0,
          supplier: String((l as any).supplier || ''),
          prNumber: '',
          prCreatedAt: '',
          poNumber: '',
          poCreatedAt: '',
          actualSupplier: '',
          prStatus: '',
          needByDate: '',
          deliveryTime: '',
          expectedArrival: today,
          receivedAt: '',
          warehouseStatus: '',
          repairArrivalAt: '',
          inboundQuantity: 0,
          outboundQuantity: 0,
          inventoryQuantity: 0,
          userName: '',
          useDate: '',
          qualityFeedback: '',
          sourceTemplateVersion: `${tpl.version || ''}${tpl.name ? ` · ${tpl.name}` : ''}`.trim(),
          changeReason: '由设备BOM引用生成',
          changeReasonNote: tag,
          urgency: 'Medium',
          status: 'Pending Review',
          comments: tag,
          version: 1,
          createdAt: nowIso,
          updatedAt: nowIso,
          customFields: {},
        } as any;
        if (String((req as any).model || '').trim()) allNewReqs.push(req);
      });
    });

    if (allNewReqs.length === 0) {
      alert(language === 'zh' ? '未生成任何采购行（请检查设备BOM明细是否为空）。' : 'No rows generated.');
      return;
    }

    const newState: AppState = {
      ...state,
      materialBoms: nextBoms,
      materialRequirements: [...allNewReqs, ...keptReqs],
    };
    setState(newState);
    await handleSaveToDatabase(newState);
    // 站内通知：引用生成需求完成/覆盖写入（服务器模式）
    try {
      if (connectionType === 'server' && apiBase && state.currentUser) {
        const base = normalizeApiBase(apiBase);
        await fetch(`${base}/api/notifications/generate-copy-done`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'x-user-id': state.currentUser.id },
          body: JSON.stringify({
            operatorId: state.currentUser.id,
            operatorName: getEmpName(state.currentUser, language),
            targetProjectId: genProjectId,
            deviceModelId: selectedDeviceId,
            deviceCount,
            bomListCount: purchasingTemplates.length,
            rowCount: allNewReqs.length,
          }),
        });
      }
    } catch (e) {
      // 通知失败不影响主流程
      console.warn('notify generate-copy-done failed:', e);
    }
    addLog?.('material', `引用生成需求：${genProjectId} ← ${selectedDeviceId}（采购BOM版本 ${purchasingTemplates.length} 个，×${deviceCount}）`);
    setShowGenerateModal(false);
    alert(language === 'zh' ? '已复制该设备所有采购BOM到目标项目（并生成采购行）。' : 'Copied all purchasing BOMs to target project.');
    setView?.('procurement');
  };

  return (
    <div className="space-y-6 w-full max-w-full pb-10">
      <div className="flex items-center justify-between gap-4 flex-wrap">
        <div>
          <h2 className="dashboard-title">{language === 'zh' ? '设备BOM中心' : 'Device BOM Center'}</h2>
          <p className="dashboard-subtitle">
            {language === 'zh'
              ? '工程师维护设备型号的母版BOM。点击“引用生成需求”可选择目标项目，并将该设备的全部采购BOM复制过去生成采购行。'
              : 'Maintain master BOM per device model; reference it to generate CS Order requirements.'}
          </p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <button className="btn-secondary text-xs" onClick={() => addDeviceModel()}>
            {language === 'zh' ? '新增设备型号' : 'New Device Model'}
          </button>
          <button className="btn-secondary text-xs" onClick={() => createTemplate()}>
            {language === 'zh' ? '新建BOM版本' : 'New BOM Version'}
          </button>
          <button className="btn-secondary text-xs" disabled={!selectedTemplateId} onClick={() => createTemplate(selectedTemplateId)}>
            {language === 'zh' ? '复制为新版本' : 'Copy as New Version'}
          </button>
          <button className="btn-secondary text-xs" disabled={!selectedTemplateId} onClick={() => void renameCurrentTemplate()}>
            {language === 'zh' ? '重命名BOM' : 'Rename BOM'}
          </button>
          <label className="btn-secondary text-xs cursor-pointer">
            {language === 'zh' ? '导入装配BOM' : 'Import Assembly BOM'}
            <input
              ref={(el) => { assemblyImportRef.current = el; }}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={importSpecificBom('assembly')}
            />
          </label>
          <label className="btn-secondary text-xs cursor-pointer">
            {language === 'zh' ? '导入采购BOM' : 'Import Purchasing BOM'}
            <input
              ref={(el) => { purchasingImportRef.current = el; }}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={importSpecificBom('purchasing')}
            />
          </label>
          <button className="btn-primary text-xs" disabled={!selectedTemplateId} onClick={saveLines}>
            {t.save}
          </button>
          <button className="btn-primary text-xs" disabled={!selectedDeviceId} onClick={() => setShowGenerateModal(true)}>
            {language === 'zh' ? '引用生成需求' : 'Generate Requirements'}
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
        <div className="card space-y-4 lg:col-span-1">
          <div className="space-y-2">
            <label className="dashboard-kicker">{language === 'zh' ? '设备型号' : 'Device Model'}</label>
            <select
              value={selectedDeviceId}
              onChange={(e) => setSelectedDeviceId(e.target.value)}
              className="input-field w-full"
            >
              <option value="">{language === 'zh' ? '请选择' : 'Select'}</option>
              {deviceOptions.map(o => (
                <option key={o.id} value={o.id}>{o.name}</option>
              ))}
            </select>
            <button
              className="btn-secondary text-xs w-full"
              onClick={() => void renameDeviceModel()}
              disabled={!selectedDeviceId}
              title={language === 'zh' ? '修改当前设备型号的名称' : 'Rename selected device model'}
            >
              {language === 'zh' ? '修改设备型号名称' : 'Rename Device Model'}
            </button>
          </div>

          <div className="space-y-2">
            <label className="dashboard-kicker">{language === 'zh' ? 'BOM类别' : 'BOM Type'}</label>
            <select
              value={selectedBomType}
              onChange={(e) => setSelectedBomType(e.target.value as any)}
              className="input-field w-full"
            >
              <option value="Mechanical">{language === 'zh' ? '机械' : 'Mechanical'}</option>
              <option value="Electrical">{language === 'zh' ? '电气' : 'Electrical'}</option>
              <option value="Standard">{language === 'zh' ? '标准件' : 'Standard'}</option>
              <option value="Spare">{language === 'zh' ? '备件' : 'Spare'}</option>
            </select>
          </div>

          <div className="space-y-2">
            <label className="dashboard-kicker">{language === 'zh' ? 'BOM视图' : 'BOM View'}</label>
            <select
              value={selectedBomKind}
              onChange={(e) => setSelectedBomKind(e.target.value as any)}
              className="input-field w-full"
            >
              <option value="assembly">{language === 'zh' ? '装配BOM（树）' : 'Assembly BOM'}</option>
              <option value="purchasing">{language === 'zh' ? '采购BOM（列表）' : 'Purchasing BOM'}</option>
            </select>
          </div>

          <div className="space-y-2">
            <label className="dashboard-kicker">{language === 'zh' ? '版本' : 'Version'}</label>
            <select
              value={selectedTemplateId}
              onChange={(e) => setSelectedTemplateId(e.target.value)}
              className="input-field w-full"
            >
              <option value="">{language === 'zh' ? '请选择' : 'Select'}</option>
              {templateOptions.map(o => (
                <option key={o.id} value={o.id}>{o.label}</option>
              ))}
            </select>
          </div>

          {lastRecon && (
            <div className="p-4 rounded-2xl border border-slate-100 bg-slate-50">
              <div className="font-black text-slate-900 text-sm">BOM对账（装配展开 vs 采购BOM）</div>
              <div className="text-xs font-bold text-slate-500 mt-1">
                差异数：{lastRecon.summary.diffCount}（装配：{lastRecon.summary.assemblyItems}项 / 采购：{lastRecon.summary.purchasingItems}项）
              </div>
              <div className="flex items-center gap-2 mt-3 flex-wrap">
                <button
                  className="btn-secondary text-xs"
                  onClick={() => {
                    const top = lastRecon.diffs.slice(0, 20).map((d: any) => `${d.part}  装配:${d.assemblyQty}  采购:${d.purchasingQty}  Δ:${d.delta}`).join('\n');
                    alert(top || '无差异');
                  }}
                >
                  查看前20条
                </button>
                <button className="btn-secondary text-xs" onClick={() => exportReconToXlsx(lastRecon)}>
                  导出差异xlsx
                </button>
              </div>
            </div>
          )}

          <div className="text-xs text-slate-500 leading-relaxed">
            {language === 'zh'
              ? '提示：先维护各个“采购BOM（列表）”版本明细，保存后点击“引用生成需求”，选择目标项目并按设备数量倍增写入采购表格。'
              : 'Tip: Maintain master BOM lines, then generate requirements by device count.'}
          </div>
        </div>

        <div className="dashboard-card overflow-hidden lg:col-span-3">
          <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-3">
            <div className="text-sm font-bold text-slate-900">
              {currentTemplate
                ? `${currentTemplate.deviceModelId} · ${currentTemplate.bomType} · ${currentTemplate.version}`
                : (language === 'zh' ? '请选择BOM版本' : 'Select BOM version')}
            </div>
            <div className="text-xs text-slate-500">
              {language === 'zh' ? '单台数量（quantityPerDevice）' : 'Per-device quantity'}
            </div>
          </div>

          <div className="p-0">
            {selectedBomKind === 'assembly' ? (
              <div className="p-4">
                <div className="text-xs font-bold text-slate-500 mb-3">
                  提示：默认展开到第2层；点击左侧箭头展开/折叠。第一列为层级序号（level path）。
                </div>
                <div className="overflow-auto max-h-[560px] border border-slate-200 rounded-2xl">
                  <table className="min-w-full text-xs">
                    <thead className="sticky top-0 bg-slate-100">
                      <tr className="text-slate-700">
                        <th className="text-left px-3 py-2 w-[120px]">序号</th>
                        <th className="text-left px-3 py-2 w-[180px]">零件代号/型号</th>
                        <th className="text-left px-3 py-2">描述</th>
                        <th className="text-right px-3 py-2 w-[120px]">数量</th>
                        <th className="text-left px-3 py-2 w-[140px]">BOM结构</th>
                        <th className="text-left px-3 py-2 w-[220px]">文件名称</th>
                      </tr>
                    </thead>
                    <tbody>
                      {visibleAssemblyNodes.map((n, idx) => {
                        const path = String(n.levelPath || '');
                        const depth = path ? path.split('.').length : 1;
                        const hasChildren = (childCountMap.get(path) || 0) > 0;
                        const expanded = !!expandedNodes[path];
                        const qty = Number(n.quantityPerDevice) || 0;
                        const unit = String(n.unit || '').trim();
                        return (
                          <tr key={n.id} className={idx % 2 === 0 ? 'bg-white' : 'bg-slate-50'}>
                            <td className="px-3 py-2 font-mono text-slate-800">
                              <div className="flex items-center gap-1" style={{ paddingLeft: Math.max(0, (depth - 1) * 14) }}>
                                {hasChildren ? (
                                  <button
                                    className="w-5 h-5 rounded hover:bg-slate-200 text-slate-700 font-black"
                                    onClick={() => setExpandedNodes(prev => ({ ...prev, [path]: !expanded }))}
                                    title={expanded ? '折叠' : '展开'}
                                  >
                                    {expanded ? '▾' : '▸'}
                                  </button>
                                ) : (
                                  <span className="inline-block w-5" />
                                )}
                                <span>{path}</span>
                              </div>
                            </td>
                            <td className="px-3 py-2 text-slate-800">{n.partCodeModel}</td>
                            <td className="px-3 py-2 text-slate-800">{n.name}</td>
                            <td className="px-3 py-2 text-right tabular-nums text-slate-900 font-bold">
                              {qty ? qty.toLocaleString() : ''}{unit ? ` ${unit}` : ''}
                            </td>
                            <td className="px-3 py-2 text-slate-600">{n.nodeType || ''}</td>
                            <td className="px-3 py-2 text-slate-600">{n.fileName || ''}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <HotTable
                ref={hotRef}
                className="custom-hot"
                data={tableData}
                colHeaders={[
                  language === 'zh' ? '零件代号/型号' : 'Part Code/Model',
                  language === 'zh' ? '版本' : 'Rev',
                  language === 'zh' ? '名称' : 'Name',
                  language === 'zh' ? '规格' : 'Specs',
                  language === 'zh' ? '单台数量' : 'Qty/Device',
                  language === 'zh' ? '单位' : 'Unit',
                  language === 'zh' ? '品牌' : 'Brand',
                  language === 'zh' ? '供应商' : 'Supplier',
                  language === 'zh' ? '备注' : 'Remark'
                ]}
                rowHeaders={true}
                height={560}
                stretchH="all"
                licenseKey="non-commercial-and-evaluation"
                readOnly={false}
                filters={true}
                dropdownMenu={true}
                contextMenu={true}
                columnSorting={true}
                manualColumnResize={true}
                manualRowResize={true}
                minSpareRows={10}
                columns={[
                  { type: 'text' },
                  { type: 'text' },
                  { type: 'text' },
                  { type: 'text' },
                  { type: 'numeric', numericFormat: { pattern: '0.####' } },
                  { type: 'text' },
                  { type: 'text' },
                  { type: 'text' },
                  { type: 'text' }
                ] as Handsontable.ColumnSettings[]}
              />
            )}
          </div>
        </div>
      </div>

      {/* 引用生成需求：项目下拉弹窗 */}
      <Portal>
        <AnimatePresence>
          {showGenerateModal && (
            <motion.div
              className="fixed inset-0 bg-black/30 z-[1000] flex items-center justify-center p-6"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowGenerateModal(false)}
            >
              <motion.div
                className="bg-white rounded-3xl shadow-2xl border border-slate-100 w-full max-w-xl overflow-hidden"
                initial={{ scale: 0.98, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.98, opacity: 0 }}
                onClick={(e) => e.stopPropagation()}
              >
                <div className="p-6 border-b border-slate-100">
                  <div className="text-lg font-black">{language === 'zh' ? '引用生成需求' : 'Generate Requirements'}</div>
                  <div className="text-xs font-bold text-slate-500 mt-1">
                    {language === 'zh'
                      ? '选择目标项目后，将当前设备型号的全部“采购BOM（列表）”复制到该项目下，并生成采购行（按台数倍增数量）。'
                      : 'Select target project; copy all purchasing BOMs of current device model into the project and generate rows (multiply by device count).'}
                  </div>
                </div>
                <div className="p-6 space-y-4">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '目标项目' : 'Target Project'}</label>
                    <select value={genProjectId} onChange={(e) => setGenProjectId(e.target.value)} className="input-field w-full">
                      <option value="">{language === 'zh' ? '请选择项目' : 'Select project'}</option>
                      {projectOptionsForGen.map(p => (
                        <option key={p.id} value={p.id}>{p.name}</option>
                      ))}
                    </select>
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{language === 'zh' ? '设备台数' : 'Device Count'}</label>
                    <input
                      type="number"
                      min={1}
                      value={genDeviceCount}
                      onChange={(e) => setGenDeviceCount(Math.max(1, Math.floor(Number(e.target.value) || 1)))}
                      className="input-field w-full"
                    />
                  </div>
                  <div className="text-[11px] font-bold text-slate-500 leading-relaxed">
                    {language === 'zh'
                      ? '注意：此操作会对每个生成的 BOM 清单执行“覆盖写入”，即覆盖该 BOM 清单下已有采购行。'
                      : 'Note: each generated BOM list will be overwritten (existing rows under that BOM list will be replaced).'}
                  </div>
                </div>
                <div className="p-6 bg-slate-50/60 border-t border-slate-100 flex gap-3">
                  <button onClick={() => setShowGenerateModal(false)} className="flex-1 btn-secondary">
                    {language === 'zh' ? '取消' : 'Cancel'}
                  </button>
                  <button onClick={() => void generateAllPurchasingBomsToProject()} className="flex-1 btn-primary" disabled={!genProjectId || !selectedDeviceId}>
                    {language === 'zh' ? '开始复制并生成' : 'Generate'}
                  </button>
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>
      </Portal>
    </div>
  );
}

// --- 3. Setup View (Master Data) ---
export function SetupView() {
  const context = useContext(AppContext);
  const [newProj, setNewProj] = useState({ name: '', csOrder: '', managerId: '', budgets: {} as { [id: string]: number }, budgetComments: {} as { [id: string]: string }, comment: '', status: 'active' as 'active' | 'archived' });
  const [newCat2, setNewCat2] = useState({ nameZh: '', nameEn: '' });
  const [newCat3, setNewCat3] = useState({ parentId: '', nameZh: '', nameEn: '' });
  const [editingCat2Id, setEditingCat2Id] = useState<string | null>(null);
  const [editingCat3Id, setEditingCat3Id] = useState<string | null>(null);
  const [editingMappingId, setEditingMappingId] = useState<string | null>(null);
  const [editCatData, setEditCatData] = useState({ nameZh: '', nameEn: '' });
  const [newMapping, setNewMapping] = useState<{
    userId: string,
    category2Ids: string[],
    category3Ids: string[],
    allocationId: string
  }>({ userId: '', category2Ids: [], category3Ids: [], allocationId: '' });

  if (!context) return null;
  const { state, setState, language, handleExportFullSystem, handleSaveToDatabase, addLog } = context;

  const isSuperAdmin = state.currentUser?.role === 'admin';
  const filteredEmployees = getVisibleEmployees(state.currentUser, state.employees).filter(e => isSuperAdmin || e.id !== 'admin-1');

  const makeTempCsOrder = (projectNo: string) => {
    const raw = (projectNo || '').trim();
    const prefix = (raw.split(/\s+/)[0] || 'P').replace(/[^\w\-]/g, '');
    let n = 1;
    const mk = (i: number) => `${prefix}${String(i).padStart(3, '0')}`; // 项目号 + 00x（001/002…）
    const used = new Set((state.projects || []).map(p => String((p as any).csOrder || '').trim()).filter(Boolean));
    while (used.has(mk(n)) && n < 999) n++;
    return mk(n);
  };

  const startEditCat2 = (c: Category2) => {
    setEditingCat2Id(c.id);
    setEditCatData({ nameZh: c.nameZh, nameEn: c.nameEn });
  };

  const saveEditCat2 = () => {
    if (!editingCat2Id) return;
    const newState = {
      ...state,
      categories2: state.categories2.map(c => c.id === editingCat2Id ? { ...c, ...editCatData } : c)
    };
    setState(newState);
    handleSaveToDatabase(newState);
    setEditingCat2Id(null);
  };

  const startEditCat3 = (c: Category3) => {
    setEditingCat3Id(c.id);
    setEditCatData({ nameZh: c.nameZh, nameEn: c.nameEn });
  };

  const saveEditCat3 = () => {
    if (!editingCat3Id) return;
    const newState = {
      ...state,
      categories3: state.categories3.map(c => c.id === editingCat3Id ? { ...c, ...editCatData } : c)
    };
    setState(newState);
    handleSaveToDatabase(newState);
    setEditingCat3Id(null);
  };

  const addProject = async () => {
    if (!newProj.name) return;
    const managerId = isSuperAdmin ? newProj.managerId : (state.currentUser?.id || '');
    if (!managerId) return;
    const id = 'p-' + Date.now();
    const csOrder = (newProj.csOrder || '').trim() || makeTempCsOrder(newProj.name);
    if (!(newProj.csOrder || '').trim()) {
      alert(language === 'zh'
        ? `未填写 CS Order，已自动生成临时号：${csOrder}（后续可在项目配置里修改）`
        : `CS Order was empty. Generated temporary CS Order: ${csOrder}`);
    }
    const proj = { id, ...newProj, csOrder, managerId, visible: true };
    const newState = { ...state, projects: [...state.projects, proj] };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '新增项目' : 'Add Project', proj.name);
    setNewProj({ name: '', csOrder: '', managerId: '', budgets: {}, budgetComments: {}, comment: '', status: 'active' });
  };

  const copyProject = (p: Project) => {
    const newName = `${p.name} (Copy)`;
    const proj = { ...p, id: 'p-' + Date.now(), name: newName, visible: true };
    const newState = { ...state, projects: [...state.projects, proj] };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '复制项目' : 'Copy Project', `${p.name} -> ${newName}`);
  };

  const [editingProjectId, setEditingProjectId] = useState<string | null>(null);

  const startEditProject = (p: any) => {
    setEditingProjectId(p.id);
    setNewProj({
      name: p.name,
      csOrder: p.csOrder || '',
      managerId: p.managerId,
      budgets: p.budgets || {},
      budgetComments: p.budgetComments || {},
      comment: p.comment || '',
      status: p.status || 'active'
    });
  };

  const updateProject = async () => {
    if (!editingProjectId || !newProj.name) return;
    const oldProj = state.projects.find(p => p.id === editingProjectId);
    if (!oldProj) return;
    const csOrder = (newProj.csOrder || '').trim() || makeTempCsOrder(newProj.name);
    if (!(newProj.csOrder || '').trim()) {
      alert(language === 'zh'
        ? `CS Order 为空，已自动生成临时号：${csOrder}（后续可修改）`
        : `CS Order is empty. Generated temporary CS Order: ${csOrder}`);
    }

    const newState = {
      ...state,
      projects: state.projects.map(p => p.id === editingProjectId ? { ...p, ...newProj, csOrder } : p)
    };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '修改项目' : 'Update Project', newProj.name);
    setEditingProjectId(null);
    setNewProj({ name: '', csOrder: '', managerId: '', budgets: {}, budgetComments: {}, comment: '', status: 'active' });
  };

  const removeProject = async (id: string) => {
    const proj = state.projects.find(p => p.id === id);
    if (!window.confirm(language === 'zh' ? `确定要删除项目 "${proj?.name}" 吗？` : `Are you sure you want to delete project "${proj?.name}"?`)) return;
    const newState = {
      ...state,
      projects: state.projects.filter(p => p.id !== id)
    };
    setState(newState);
    handleSaveToDatabase(newState);
    if (proj) addLog(language === 'zh' ? '删除项目' : 'Remove Project', proj.name);
  };

  const addCategory2 = async () => {
    if (!newCat2.nameZh || !newCat2.nameEn) return;
    const id = 'c2-' + Date.now();
    const cat = { id, ...newCat2 };
    const newState = { ...state, categories2: [...state.categories2, cat] };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '新增分类2' : 'Add Category 2', getCatName(cat, language));
    setNewCat2({ nameZh: '', nameEn: '' });
  };

  const removeCategory2 = async (id: string) => {
    const cat = state.categories2.find(c => c.id === id);
    if (!window.confirm(language === 'zh' ? `确定要删除分类 "${getCatName(cat, language)}" 吗？` : `Are you sure?` )) return;
    const newState = {
      ...state,
      categories2: state.categories2.filter(c => c.id !== id),
      categories3: state.categories3.filter(c => c.parentId !== id)
    };
    setState(newState);
    handleSaveToDatabase(newState);
    if (cat) addLog(language === 'zh' ? '删除分类2' : 'Remove Category 2', getCatName(cat, language));
  };

  const addCategory3 = async () => {
    if (!newCat3.nameZh || !newCat3.nameEn || !newCat3.parentId) return;
    const id = 'c3-' + Date.now();
    const cat = { id, ...newCat3 };
    const newState = { ...state, categories3: [...state.categories3, cat] };
    setState(newState);
    handleSaveToDatabase(newState);
    addLog(language === 'zh' ? '新增分类3' : 'Add Category 3', getCatName(cat, language));
    setNewCat3({ ...newCat3, nameZh: '', nameEn: '' });
  };

  const removeCategory3 = async (id: string) => {
    const cat = state.categories3.find(c => c.id === id);
    if (!window.confirm(language === 'zh' ? `确定要删除分类 "${getCatName(cat, language)}" 吗？` : `Are you sure?` )) return;
    const newState = {
      ...state,
      categories3: state.categories3.filter(c => c.id !== id)
    };
    setState(newState);
    handleSaveToDatabase(newState);
    if (cat) addLog(language === 'zh' ? '删除分类3' : 'Remove Category 3', getCatName(cat, language));
  };

  const startEditMapping = (m: any) => {
    setEditingMappingId(m.id);
    setNewMapping({
      userId: m.userId || '',
      category2Ids: m.category2Ids || [],
      category3Ids: m.category3Ids || [],
      allocationId: m.allocationId || ''
    });
    // Scroll to form
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const addMapping = async () => {
    if (newMapping.category2Ids.length === 0 || newMapping.category3Ids.length === 0 || !newMapping.allocationId) return;
    
    // Validation: Check for conflicts
    const conflicts: string[] = [];
    const currentMappings = state.userMappings || [];
    
    newMapping.category2Ids.forEach(c2Id => {
      newMapping.category3Ids.forEach(c3Id => {
        const existing = currentMappings.find(m => 
          m.id !== editingMappingId && // Ignore current mapping if editing
          ((m.userId === newMapping.userId) || 
           ((!m.userId || m.userId === '') && (!newMapping.userId || newMapping.userId === ''))) &&
          m.category2Ids?.includes(c2Id) &&
          m.category3Ids?.includes(c3Id)
        );
        
        if (existing && existing.allocationId !== newMapping.allocationId) {
          const c2 = state.categories2.find(c => c.id === c2Id);
          const c3 = state.categories3.find(c => c.id === c3Id);
          conflicts.push(`${getCatName(c2, language)} > ${getCatName(c3, language)}`);
        }
      });
    });

    if (conflicts.length > 0) {
      alert((language === 'zh' ? '冲突：以下分类组合已分配给其他项目分工：\n' : 'Conflict: The following category combinations are already assigned to other allocations:\n') + conflicts.join('\n'));
      return;
    }

    let newState;
    if (editingMappingId) {
      newState = {
        ...state,
        userMappings: state.userMappings.map(m => m.id === editingMappingId ? { ...m, ...newMapping } : m)
      };
      addLog(language === 'zh' ? '修改映射' : 'Update Mapping', `Mapping: ${editingMappingId}`);
    } else {
      const id = 'um-' + Date.now();
      const mapping = { ...newMapping, id };
      newState = { ...state, userMappings: [...state.userMappings, mapping] };
      addLog(language === 'zh' ? '新增映射' : 'Add Mapping', `Mapping: ${id}`);
    }
    
    setState(newState);
    handleSaveToDatabase(newState);
    setNewMapping({ userId: '', category2Ids: [], category3Ids: [], allocationId: '' });
    setEditingMappingId(null);
  };

  // --- 批量导入“员工多角色 -> 项目分工映射”（从你提供的Excel格式） ---
  const mappingImportRef = useRef<HTMLInputElement | null>(null);
  const applyRoleMappingRows = async (rows: RoleMappingRow[], source: string) => {
    const { mappings, warnings } = buildUserMappingsFromRoleRows(rows, state);
    if (!mappings.length) {
      alert(language === 'zh' ? `未生成任何映射（${source}）。` : `No mappings generated (${source}).`);
      return;
    }
    const userIds = new Set(mappings.map(m => m.userId).filter(Boolean));
    const kept = (state.userMappings || []).filter(m => !m.userId || !userIds.has(m.userId));
    const newState = { ...state, userMappings: [...kept, ...mappings] };
    setState(newState);
    await handleSaveToDatabase(newState);
    const msg = language === 'zh'
      ? `已写入映射：${mappings.length} 条（覆盖涉及人员：${userIds.size} 人）。${warnings.length ? `\n\n注意：有 ${warnings.length} 条未匹配项（已在控制台输出）。` : ''}`
      : `Mappings saved: ${mappings.length} (employees overwritten: ${userIds.size}).`;
    alert(msg);
    if (warnings.length) console.warn('Role mapping warnings:', warnings);
  };

  const importRoleMappingsFromXlsx = async (file: File) => {
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false }) as any[];
        if (!data.length) return alert('Excel为空或无法识别');

        const stripHours = (s: string) => String(s || '').replace(/\([^\)]*h\)\s*$/i, '').trim();
        const rows: RoleMappingRow[] = data.map((r: any) => {
          const emp = String(r['员工'] || r['Employee'] || '').trim();
          const c2 = String(r['分类二'] || r['Category 2'] || '').trim();
          const c3 = String(r['分类三(Top)'] || r['Category 3'] || '').trim();
          const c3List = c3.split(/[；;]+/).map(x => stripHours(x)).filter(Boolean);
          const alloc = String(r['对应项目分工'] || r['Auto Allocation'] || r['Allocation'] || '').trim();
          const note = String(r['备注'] || '').trim();
          return { employeeName: emp, category2: c2, category3List: c3List, allocationName: alloc, note };
        }).filter(r => r.employeeName && r.category2 && r.category3List.length && r.allocationName);

        await applyRoleMappingRows(rows, `Excel: ${file.name}`);
      } catch (e: any) {
        alert(`导入失败：${e?.message || e}`);
      } finally {
        if (mappingImportRef.current) mappingImportRef.current.value = '';
      }
    };
    reader.readAsBinaryString(file);
  };

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-12">
      {/* Project Management Section */}
      <section className="space-y-6">
        <h3 className="text-xl font-black text-slate-900 flex items-center gap-3">
          <Briefcase className="text-blue-600" />
          项目管理
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-4 gap-8 items-start">
          {/* Project List */}
          <div className="lg:col-span-1 bg-white p-8 rounded-[40px] shadow-sm border border-slate-200 space-y-4 h-full max-h-[1200px] overflow-y-auto custom-scrollbar">
            <h4 className="text-sm font-black text-slate-400 uppercase tracking-widest mb-4">项目列表</h4>
            <div className="space-y-3">
              {state.projects.filter(p => p.id !== 'p-general' && p.id !== 'p-leave').filter(p => isSuperAdmin || p.managerId === state.currentUser?.id).map(p => (
                <div 
                  key={p.id} 
                  onClick={() => startEditProject(p)}
                  className={cn(
                    "flex justify-between items-center p-4 rounded-2xl border transition-all cursor-pointer group",
                    editingProjectId === p.id ? "bg-blue-50 border-blue-200" : "bg-slate-50 border-slate-100 hover:border-blue-200"
                  )}
                >
                  <div>
                    <p className="font-black text-slate-900 group-hover:text-blue-600 transition-colors">{p.name}</p>
                    <p className="text-[10px] font-bold text-slate-400 uppercase">
                      PM: {getEmpName(state.employees.find(e => e.id === p.managerId) || null, language)}
                    </p>
                  </div>
                  <div className="flex items-center gap-3">
                    <button 
                      onClick={(e) => { e.stopPropagation(); copyProject(p); }}
                      className="text-slate-300 hover:text-indigo-500 transition-colors"
                      title="复制项目"
                    >
                      <Copy size={16} />
                    </button>
                    <Pencil size={14} className="text-slate-300 group-hover:text-blue-500 transition-colors" />
                    {isSuperAdmin && (
                      <button 
                        onClick={(e) => { e.stopPropagation(); removeProject(p.id); }} 
                        className="text-slate-300 hover:text-rose-500 transition-colors"
                      >
                        <Trash2 size={18} />
                      </button>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Project Form */}
          <div className="lg:col-span-3 bg-white p-8 rounded-[40px] shadow-sm border border-slate-200 space-y-6">
            <div className="flex justify-between items-center">
              <h4 className="text-sm font-black text-slate-400 uppercase tracking-widest">{editingProjectId ? '编辑项目' : '创建项目'}</h4>
              {isSuperAdmin && (
                <button 
                  onClick={() => {
                    const nameZh = window.prompt('输入新分工名称 (中文)');
                    const nameEn = window.prompt('输入新分工名称 (英文)');
                    const type = window.confirm('是否为白领 (WC)?') ? 'WC' : 'BC';
                    if (nameZh && nameEn) {
                      const id = 'alloc-' + Date.now();
                      const newAlloc = { id, nameZh, nameEn, type: type as 'WC' | 'BC' };
                      const newState = { ...state, allocations: [...state.allocations, newAlloc] };
                      setState(newState);
                      handleSaveToDatabase(newState);
                      addLog('新增分工', nameZh);
                    }
                  }}
                  className="text-[10px] font-black text-blue-600 uppercase tracking-widest flex items-center gap-1 hover:text-blue-700 transition-colors"
                >
                  <Plus size={12} />
                  管理分工
                </button>
              )}
            </div>
            <div className="grid grid-cols-1 gap-4">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">项目名称</label>
                <input 
                  type="text" 
                  placeholder="项目名称"
                  value={newProj.name}
                  onChange={e => setNewProj({...newProj, name: e.target.value})}
                  className="input-field"
                />
              </div>
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">CS Order（必填）</label>
                  <button
                    type="button"
                    onClick={() => {
                      if (!newProj.name.trim()) return alert('请先填写项目名称/项目号');
                      setNewProj({ ...newProj, csOrder: makeTempCsOrder(newProj.name) });
                    }}
                    className="text-[10px] font-black text-blue-600 hover:text-blue-700"
                  >
                    自动生成临时号
                  </button>
                </div>
                <input
                  type="text"
                  placeholder="例如：项目号001（无CS Order时先用临时号，后续可修改）"
                  value={(newProj as any).csOrder || ''}
                  onChange={e => setNewProj({ ...newProj, csOrder: e.target.value })}
                  className="input-field"
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">项目经理</label>
                  <select 
                    value={isSuperAdmin ? newProj.managerId : (state.currentUser?.id || '')}
                    onChange={e => setNewProj({ ...newProj, managerId: e.target.value })}
                    className="input-field"
                    disabled={!isSuperAdmin}
                  >
                    <option value="">选择项目经理</option>
                    {filteredEmployees.filter(e => e.role === 'admin').map(emp => (
                      <option key={emp.id} value={emp.id}>{getEmpName(emp, language)}</option>
                    ))}
                  </select>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">项目状态</label>
                  <select 
                    value={newProj.status || 'active'}
                    onChange={e => setNewProj({ ...newProj, status: e.target.value as any })}
                    className="input-field"
                  >
                    <option value="active">进行中</option>
                    <option value="archived">已归档</option>
                  </select>
                </div>
              </div>
              <input 
                type="text" 
                placeholder="项目备注"
                value={newProj.comment}
                onChange={e => setNewProj({...newProj, comment: e.target.value})}
                className="input-field"
              />
            </div>
            <div className="grid grid-cols-2 gap-4 p-2">
              {[...state.allocations].map(alloc => (
                <div key={alloc.id} className="space-y-1 group/alloc">
                  <div className="flex justify-between items-center px-1">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{alloc.nameZh}</label>
                    {isSuperAdmin && (
                      <div className="flex gap-1 opacity-0 group-hover/alloc:opacity-100 transition-opacity">
                        <button 
                          onClick={() => {
                            const newZh = window.prompt('修改中文名', alloc.nameZh);
                            const newEn = window.prompt('修改英文名', alloc.nameEn);
                            if (newZh && newEn) {
                              const newState = {
                                ...state,
                                allocations: state.allocations.map(a => a.id === alloc.id ? { ...a, nameZh: newZh, nameEn: newEn } : a)
                              };
                              setState(newState);
                              handleSaveToDatabase(newState);
                              addLog('修改分工名称', newZh);
                            }
                          }}
                          className="text-slate-300 hover:text-blue-500"
                        >
                          <Pencil size={10} />
                        </button>
                        <button 
                          onClick={() => {
                            if (window.confirm(`确定要删除分工 "${alloc.nameZh}" 吗？`)) {
                              const newState = {
                                ...state,
                                allocations: state.allocations.filter(a => a.id !== alloc.id),
                                deletedAllocations: [alloc, ...state.deletedAllocations].slice(0, 50)
                              };
                              setState(newState);
                              handleSaveToDatabase(newState);
                              addLog('删除分工', alloc.nameZh);
                            }
                          }}
                          className="text-slate-300 hover:text-rose-500"
                        >
                          <Trash2 size={10} />
                        </button>
                      </div>
                    )}
                  </div>
                  <div className="space-y-1">
                    <input 
                      type="number"
                      placeholder="预算工时"
                      value={newProj.budgets[alloc.id] || ''}
                      onChange={e => setNewProj({
                        ...newProj, 
                        budgets: { ...newProj.budgets, [alloc.id]: parseFloat(e.target.value) || 0 }
                      })}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 text-sm font-bold outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all"
                    />
                    <input 
                      type="text"
                      placeholder="备注"
                      value={newProj.budgetComments?.[alloc.id] || ''}
                      onChange={e => setNewProj({
                        ...newProj, 
                        budgetComments: { ...(newProj.budgetComments || {}), [alloc.id]: e.target.value }
                      })}
                      className="w-full bg-slate-50 border border-slate-100 rounded-lg px-3 py-1 text-[10px] font-bold text-slate-400 outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all"
                    />
                  </div>
                </div>
              ))}
            </div>
            {editingProjectId ? (
              <div className="flex gap-4">
                <button onClick={updateProject} className="btn-primary flex-1">
                  保存修改
                </button>
                <button onClick={() => { setEditingProjectId(null); setNewProj({ name: '', managerId: '', budgets: {}, budgetComments: {}, comment: '', status: 'active' }); }} className="bg-slate-100 text-slate-600 px-8 py-4 rounded-2xl font-black uppercase tracking-widest hover:bg-slate-200 transition-all">
                  取消
                </button>
              </div>
            ) : (
              <button onClick={addProject} className="btn-primary w-full">
                创建项目
              </button>
            )}
          </div>
        </div>
      </section>

      {/* Category Management Section */}
      {isSuperAdmin && (
        <section className="space-y-6">
          <h3 className="text-xl font-black text-slate-900 flex items-center gap-3">
            <Layers className="text-indigo-600" />
            分类管理
          </h3>
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
            {/* Category 2 */}
            <div className="bg-white p-8 rounded-[40px] shadow-sm border border-slate-200 space-y-6">
              <h4 className="text-sm font-black text-slate-400 uppercase tracking-widest">分类2</h4>
              <div className="space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <input 
                    type="text" 
                    placeholder="中文名称"
                    value={newCat2.nameZh || ''}
                    onChange={e => setNewCat2({ ...newCat2, nameZh: e.target.value })}
                    className="input-field"
                  />
                  <input 
                    type="text" 
                    placeholder="英文名称"
                    value={newCat2.nameEn || ''}
                    onChange={e => setNewCat2({ ...newCat2, nameEn: e.target.value })}
                    className="input-field"
                  />
                </div>
                <button onClick={addCategory2} className="btn-primary w-full">
                  <Plus size={20} />
                  新增分类2
                </button>
              </div>
              <div className="space-y-2 max-h-[300px] overflow-y-auto">
                {state.categories2.map(c => (
                  <div key={c.id} className="flex justify-between items-center p-3 bg-slate-50 rounded-xl border border-slate-100">
                    {editingCat2Id === c.id ? (
                      <div className="flex gap-2 flex-1 mr-4">
                        <input 
                          type="text" 
                          value={editCatData.nameZh} 
                          onChange={e => setEditCatData({...editCatData, nameZh: e.target.value})}
                          className="flex-1 bg-white border border-slate-200 rounded-lg px-2 py-1 text-sm font-bold"
                        />
                        <input 
                          type="text" 
                          value={editCatData.nameEn} 
                          onChange={e => setEditCatData({...editCatData, nameEn: e.target.value})}
                          className="flex-1 bg-white border border-slate-200 rounded-lg px-2 py-1 text-sm font-bold"
                        />
                        <button onClick={saveEditCat2} className="text-blue-600"><CheckCircle2 size={18} /></button>
                      </div>
                    ) : (
                      <div className="flex-1">
                        <p className="font-bold text-slate-900">{c.nameZh}</p>
                      </div>
                    )}
                    <div className="flex gap-2">
                      <button onClick={() => startEditCat2(c)} className="text-slate-300 hover:text-blue-600"><Pencil size={14} /></button>
                      <button onClick={() => removeCategory2(c.id)} className="text-slate-300 hover:text-rose-600"><Trash2 size={14} /></button>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* Category 3 */}
            <div className="bg-white p-8 rounded-[40px] shadow-sm border border-slate-200 space-y-6">
              <h4 className="text-sm font-black text-slate-400 uppercase tracking-widest">分类3</h4>
              <div className="space-y-4">
                <div className="grid grid-cols-1 gap-4">
                  <select 
                    value={newCat3.parentId}
                    onChange={e => setNewCat3({ ...newCat3, parentId: e.target.value })}
                    className="input-field"
                  >
                    <option value="">选择分类2</option>
                    {state.categories2.map(c => (
                      <option key={c.id} value={c.id}>{getCatName(c, language)}</option>
                    ))}
                  </select>
                  <div className="grid grid-cols-2 gap-4">
                    <input 
                      type="text" 
                      placeholder={language === 'zh' ? '中文名称' : 'Name (ZH)'}
                      value={newCat3.nameZh || ''}
                      onChange={e => setNewCat3({ ...newCat3, nameZh: e.target.value })}
                      className="input-field"
                    />
                    <input 
                      type="text" 
                      placeholder={language === 'zh' ? '英文名称' : 'Name (EN)'}
                      value={newCat3.nameEn || ''}
                      onChange={e => setNewCat3({ ...newCat3, nameEn: e.target.value })}
                      className="input-field"
                    />
                  </div>
                </div>
                <button onClick={addCategory3} className="btn-primary w-full">
                  <Plus size={20} />
                  {language === 'zh' ? '新增分类3' : 'Add Category 3'}
                </button>
              </div>
              <div className="space-y-2 max-h-[300px] overflow-y-auto">
                {state.categories3.filter(c => !newCat3.parentId || c.parentId === newCat3.parentId).map(c => (
                  <div key={c.id} className="flex justify-between items-center p-3 bg-slate-50 rounded-xl border border-slate-100">
                    {editingCat3Id === c.id ? (
                      <div className="flex gap-2 flex-1 mr-4">
                        <input 
                          type="text" 
                          value={editCatData.nameZh} 
                          onChange={e => setEditCatData({...editCatData, nameZh: e.target.value})}
                          className="flex-1 bg-white border border-slate-200 rounded-lg px-2 py-1 text-sm font-bold"
                        />
                        <input 
                          type="text" 
                          value={editCatData.nameEn} 
                          onChange={e => setEditCatData({...editCatData, nameEn: e.target.value})}
                          className="flex-1 bg-white border border-slate-200 rounded-lg px-2 py-1 text-sm font-bold"
                        />
                        <button onClick={saveEditCat3} className="text-blue-600"><CheckCircle2 size={18} /></button>
                      </div>
                    ) : (
                      <div className="flex-1">
                        <p className="font-bold text-slate-900">{getCatName(c, language)}</p>
                        <p className="text-[10px] text-slate-400 uppercase">{getCatName(state.categories2.find(p => p.id === c.parentId), language)}</p>
                      </div>
                    )}
                    <div className="flex gap-2">
                      <button onClick={() => startEditCat3(c)} className="text-slate-300 hover:text-blue-600"><Pencil size={14} /></button>
                      <button onClick={() => removeCategory3(c.id)} className="text-slate-300 hover:text-rose-600"><Trash2 size={14} /></button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </section>
      )}

      <section className="space-y-6 pt-12 border-t border-slate-100">
        <div className="flex justify-between items-center">
          <h3 className="text-xl font-black text-slate-900 flex items-center gap-3">
            {language === 'zh' ? '一人多角色填报映射' : 'One Person Multi-Role Mapping'}
          </h3>
          <div className="flex items-center gap-2 px-4 py-2 bg-blue-50 rounded-xl text-blue-600">
            <span className="text-[10px] font-black uppercase tracking-widest">
              {language === 'zh' ? '说明：定义“人员+分类”对应的“项目分工”，实现填报时身份的自动切换。' : 'Note: Map "User + Category" to "Project Allocation" for automatic identity switching.'}
            </span>
          </div>
        </div>
        <div className="bg-white p-8 rounded-[40px] shadow-sm border border-slate-200 space-y-6">
          <div className="flex justify-between items-center">
            <h4 className="text-sm font-black text-slate-400 uppercase tracking-widest">
              {editingMappingId ? (language === 'zh' ? '编辑映射' : 'Edit Mapping') : (language === 'zh' ? '新增映射' : 'New Mapping')}
            </h4>
            <div className="flex items-center gap-2">
              <button
                onClick={() => applyRoleMappingRows([...EMBEDDED_ROLE_MAPPING_ROWS], '内置映射(v3)')}
                className="btn-secondary py-2 px-3 text-[10px] flex items-center gap-2"
              >
                {language === 'zh' ? '应用已确认映射(v3)' : 'Apply Embedded Mapping (v3)'}
              </button>
              <label className="btn-secondary py-2 px-3 text-[10px] cursor-pointer flex items-center gap-2">
                {language === 'zh' ? '导入映射xlsx(覆盖)' : 'Import Mapping xlsx'}
                <input
                  ref={(el) => { mappingImportRef.current = el; }}
                  type="file"
                  accept=".xlsx,.xls"
                  className="hidden"
                  onChange={e => e.target.files?.[0] && importRoleMappingsFromXlsx(e.target.files[0])}
                />
              </label>
              {editingMappingId && (
                <button 
                  onClick={() => {
                    setEditingMappingId(null);
                    setNewMapping({ userId: '', category2Ids: [], category3Ids: [], allocationId: '' });
                  }}
                  className="text-[10px] font-black text-rose-600 uppercase tracking-widest flex items-center gap-1 hover:text-rose-700 transition-colors"
                >
                  <X size={12} />
                  {language === 'zh' ? '取消编辑' : 'Cancel Edit'}
                </button>
              )}
            </div>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-5 gap-4 items-start bg-slate-50 p-6 rounded-3xl border border-slate-100">
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">{language === 'zh' ? '选择人员 (可选)' : 'Select Employee (Optional)'}</label>
              <select 
                value={newMapping.userId || ''}
                onChange={e => setNewMapping({...newMapping, userId: e.target.value})}
                className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs font-bold outline-none"
              >
                <option value="">{language === 'zh' ? '所有人员 (全局映射)' : 'All Employees (Global)'}</option>
                {filteredEmployees.map(emp => <option key={emp.id} value={emp.id}>{getEmpName(emp, language)}</option>)}
              </select>
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Category 2 (多选)</label>
              <div className="bg-white border border-slate-200 rounded-xl p-2 max-h-[120px] overflow-y-auto space-y-1">
                {state.categories2.map(c => (
                  <label key={c.id} className="flex items-center gap-2 text-[10px] font-bold cursor-pointer hover:bg-slate-50 p-1 rounded">
                    <input 
                      type="checkbox" 
                      checked={newMapping.category2Ids.includes(c.id)}
                      onChange={e => {
                        const ids = e.target.checked 
                          ? [...newMapping.category2Ids, c.id]
                          : newMapping.category2Ids.filter(id => id !== c.id);
                        setNewMapping({ ...newMapping, category2Ids: ids, category3Ids: [] });
                      }}
                    />
                    {getCatName(c, language)}
                  </label>
                ))}
              </div>
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Category 3 (多选)</label>
              <div className="bg-white border border-slate-200 rounded-xl p-2 max-h-[120px] overflow-y-auto space-y-1">
                {state.categories3.filter(c => newMapping.category2Ids.includes(c.parentId)).map(c => (
                  <label key={c.id} className="flex items-center gap-2 text-[10px] font-bold cursor-pointer hover:bg-slate-50 p-1 rounded">
                    <input 
                      type="checkbox" 
                      checked={newMapping.category3Ids.includes(c.id)}
                      onChange={e => {
                        const ids = e.target.checked 
                          ? [...newMapping.category3Ids, c.id]
                          : newMapping.category3Ids.filter(id => id !== c.id);
                        setNewMapping({ ...newMapping, category3Ids: ids });
                      }}
                    />
                    {getCatName(c, language)}
                  </label>
                ))}
              </div>
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">{language === 'zh' ? '对应项目分工' : 'Project Allocation'}</label>
              <select 
                value={newMapping.allocationId}
                onChange={e => setNewMapping({...newMapping, allocationId: e.target.value})}
                className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs font-bold outline-none"
              >
                <option value="">Select Allocation</option>
                {state.allocations.map(a => <option key={a.id} value={a.id}>{language === 'zh' ? a.nameZh : a.nameEn}</option>)}
              </select>
            </div>
            <div className="pt-6">
              <button 
                onClick={addMapping}
                className={cn(
                  "w-full px-4 py-2 rounded-xl font-black text-[10px] uppercase tracking-widest shadow-lg transition-all",
                  editingMappingId ? "bg-indigo-600 hover:bg-indigo-500 text-white" : "bg-blue-600 hover:bg-blue-500 text-white"
                )}
              >
                {editingMappingId ? (language === 'zh' ? '保存修改' : 'Save Changes') : (language === 'zh' ? '添加对应' : 'Add Mapping')}
              </button>
            </div>
          </div>

          <div className="overflow-x-auto rounded-3xl border border-slate-100">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-100">
                  <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Scope</th>
                  <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Category 2</th>
                  <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Category 3</th>
                  <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Auto Allocation</th>
                  <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">Action</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {state.userMappings.map(m => (
                  <tr key={m.id} className="hover:bg-slate-50/30 transition-colors group">
                    <td className="px-6 py-4 text-xs font-bold text-slate-600">
                      {m.userId ? getEmpName(state.employees.find(e => e.id === m.userId), language) : 'Global (All)'}
                    </td>
                    <td className="px-6 py-4 text-[10px] font-bold text-slate-500 max-w-[200px]">
                      <div className="flex flex-wrap gap-1">
                        {m.category2Ids?.map(id => (
                          <span key={id} className="bg-slate-100 px-1.5 py-0.5 rounded">{getCatName(state.categories2.find(c => c.id === id), language)}</span>
                        ))}
                      </div>
                    </td>
                    <td className="px-6 py-4 text-[10px] font-bold text-slate-500 max-w-[300px]">
                      <div className="flex flex-wrap gap-1">
                        {m.category3Ids?.map(id => (
                          <span key={id} className="bg-slate-100 px-1.5 py-0.5 rounded">{getCatName(state.categories3.find(c => c.id === id), language)}</span>
                        ))}
                      </div>
                    </td>
                    <td className="px-6 py-4 text-xs font-bold text-blue-600">
                      {state.allocations.find(a => a.id === m.allocationId)?.nameEn || state.allocations.find(a => a.id === m.allocationId)?.nameZh}
                    </td>
                    <td className="px-6 py-4 text-right">
                      <div className="flex items-center justify-end gap-2">
                        <button 
                          onClick={() => startEditMapping(m)}
                          className="flex items-center gap-1 bg-blue-50 text-blue-600 px-3 py-1.5 rounded-lg font-black text-[10px] uppercase tracking-widest hover:bg-blue-600 hover:text-white transition-all shadow-sm"
                        >
                          <Pencil size={12} />
                          {language === 'zh' ? '修改' : 'Edit'}
                        </button>
                        <button 
                          onClick={() => {
                            if (window.confirm(language === 'zh' ? '确定删除此映射吗？' : 'Delete this mapping?')) {
                              const newState = { ...state, userMappings: state.userMappings.filter(um => um.id !== m.id) };
                              setState(newState);
                              handleSaveToDatabase(newState);
                            }
                          }}
                          className="p-1.5 text-slate-300 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                          title={language === 'zh' ? '删除' : 'Delete'}
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </section>
    </motion.div>
  );
}

export function PageAccessView() {
  const context = useContext(AppContext);
  if (!context) return null;
  const { state, setState, handleSaveToDatabase } = context;

  const isAdmin = state.currentUser?.role === 'admin';
  if (!isAdmin) {
    return (
      <div className="p-10">
        <div className="bg-white rounded-3xl border border-slate-100 p-8">
          <div className="text-xl font-black text-slate-900">无权限</div>
          <div className="text-sm font-bold text-slate-500 mt-2">仅管理员可配置页面可见性。</div>
        </div>
      </div>
    );
  }

  const pages: Array<{ id: string; name: string }> = [
    { id: 'entry', name: '工时填报' },
    { id: 'procurement', name: '采购管理' },
    { id: 'device_bom', name: '设备BOM中心' },
    { id: 'production_forecast', name: '生产预测排班' },
    { id: 'kit_readiness', name: '齐套检查' },
    { id: 'pm_view', name: '项目看板' },
    { id: 'setup', name: '项目配置' },
  ];

  const employees = state.employees.filter(e => e.id !== 'admin-1');

  const getCfg = (pageId: string) => {
    const cfg = state.pageAccess?.[pageId];
    return cfg || { mode: 'all' as const, userIds: [] as string[] };
  };

  const saveCfg = (pageId: string, next: { mode: 'all' | 'restricted'; userIds: string[] }) => {
    setState(prev => {
      const newState = {
        ...prev,
        pageAccess: { ...(prev.pageAccess || {}), [pageId]: next }
      };
      void handleSaveToDatabase(newState);
      return newState;
    });
  };

  return (
    <div className="p-10 space-y-6">
      <div className="bg-white rounded-3xl border border-slate-100 p-8">
        <div className="flex items-start justify-between gap-6">
          <div>
            <div className="text-2xl font-black text-slate-900">页面可见性配置</div>
            <div className="text-sm font-bold text-slate-500 mt-2">
              仅影响左侧导航栏显示与页面访问入口（管理员始终可见）。
            </div>
          </div>
        </div>
      </div>

      <div className="bg-white rounded-3xl border border-slate-100 p-6 space-y-4">
        {pages.map(p => {
          const cfg = getCfg(p.id);
          const restricted = cfg.mode === 'restricted';
          const allowedNames = restricted
            ? employees.filter(e => cfg.userIds.includes(e.id)).map(e => getEmpName(e, 'zh')).join('、')
            : '';

          return (
            <div key={p.id} className="border border-slate-100 rounded-2xl p-5 space-y-3">
              <div className="flex items-center justify-between gap-4">
                <div className="font-black text-slate-900">{p.name}</div>
                <div className="flex items-center gap-2">
                  <button
                    className={cn("px-3 py-1.5 rounded-xl text-[10px] font-black border",
                      !restricted ? "bg-slate-900 text-white border-slate-900" : "bg-white text-slate-500 border-slate-200")}
                    onClick={() => saveCfg(p.id, { mode: 'all', userIds: [] })}
                  >
                    全部可见
                  </button>
                  <button
                    className={cn("px-3 py-1.5 rounded-xl text-[10px] font-black border",
                      restricted ? "bg-blue-600 text-white border-blue-600" : "bg-white text-slate-500 border-slate-200")}
                    onClick={() => saveCfg(p.id, { mode: 'restricted', userIds: cfg.userIds || [] })}
                  >
                    指定人员
                  </button>
                </div>
              </div>

              {restricted && (
                <>
                  <div className="flex items-center gap-1 flex-wrap">
                    {employees.map(emp => {
                      const selected = cfg.userIds.includes(emp.id);
                      return (
                        <button
                          key={emp.id}
                          onClick={() => {
                            const nextIds = selected ? cfg.userIds.filter(id => id !== emp.id) : [...cfg.userIds, emp.id];
                            saveCfg(p.id, { mode: 'restricted', userIds: nextIds });
                          }}
                          className={cn(
                            "px-2 py-0.5 rounded-md text-[10px] font-bold border",
                            selected ? "bg-blue-600 text-white border-blue-600" : "bg-white text-slate-500 border-slate-200"
                          )}
                        >
                          {getEmpName(emp, 'zh')}
                        </button>
                      );
                    })}
                  </div>
                  <div className="text-[10px] font-black text-slate-400">
                    已允许：{allowedNames || '（未选择任何人，将仅管理员可见）'}
                  </div>
                </>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
