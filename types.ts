export type UserRole = 'admin' | 'member' | 'purchaser' | 'pm' | 'approver' | 'warehouse';
export type CollarType = 'WC' | 'BC';

export interface UserProfile {
  uid: string;
  name: string;
  email: string;
  role: UserRole;
  type: CollarType;
}

export interface Project {
  id: string;
  name: string;
  wcBudget: number;
  bcBudget: number;
  wcUsed: number;
  bcUsed: number;
  active: boolean;
  managerId?: string;
}

export interface Category {
  id: string;
  level: 2 | 3;
  name: string;
  parentId?: string;
}

export interface TimesheetEntry {
  id: string;
  userId: string;
  userName: string;
  userType: CollarType;
  date: string; // YYYY-MM-DD
  projectId: string;
  projectName: string;
  category2: string;
  category3: string;
  hours: number;
  remark: string;
  createdAt: string; 
  submittedAt?: string; // NEW: Submission time with minute precision
}

export interface UserSettings {
  [userId: string]: {
    lastProject?: string;
    lastCategory2?: string;
    // Add other persistent preferences here
  };
}

// --- Procurement Types ---

export type MaterialType = 'Mechanical' | 'Electrical';

export type MaterialUrgency = 'Low' | 'Medium' | 'High' | 'Urgent';

export type MaterialStatus = 
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

export interface ApprovalStep {
  role: UserRole;
  userId?: string;
  status: 'Pending' | 'Approved' | 'Rejected';
  comment?: string;
  timestamp?: string;
}

export interface MaterialRequirement {
  id: string;
  projectId: string;
  creatorId: string;
  type: MaterialType;
  code: string;
  name: string;
  model: string;          // NEW
  brand: string;          // NEW
  specs: string;
  quantity: number;
  urgency: MaterialUrgency;
  expectedArrival: string; // ISO date
  
  // Progress Tracking Fields (Purchaser responsible)
  prProgress?: string;      // NEW
  biddingProgress?: string; // NEW
  approvalProgress?: string; // NEW
  decisionProgress?: string; // NEW
  poCompletion?: string;    // NEW
  
  // Warehouse Fields
  actualArrivalTime?: string;    // NEW
  actualArrivalQuantity?: number; // NEW
  inventoryQuantity?: number;     // NEW

  technicalDocs?: string; // URL or description
  status: MaterialStatus;
  comments: string;
  version: number;
  updatedAt: string;
  createdAt: string;
}

export interface BiddingTask {
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

export interface PurchaseRequest {
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

export interface PurchaseOrder {
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
