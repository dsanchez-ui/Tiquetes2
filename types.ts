
export enum UserRole {
  REQUESTER = 'REQUESTER',
  ANALYST = 'ANALYST',
  APPROVER = 'APPROVER' 
}

export enum RequestStatus {
  PENDING_OPTIONS = 'PENDIENTE_OPCIONES', 
  PENDING_SELECTION = 'PENDIENTE_SELECCION', 
  PENDING_APPROVAL = 'PENDIENTE_APROBACION', 
  APPROVED = 'APROBADO',
  REJECTED = 'DENEGADO',
  PROCESSED = 'PROCESADO' 
}

export interface Passenger {
  name: string;
  idNumber: string;
  email?: string; // Added email
}

export interface Integrant {
  idNumber: string;
  name: string;
  email: string;
  approverName: string;
  approverEmail: string;
}

export interface FlightDetails {
  airline: string;
  flightTime: string;
  flightNumber?: string;
  notes: string;
}

export interface Option {
  id: string; // A, B, C
  totalPrice: number;
  
  // Flight Details separated by leg
  outbound: FlightDetails;
  inbound?: FlightDetails; // Optional if one-way, but distinct if exists
  
  hotel?: string;
  generalNotes?: string;
}

export interface SupportFile {
  id: string;
  name: string;
  url: string;
  mimeType: string;
  date: string;
}

export interface SupportData {
  folderId: string;
  folderUrl: string;
  files: SupportFile[];
}

export interface TravelRequest {
  requestId: string; 
  timestamp: string;
  requesterEmail: string;
  
  // Company Info
  company: string;
  businessUnit: string;
  site: string;
  costCenter: string; 
  costCenterName?: string; 
  variousCostCenters?: string; // New field for comma-separated list
  workOrder?: string;
  
  // Trip Info
  origin: string;
  destination: string;
  departureDate: string;
  returnDate?: string; // Optional for One-Way
  departureTimePreference: string;
  returnTimePreference?: string; // Optional for One-Way
  
  // Passengers
  passengers: Passenger[];
  
  // Hotel
  requiresHotel: boolean;
  hotelName?: string;
  nights?: number;
  
  // System/Process Fields
  status: RequestStatus; 
  approverEmail?: string; 
  analystOptions?: Option[]; // Array of structured Options
  selectedOption?: Option; // Single structured Option
  totalCost?: number;
  comments?: string; // Observaciones
  
  // Supports (Post-Approval)
  supportData?: SupportData;
}

export interface CostCenterMaster {
  code: string;
  name: string;
  businessUnit: string;
}

export interface ApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
}
