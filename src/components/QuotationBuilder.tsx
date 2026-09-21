import React, { useState, useEffect, useRef } from "react";
import { Download, Printer, Calendar, Save, Trash2, Plus, Minus, Check, RefreshCw, Copy, X, FileSpreadsheet, Layers, ListPlus, ArrowDownToLine, CheckCheck, Scissors, WrapText, Ship, Percent, FolderOpen, FileEdit, Building2, ArrowLeftRight, User, MapPin, Hash, FileText, Anchor } from "lucide-react";
import { db } from "../lib/firebase";
import { collection, doc, setDoc, deleteDoc, onSnapshot, query, orderBy } from "firebase/firestore";
import { numberToWords } from "../utils/numberToWords";
import { parseClipboardData, parseTSV, cleanCellText } from "../utils/tsvParser";
import { QuotationRow, MergedRegion, SavedDocument, CellFormat, CellFormatMap, CellBorders, CompanyId, CompanyProfile } from "../types";
import { COMPANY_PROFILES } from "../utils/companyProfiles";
import ExcelRibbonToolbar from "./ExcelRibbonToolbar";
import RichTextCell from "./RichTextCell";
import FloatingTextToolbar from "./FloatingTextToolbar";
import ErpSidebar from "./ErpSidebar";
import ErpTopNav from "./ErpTopNav";
import ErpDashboard from "./ErpDashboard";
import { stripHtml, parseNumericInput, applyInlineFormatting, hasActiveSelectionInEditable } from "../utils/textFormatter";
import { generateExcelDocument } from "../utils/excelGenerator";

// Lazy-loaded secondary components for instant initial app startup
const SavedDocumentsPanel = React.lazy(() => import("./SavedDocumentsPanel"));
const ExcelPasteModal = React.lazy(() => import("./ExcelPasteModal"));

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: null,
      email: null
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// Translate OKLCH colors to standard sRGB for canvas compatibility (used for html2pdf rendering)
function oklchToRgb(l: number, c: number, h: number): [number, number, number] {
  const hRad = (h * Math.PI) / 180;
  const a = c * Math.cos(hRad);
  const b = c * Math.sin(hRad);
  
  const l_ = l + 0.3963377774 * a + 0.2158037573 * b;
  const m_ = l - 0.1055613458 * a - 0.0638541728 * b;
  const s_ = l - 0.0894841775 * a - 1.2914855414 * b;
  
  const l3 = l_ * l_ * l_;
  const m3 = m_ * m_ * m_;
  const s3 = s_ * s_ * s_;
  
  const rLinear = +4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
  const gLinear = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
  const bLinear = -0.0041960863 * l3 - 0.703418614 * m3 + 1.7076146995 * s3;
  
  const toSRGB = (x: number) => {
    const clamped = Math.max(0, Math.min(1, x));
    return clamped <= 0.0031308
      ? Math.round(clamped * 12.92 * 255)
      : Math.round((1.055 * Math.pow(clamped, 1 / 2.4) - 0.055) * 255);
  };
  
  return [toSRGB(rLinear), toSRGB(gLinear), toSRGB(bLinear)];
}

function replaceOklchInCss(cssText: string): string {
  if (!cssText || typeof cssText !== 'string' || !cssText.includes("oklch")) {
    return cssText;
  }
  
  const oklchRegex = /oklch\(([^)]+)\)/g;
  return cssText.replace(oklchRegex, (match, innerText) => {
    try {
      const parts = innerText.trim().split(/\s+/).filter((p: string) => p !== '/');
      if (parts.length >= 3) {
        const lStr = parts[0];
        const cStr = parts[1];
        const hStr = parts[2];
        const aStr = parts[3];
        
        let l = lStr.endsWith('%') ? parseFloat(lStr) / 100 : parseFloat(lStr);
        let c = cStr.endsWith('%') ? parseFloat(cStr) / 100 : parseFloat(cStr);
        let h = hStr.endsWith('deg') ? parseFloat(hStr) : parseFloat(hStr);
        
        let alpha = 1;
        if (aStr) {
          alpha = aStr.endsWith('%') ? parseFloat(aStr) / 100 : parseFloat(aStr);
        }
        
        if (isNaN(l) || isNaN(c) || isNaN(h)) return match;
        
        const [r, g, b] = oklchToRgb(l, c, h);
        if (aStr !== undefined) {
          return `rgba(${r}, ${g}, ${b}, ${alpha})`;
        } else {
          return `rgb(${r}, ${g}, ${b})`;
        }
      }
      return match;
    } catch (e) {
      return match;
    }
  });
}

const cleanHtmlText = (str?: string): string => {
  if (!str) return "";
  return str.replace(/<[^>]*>/g, "").replace(/&nbsp;/g, " ").trim();
};

const getDraftStorageKey = (company: CompanyId = "zainee") => `${company}_active_draft_v2`;

const getInitialCompany = (): CompanyId => "zainee";

const loadInitialDraft = (company: CompanyId = "zainee") => {
  if (typeof window === "undefined") return null;
  try {
    const raw = localStorage.getItem(getDraftStorageKey(company));
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
};

export default function QuotationBuilder() {
  const initialCompany = "zainee";
  const initialDraft = useRef(loadInitialDraft(initialCompany)).current;

  // Business Branding & Entity state
  const [activeCompany, setActiveCompany] = useState<CompanyId>("zainee");

  const [docType, setDocType] = useState<"quotation" | "challan" | "invoice">(() => initialDraft?.docType || "quotation");
  const [dateVal, setDateVal] = useState(() => {
    if (initialDraft?.dateVal) return initialDraft.dateVal;
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, "0");
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const yyyy = today.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  });
  const [messers, setMessers] = useState(() => initialDraft?.messers || "");
  const [address, setAddress] = useState(() => initialDraft?.address || "");
  const [vesselName, setVesselName] = useState(() => initialDraft?.vesselName || "");
  const [portBerth, setPortBerth] = useState(() => initialDraft?.portBerth || "");
  const [includeVesselName, setIncludeVesselName] = useState<boolean>(() => {
    if (initialDraft?.includeVesselName !== undefined) return Boolean(initialDraft.includeVesselName);
    return true;
  });
  const [includePortBerth, setIncludePortBerth] = useState<boolean>(() => {
    if (initialDraft?.includePortBerth !== undefined) return Boolean(initialDraft.includePortBerth);
    return true;
  });
  const [currency, setCurrency] = useState<string>(() => (initialDraft?.currency === "USD" ? "Taka" : initialDraft?.currency || "Taka"));
  const [challanNo, setChallanNo] = useState(() => initialDraft?.challanNo || "");
  const [requisitionNo, setRequisitionNo] = useState(() => initialDraft?.requisitionNo || "");
  const [invoiceNo, setInvoiceNo] = useState(() => initialDraft?.invoiceNo || "");
  const [poNumber, setPoNumber] = useState(() => initialDraft?.poNumber || "");
  const [quotationNo, setQuotationNo] = useState(() => initialDraft?.quotationNo || "");
  const [includeInvoiceNo, setIncludeInvoiceNo] = useState<boolean>(() => {
    if (initialDraft?.includeInvoiceNo !== undefined) return Boolean(initialDraft.includeInvoiceNo);
    return true;
  });
  const [includeChallanNo, setIncludeChallanNo] = useState<boolean>(() => {
    if (initialDraft?.includeChallanNo !== undefined) return Boolean(initialDraft.includeChallanNo);
    return true;
  });
  const [includeQuotationNo, setIncludeQuotationNo] = useState<boolean>(() => {
    if (initialDraft?.includeQuotationNo !== undefined) return Boolean(initialDraft.includeQuotationNo);
    return true;
  });
  const [includeRequisitionNo, setIncludeRequisitionNo] = useState<boolean>(() => {
    if (initialDraft?.includeRequisitionNo !== undefined) return Boolean(initialDraft.includeRequisitionNo);
    return true;
  });
  const [includePoNumber, setIncludePoNumber] = useState<boolean>(() => {
    if (initialDraft?.includePoNumber !== undefined) return Boolean(initialDraft.includePoNumber);
    return true;
  });
  const [vatPercent, setVatPercent] = useState<string>(() => initialDraft?.vatPercent !== undefined ? String(initialDraft.vatPercent) : "0");
  const [transportationFee, setTransportationFee] = useState<string>(() => initialDraft?.transportationFee !== undefined ? String(initialDraft.transportationFee) : "0");
  const [includeDiscount, setIncludeDiscount] = useState<boolean>(() => {
    if (initialDraft?.includeDiscount !== undefined) return Boolean(initialDraft.includeDiscount);
    if (initialDraft?.discountValue !== undefined && Number(initialDraft.discountValue) > 0) return true;
    if (initialDraft?.discountPercent !== undefined && Number(initialDraft.discountPercent) > 0) return true;
    return false;
  });
  const [discountType, setDiscountType] = useState<"percentage" | "fixed">(() => {
    if (initialDraft?.discountType === "fixed" || initialDraft?.discountType === "percentage") {
      return initialDraft.discountType;
    }
    return "percentage";
  });
  const [discountValue, setDiscountValue] = useState<string>(() => {
    if (initialDraft?.discountValue !== undefined) return String(initialDraft.discountValue);
    if (initialDraft?.discountPercent !== undefined) return String(initialDraft.discountPercent);
    return "0";
  });
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);
  const [isGeneratingExcel, setIsGeneratingExcel] = useState(false);

  // In-app storage & Auto-Save states
  const [savedDocs, setSavedDocs] = useState<SavedDocument[]>([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedTypeFilter, setSelectedTypeFilter] = useState<"all" | "quotation" | "challan" | "invoice">("all");
  
  // Guard currentDocId: ensure active document uses a fresh zainee ID
  const [currentDocId, setCurrentDocId] = useState<string | null>(() => {
    const rawId = initialDraft?.currentDocId;
    if (rawId && rawId.startsWith("zainee-")) {
      return rawId;
    }
    return "zainee-" + Math.random().toString(36).substring(2, 10) + '-' + Date.now().toString(36);
  });
  const [autoSaveEnabled] = useState<boolean>(true);
  const [lastSavedTime, setLastSavedTime] = useState<string | null>(null);
  const [saveStatus, setSaveStatus] = useState<"idle" | "saving" | "saved" | "error">("idle");

  // Active ERP View state ("dashboard" | "editor" | "saved-docs")
  const [activeView, setActiveView] = useState<"dashboard" | "editor" | "saved-docs">("editor");
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState<boolean>(false);
  const [isMobileSidebarOpen, setIsMobileSidebarOpen] = useState<boolean>(false);

  const currentCompany = COMPANY_PROFILES.zainee || COMPANY_PROFILES.comilla;

  const handleSwitchCompany = (_newCompany: CompanyId) => {
    // Single company mode: Zainee Enterprise
  };

  // Excel Paste Modal and Batch Row Adder states
  const [isExcelModalOpen, setIsExcelModalOpen] = useState(false);
  const [customRowCountInput, setCustomRowCountInput] = useState<string>("10");
  const [customSubtractCountInput, setCustomSubtractCountInput] = useState<string>("10");
  const [targetTotalRowCountInput, setTargetTotalRowCountInput] = useState<string>("");
  const [toastMessage, setToastMessage] = useState<{ text: string; type?: "info" | "success" } | null>(null);

  const showToast = (text: string, type: "info" | "success" = "success") => {
    setToastMessage({ text, type });
    setTimeout(() => {
      setToastMessage(null);
    }, 3500);
  };

  // Grid rows: starts with saved draft rows or 35 blank rows by default
  const [rows, setRows] = useState<QuotationRow[]>(() => {
    if (initialDraft?.rows && Array.isArray(initialDraft.rows) && initialDraft.rows.length > 0) {
      return initialDraft.rows;
    }
    const initialRows: QuotationRow[] = [];
    for (let i = 1; i <= 35; i++) {
      initialRows.push({
        sl: i,
        desc: "",
        qty: "",
        unit: "",
        price: "",
        amount: 0,
      });
    }
    return initialRows;
  });

  const [mergedRegions, setMergedRegions] = useState<MergedRegion[]>(() => initialDraft?.mergedRegions || []);
  const [cellFormats, setCellFormats] = useState<CellFormatMap>(() => initialDraft?.cellFormats || {});
  const [selectedRowIndex, setSelectedRowIndex] = useState<number>(0);
  const [selectedCell, setSelectedCell] = useState<{ rowIndex: number; colIndex: number } | null>({ rowIndex: 0, colIndex: 0 });
  const [isSelecting, setIsSelecting] = useState<boolean>(false);
  const [selectionStart, setSelectionStart] = useState<{ rowIndex: number; colIndex: number } | null>({ rowIndex: 0, colIndex: 0 });
  const [selectionEnd, setSelectionEnd] = useState<{ rowIndex: number; colIndex: number } | null>({ rowIndex: 0, colIndex: 0 });

  const [contextMenu, setContextMenu] = useState<{
    visible: boolean;
    x: number;
    y: number;
    rowIndex: number;
    colIndex: number;
  } | null>(null);

  const dateRef = useRef<HTMLInputElement>(null);
  const autoSaveTimerRef = useRef<NodeJS.Timeout | null>(null);

  // Document History Stack for Full Undo/Redo via Shortcuts
  interface DocumentSnapshot {
    rows: QuotationRow[];
    mergedRegions: MergedRegion[];
    cellFormats: CellFormatMap;
    docType: "quotation" | "challan" | "invoice";
    messers: string;
    address: string;
    vesselName: string;
    portBerth: string;
    includeVesselName: boolean;
    includePortBerth: boolean;
    currency: string;
    challanNo: string;
    requisitionNo: string;
    invoiceNo: string;
    poNumber: string;
    quotationNo: string;
    includeInvoiceNo: boolean;
    includeChallanNo: boolean;
    includeQuotationNo: boolean;
    includeRequisitionNo: boolean;
    includePoNumber: boolean;
    vatPercent: string;
    transportationFee: string;
    includeDiscount: boolean;
    discountType: "percentage" | "fixed";
    discountValue: string;
  }

  const undoStackRef = useRef<DocumentSnapshot[]>([]);
  const redoStackRef = useRef<DocumentSnapshot[]>([]);
  const isUndoRedoActionRef = useRef<boolean>(false);
  const lastTypingTimeRef = useRef<number>(0);
  const [historyChangeCount, setHistoryChangeCount] = useState<number>(0);

  const getCurrentSnapshot = (): DocumentSnapshot => ({
    rows: JSON.parse(JSON.stringify(rows)),
    mergedRegions: JSON.parse(JSON.stringify(mergedRegions)),
    cellFormats: JSON.parse(JSON.stringify(cellFormats)),
    docType,
    messers,
    address,
    vesselName,
    portBerth,
    includeVesselName,
    includePortBerth,
    currency,
    challanNo,
    requisitionNo,
    invoiceNo,
    poNumber,
    quotationNo,
    includeInvoiceNo,
    includeChallanNo,
    includeQuotationNo,
    includeRequisitionNo,
    includePoNumber,
    vatPercent,
    transportationFee,
    includeDiscount,
    discountType,
    discountValue,
  });

  const recordChange = () => {
    if (isUndoRedoActionRef.current) return;
    const snap = getCurrentSnapshot();
    undoStackRef.current.push(snap);
    if (undoStackRef.current.length > 60) {
      undoStackRef.current.shift();
    }
    redoStackRef.current = [];
    setHistoryChangeCount((c) => c + 1);
  };

  const handleUndo = () => {
    if (undoStackRef.current.length === 0) {
      showToast("Nothing to undo", "info");
      return;
    }
    const current = getCurrentSnapshot();
    const previous = undoStackRef.current.pop()!;
    redoStackRef.current.push(current);
    if (redoStackRef.current.length > 60) {
      redoStackRef.current.shift();
    }

    isUndoRedoActionRef.current = true;
    setRows(previous.rows);
    setMergedRegions(previous.mergedRegions);
    setCellFormats(previous.cellFormats);
    setDocType(previous.docType);
    setMessers(previous.messers);
    setAddress(previous.address);
    setVesselName(previous.vesselName);
    setPortBerth(previous.portBerth);
    setIncludeVesselName(previous.includeVesselName);
    setIncludePortBerth(previous.includePortBerth);
    setCurrency(previous.currency);
    setChallanNo(previous.challanNo);
    setRequisitionNo(previous.requisitionNo);
    setInvoiceNo(previous.invoiceNo);
    setPoNumber(previous.poNumber);
    setQuotationNo(previous.quotationNo || "");
    setIncludeInvoiceNo(previous.includeInvoiceNo !== undefined ? previous.includeInvoiceNo : true);
    setIncludeChallanNo(previous.includeChallanNo !== undefined ? previous.includeChallanNo : true);
    setIncludeQuotationNo(previous.includeQuotationNo !== undefined ? previous.includeQuotationNo : true);
    setIncludeRequisitionNo(previous.includeRequisitionNo !== undefined ? previous.includeRequisitionNo : true);
    setIncludePoNumber(previous.includePoNumber !== undefined ? previous.includePoNumber : true);
    setVatPercent(previous.vatPercent);
    setTransportationFee(previous.transportationFee);
    setIncludeDiscount(previous.includeDiscount);
    setDiscountType(previous.discountType);
    setDiscountValue(previous.discountValue);
    setHistoryChangeCount((c) => c + 1);

    setTimeout(() => {
      isUndoRedoActionRef.current = false;
    }, 50);

    showToast("Undo (Ctrl+Z)", "info");
  };

  const handleRedo = () => {
    if (redoStackRef.current.length === 0) {
      showToast("Nothing to redo", "info");
      return;
    }
    const current = getCurrentSnapshot();
    const next = redoStackRef.current.pop()!;
    undoStackRef.current.push(current);
    if (undoStackRef.current.length > 60) {
      undoStackRef.current.shift();
    }

    isUndoRedoActionRef.current = true;
    setRows(next.rows);
    setMergedRegions(next.mergedRegions);
    setCellFormats(next.cellFormats);
    setDocType(next.docType);
    setMessers(next.messers);
    setAddress(next.address);
    setVesselName(next.vesselName);
    setPortBerth(next.portBerth);
    setIncludeVesselName(next.includeVesselName);
    setIncludePortBerth(next.includePortBerth);
    setCurrency(next.currency);
    setChallanNo(next.challanNo);
    setRequisitionNo(next.requisitionNo);
    setInvoiceNo(next.invoiceNo);
    setPoNumber(next.poNumber);
    setQuotationNo(next.quotationNo || "");
    setIncludeInvoiceNo(next.includeInvoiceNo !== undefined ? next.includeInvoiceNo : true);
    setIncludeChallanNo(next.includeChallanNo !== undefined ? next.includeChallanNo : true);
    setIncludeQuotationNo(next.includeQuotationNo !== undefined ? next.includeQuotationNo : true);
    setIncludeRequisitionNo(next.includeRequisitionNo !== undefined ? next.includeRequisitionNo : true);
    setIncludePoNumber(next.includePoNumber !== undefined ? next.includePoNumber : true);
    setVatPercent(next.vatPercent);
    setTransportationFee(next.transportationFee);
    setIncludeDiscount(next.includeDiscount);
    setDiscountType(next.discountType);
    setDiscountValue(next.discountValue);
    setHistoryChangeCount((c) => c + 1);

    setTimeout(() => {
      isUndoRedoActionRef.current = false;
    }, 50);

    showToast("Redo (Ctrl+Y)", "info");
  };

  const triggerDatePicker = () => {
    if (dateRef.current) {
      try {
        dateRef.current.showPicker();
      } catch (e) {
        dateRef.current.focus();
        dateRef.current.click();
      }
    }
  };

  const handleDatePickerChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const dateStr = e.target.value;
    if (!dateStr) return;
    const [yyyy, mm, dd] = dateStr.split("-");
    setDateVal(`${dd}/${mm}/${yyyy}`);
  };

  // Listen to Firestore documents from "documents" (Comilla Traders)
  useEffect(() => {
    const parseDocSnapshot = (docSnap: any): SavedDocument => {
      const data = docSnap.data();
      const docRows = (data.rows || []).map((r: any) => ({
        sl: Number(r.sl) || 0,
        desc: String(r.desc ?? ""),
        qty: String(r.qty ?? ""),
        unit: String(r.unit ?? ""),
        price: String(r.price ?? ""),
        amount: Number(r.amount) || 0,
      }));
      
      const docMergedRegions: MergedRegion[] = Array.isArray(data.mergedRegions)
        ? data.mergedRegions.map((m: any) => ({
            id: String(m.id ?? `region-${Math.random().toString(36).substring(2, 9)}`),
            startRow: Number(m.startRow) || 0,
            endRow: Number(m.endRow) || 0,
            startCol: Number(m.startCol) ?? 0,
            endCol: Number(m.endCol) ?? 0,
          }))
        : [];

      return {
        id: docSnap.id,
        companyId: "zainee",
        companyName: data.companyName || COMPANY_PROFILES.zainee.name,
        name: data.name || "",
        createdAt: data.createdAt || "",
        updatedAt: data.updatedAt || "",
        docType: data.docType || "quotation",
        dateVal: data.dateVal || "",
        messers: data.messers || "",
        address: data.address || "",
        vesselName: data.vesselName || "",
        portBerth: data.portBerth || "",
        currency: (data.currency === "USD" || !data.currency) ? "Taka" : data.currency,
        discountPercent: data.discountPercent || 0,
        includeDiscount: data.includeDiscount !== undefined ? Boolean(data.includeDiscount) : ((data.discountValue && data.discountValue > 0) || (data.discountPercent && data.discountPercent > 0)),
        discountType: data.discountType || "percentage",
        discountValue: data.discountValue !== undefined ? data.discountValue : (data.discountPercent || 0),
        challanNo: data.challanNo || "",
        requisitionNo: data.requisitionNo || "",
        invoiceNo: data.invoiceNo || "",
        poNumber: data.poNumber || "",
        rows: docRows,
        mergedRegions: docMergedRegions,
        cellFormats: (data.cellFormats as CellFormatMap) || {},
        vatPercent: data.vatPercent,
        transportationFee: data.transportationFee
      };
    };

    let zaineeDocs: SavedDocument[] = [];

    const updateAllDocs = () => {
      const all = [...zaineeDocs].sort((a, b) => 
        new Date(b.updatedAt || 0).getTime() - new Date(a.updatedAt || 0).getTime()
      );
      // Deduplicate by ID
      const uniqueDocs = Array.from(new Map(all.map(d => [d.id, d])).values());
      setSavedDocs(uniqueDocs);
    };

    let fallbackUnsub: (() => void) | null = null;
    const qZainee = query(collection(db, "zainee_documents"), orderBy("updatedAt", "desc"));
    const unsubZainee = onSnapshot(qZainee, (snapshot) => {
      zaineeDocs = snapshot.docs.map(d => parseDocSnapshot(d));
      updateAllDocs();
    }, (error) => {
      console.warn("zainee_documents listener notice:", error);
      try {
        fallbackUnsub = onSnapshot(collection(db, "zainee_documents"), (fallbackSnapshot) => {
          zaineeDocs = fallbackSnapshot.docs.map(d => parseDocSnapshot(d));
          updateAllDocs();
        }, (err) => {
          console.warn("zainee_documents fallback collection error:", err);
        });
      } catch (err) {
        console.warn("Failed to subscribe to zainee_documents fallback:", err);
      }
    });

    return () => {
      unsubZainee();
      if (fallbackUnsub) {
        fallbackUnsub();
      }
    };
  }, []);

  const generateUUID = () => {
    return "zainee-" + Math.random().toString(36).substring(2, 10) + '-' + Date.now().toString(36);
  };

  const saveCurrentDocToApp = async (customName?: string) => {
    if (autoSaveTimerRef.current) {
      clearTimeout(autoSaveTimerRef.current);
      autoSaveTimerRef.current = null;
    }

    const now = new Date().toISOString();
    let docIdentifier = "";
    if (docType === "challan" && challanNo) {
      docIdentifier = ` (Challan #${challanNo})`;
    } else if (docType === "invoice" && invoiceNo) {
      docIdentifier = ` (Invoice #${invoiceNo})`;
    }

    const docTypeLabel = docType === "invoice" ? "Invoice" : docType === "challan" ? "Challan" : "Quotation";
    const defaultName = `${docTypeLabel}${docIdentifier} - ${messers || "Unnamed Client"} (${dateVal})`;
    const nameToUse = customName || savedDocs.find(d => d.id === currentDocId)?.name || defaultName;

    const targetCollection = "zainee_documents";
    const docId = (currentDocId && currentDocId.startsWith("zainee-")) ? currentDocId : generateUUID();
    if (currentDocId !== docId) {
      setCurrentDocId(docId);
    }

    const sanitizedRows = rows.map(r => ({
      sl: Number(r.sl) || 0,
      desc: String(r.desc ?? ""),
      qty: String(r.qty ?? ""),
      unit: String(r.unit ?? ""),
      price: String(r.price ?? ""),
      amount: Number(r.amount) || 0
    }));

    const sanitizedMergedRegions = mergedRegions.map(m => ({
      id: String(m.id),
      startRow: Number(m.startRow) || 0,
      endRow: Number(m.endRow) || 0,
      startCol: Number(m.startCol) ?? 0,
      endCol: Number(m.endCol) ?? 0
    }));

    const docData: SavedDocument = {
      id: docId,
      companyId: "zainee",
      companyName: "Zainee Enterprise",
      name: String(nameToUse || "Unnamed Document"),
      createdAt: String(savedDocs.find(d => d.id === docId)?.createdAt || now),
      updatedAt: String(now),
      docType: docType as "quotation" | "challan" | "invoice",
      dateVal: String(dateVal || ""),
      messers: String(messers || ""),
      address: String(address || ""),
      vesselName: String(vesselName || ""),
      portBerth: String(portBerth || ""),
      includeVesselName: Boolean(includeVesselName),
      includePortBerth: Boolean(includePortBerth),
      currency: String(currency || ""),
      discountPercent: discountType === "percentage" ? (parseFloat(discountValue) || 0) : 0,
      includeDiscount: Boolean(includeDiscount),
      discountType,
      discountValue: parseFloat(discountValue) || 0,
      challanNo: String(challanNo || ""),
      requisitionNo: String(requisitionNo || ""),
      invoiceNo: String(invoiceNo || ""),
      poNumber: String(poNumber || ""),
      quotationNo: String(quotationNo || ""),
      includeInvoiceNo: Boolean(includeInvoiceNo),
      includeChallanNo: Boolean(includeChallanNo),
      includeQuotationNo: Boolean(includeQuotationNo),
      includeRequisitionNo: Boolean(includeRequisitionNo),
      includePoNumber: Boolean(includePoNumber),
      rows: sanitizedRows,
      mergedRegions: sanitizedMergedRegions,
      cellFormats: { ...cellFormats },
      vatPercent: parseFloat(vatPercent) || 0,
      transportationFee: parseFloat(transportationFee) || 0
    };

    setSaveStatus("saving");
    try {
      await setDoc(doc(db, targetCollection, docId), docData);
      if (currentDocId !== docId) {
        setCurrentDocId(docId);
      }
      const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
      setLastSavedTime(timeStr);
      setSaveStatus("saved");
      setTimeout(() => setSaveStatus("idle"), 3000);
      showToast("Saved to Zainee Enterprise", "success");
    } catch (e) {
      console.error("Error saving document:", e);
      setSaveStatus("error");
      setTimeout(() => setSaveStatus("idle"), 3000);
      handleFirestoreError(e, OperationType.WRITE, `${targetCollection}/${docId}`);
    }
  };

  const resetSheetFields = (companyToReset: CompanyId = activeCompany) => {
    setDocType("quotation");
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, "0");
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const yyyy = today.getFullYear();
    setDateVal(`${dd}/${mm}/${yyyy}`);
    setMessers("");
    setAddress("");
    setVesselName("");
    setPortBerth("");
    setIncludeVesselName(true);
    setIncludePortBerth(true);
    setCurrency("Taka");
    setIncludeDiscount(false);
    setDiscountType("percentage");
    setDiscountValue("0");
    setChallanNo("");
    setRequisitionNo("");
    setInvoiceNo("");
    setPoNumber("");
    setQuotationNo("");
    setIncludeInvoiceNo(true);
    setIncludeChallanNo(true);
    setIncludeQuotationNo(true);
    setIncludeRequisitionNo(true);
    setIncludePoNumber(true);
    setVatPercent("0");
    setTransportationFee("0");
    
    const initialRows: QuotationRow[] = [];
    for (let i = 1; i <= 35; i++) {
      initialRows.push({
        sl: i,
        desc: "",
        qty: "",
        unit: "",
        price: "",
        amount: 0,
      });
    }
    setRows(initialRows);
    setMergedRegions([]);
    setCellFormats({});
    setCurrentDocId(null);
    setLastSavedTime(null);
    if (typeof window !== "undefined") {
      localStorage.removeItem(getDraftStorageKey(companyToReset));
    }
  };

  const loadSavedDoc = (doc: SavedDocument) => {
    setDocType(doc.docType);
    setDateVal(doc.dateVal);
    setMessers(doc.messers);
    setAddress(doc.address);
    setVesselName(doc.vesselName || "");
    setPortBerth(doc.portBerth || "");
    setIncludeVesselName(doc.includeVesselName !== undefined ? Boolean(doc.includeVesselName) : (doc.vesselName !== undefined && doc.vesselName.trim() !== "" ? true : true));
    setIncludePortBerth(doc.includePortBerth !== undefined ? Boolean(doc.includePortBerth) : (doc.portBerth !== undefined && doc.portBerth.trim() !== "" ? true : true));
    setCurrency(doc.currency === "USD" || !doc.currency ? "Taka" : doc.currency);
    
    const docHasDiscount = doc.includeDiscount !== undefined 
      ? Boolean(doc.includeDiscount) 
      : ((doc.discountValue !== undefined && Number(doc.discountValue) > 0) || (doc.discountPercent !== undefined && Number(doc.discountPercent) > 0));
    setIncludeDiscount(docHasDiscount);
    setDiscountType(doc.discountType === "fixed" ? "fixed" : "percentage");
    setDiscountValue(doc.discountValue !== undefined ? String(doc.discountValue) : (doc.discountPercent !== undefined ? String(doc.discountPercent) : "0"));

    setChallanNo(doc.challanNo || "");
    setRequisitionNo(doc.requisitionNo || "");
    setInvoiceNo(doc.invoiceNo || "");
    setPoNumber(doc.poNumber || "");
    setQuotationNo(doc.quotationNo || "");
    setIncludeInvoiceNo(doc.includeInvoiceNo !== undefined ? Boolean(doc.includeInvoiceNo) : true);
    setIncludeChallanNo(doc.includeChallanNo !== undefined ? Boolean(doc.includeChallanNo) : true);
    setIncludeQuotationNo(doc.includeQuotationNo !== undefined ? Boolean(doc.includeQuotationNo) : true);
    setIncludeRequisitionNo(doc.includeRequisitionNo !== undefined ? Boolean(doc.includeRequisitionNo) : true);
    setIncludePoNumber(doc.includePoNumber !== undefined ? Boolean(doc.includePoNumber) : true);
    setRows(doc.rows.map(r => ({ ...r })));
    setMergedRegions((doc.mergedRegions || []).map(m => ({ ...m })));
    setCellFormats(doc.cellFormats ? { ...doc.cellFormats } : {});
    setCurrentDocId(doc.id);
    setLastSavedTime(null);
    setVatPercent(doc.vatPercent !== undefined ? String(doc.vatPercent) : "0");
    setTransportationFee(doc.transportationFee !== undefined ? String(doc.transportationFee) : "0");

    if (typeof window !== "undefined") {
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  };

  const deleteSavedDoc = async (id: string, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    
    const docToDelete = savedDocs.find(d => d.id === id);
    const confirmMsg = `Delete "${docToDelete?.name || 'this document'}"?`;

    if (window.confirm(confirmMsg)) {
      try {
        await deleteDoc(doc(db, "zainee_documents", id));
        
        if (currentDocId === id) {
          resetSheetFields("zainee");
        }
        showToast("Document deleted.", "info");
      } catch (e) {
        console.error("Error deleting document:", e);
        handleFirestoreError(e, OperationType.DELETE, `zainee_documents/${id}`);
      }
    }
  };

  const renameSavedDoc = async (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const documentObj = savedDocs.find(d => d.id === id);
    if (!documentObj) return;

    const newName = window.prompt("Rename this document:", documentObj.name);
    if (newName && newName.trim() !== "") {
      try {
        const updatedData = {
          ...documentObj,
          name: newName.trim(),
          updatedAt: new Date().toISOString()
        };
        await setDoc(doc(db, "zainee_documents", id), updatedData);
        showToast("Document renamed");
      } catch (e) {
        console.error("Error renaming document:", e);
        handleFirestoreError(e, OperationType.WRITE, `zainee_documents/${id}`);
      }
    }
  };

  const startNewDoc = (targetType?: "quotation" | "challan" | "invoice", skipConfirm = false) => {
    if (skipConfirm || window.confirm("Start a new document? Unsaved changes on your active sheet will be overwritten.")) {
      resetSheetFields(activeCompany);
      if (targetType) {
        setDocType(targetType);
      }
    }
  };

  const duplicateCurrentDoc = async () => {
    const targetCollection = "zainee_documents";
    const defaultName = `Copy of ${messers ? messers.trim() : "Quotation"} (${dateVal})`;
    const docName = window.prompt("Enter a name for the duplicated copy:", defaultName);
    if (!docName || docName.trim() === "") return;

    setSaveStatus("saving");
    const newId = generateUUID();
    try {
      const docPayload: SavedDocument = {
        id: newId,
        companyId: "zainee",
        companyName: "Zainee Enterprise",
        name: docName.trim(),
        docType,
        dateVal,
        messers,
        address,
        vesselName: vesselName || "",
        portBerth: portBerth || "",
        includeVesselName: Boolean(includeVesselName),
        includePortBerth: Boolean(includePortBerth),
        currency: currency || "",
        discountPercent: discountType === "percentage" ? (parseFloat(discountValue) || 0) : 0,
        includeDiscount: Boolean(includeDiscount),
        discountType,
        discountValue: parseFloat(discountValue) || 0,
        challanNo: challanNo || "",
        requisitionNo: requisitionNo || "",
        invoiceNo: invoiceNo || "",
        poNumber: poNumber || "",
        quotationNo: quotationNo || "",
        includeInvoiceNo: Boolean(includeInvoiceNo),
        includeChallanNo: Boolean(includeChallanNo),
        includeQuotationNo: Boolean(includeQuotationNo),
        includeRequisitionNo: Boolean(includeRequisitionNo),
        includePoNumber: Boolean(includePoNumber),
        rows: rows.map(r => ({ ...r })),
        mergedRegions: mergedRegions.map(m => ({ ...m })),
        cellFormats: { ...cellFormats },
        vatPercent: parseFloat(vatPercent) || 0,
        transportationFee: parseFloat(transportationFee) || 0,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };

      await setDoc(doc(db, targetCollection, newId), docPayload);
      setCurrentDocId(newId);
      setSaveStatus("saved");
      setTimeout(() => setSaveStatus("idle"), 3000);
      showToast("Document duplicated", "success");
    } catch (err) {
      console.error("Error duplicating document:", err);
      setSaveStatus("error");
      setTimeout(() => setSaveStatus("idle"), 3000);
      handleFirestoreError(err, OperationType.WRITE, `${targetCollection}/${newId}`);
    }
  };

  // Debounced Auto-Save
  useEffect(() => {
    if (!autoSaveEnabled) return;

    const hasAnyContent = 
      currentDocId !== null ||
      messers.trim() !== "" || 
      address.trim() !== "" || 
      challanNo.trim() !== "" || 
      invoiceNo.trim() !== "" ||
      poNumber.trim() !== "" ||
      rows.some(r => r.desc.trim() !== "");

    if (!hasAnyContent) return;

    const timer = setTimeout(async () => {
      const targetCollection = "zainee_documents";
      const docId = (currentDocId && currentDocId.startsWith("zainee-")) ? currentDocId : generateUUID();
      if (currentDocId !== docId) {
        setCurrentDocId(docId);
      }
      const now = new Date().toISOString();
      let docIdentifier = "";
      if (docType === "challan" && challanNo) {
        docIdentifier = ` (Challan #${challanNo})`;
      } else if (docType === "invoice" && invoiceNo) {
        docIdentifier = ` (Invoice #${invoiceNo})`;
      }

      const docTypeLabel = docType === "invoice" ? "Invoice" : docType === "challan" ? "Challan" : "Quotation";
      const defaultName = `${docTypeLabel}${docIdentifier} - ${messers || "Unnamed Client"} (${dateVal})`;
      const nameToUse = savedDocs.find(d => d.id === currentDocId)?.name || defaultName;

      const sanitizedRows = rows.map(r => ({
        sl: Number(r.sl) || 0,
        desc: String(r.desc ?? ""),
        qty: String(r.qty ?? ""),
        unit: String(r.unit ?? ""),
        price: String(r.price ?? ""),
        amount: Number(r.amount) || 0
      }));

      const sanitizedMergedRegions = mergedRegions.map(m => ({
        id: String(m.id),
        startRow: Number(m.startRow) || 0,
        endRow: Number(m.endRow) || 0,
        startCol: Number(m.startCol) ?? 0,
        endCol: Number(m.endCol) ?? 0
      }));

      const docData: SavedDocument = {
        id: docId,
        companyId: "zainee",
        companyName: "Zainee Enterprise",
        name: String(nameToUse || "Unnamed Document"),
        createdAt: String(savedDocs.find(d => d.id === docId)?.createdAt || now),
        updatedAt: String(now),
        docType: docType as "quotation" | "challan" | "invoice",
        dateVal: String(dateVal || ""),
        messers: String(messers || ""),
        address: String(address || ""),
        vesselName: String(vesselName || ""),
        portBerth: String(portBerth || ""),
        includeVesselName: Boolean(includeVesselName),
        includePortBerth: Boolean(includePortBerth),
        currency: String(currency || ""),
        discountPercent: discountType === "percentage" ? (parseFloat(discountValue) || 0) : 0,
        includeDiscount: Boolean(includeDiscount),
        discountType,
        discountValue: parseFloat(discountValue) || 0,
        challanNo: String(challanNo || ""),
        requisitionNo: String(requisitionNo || ""),
        invoiceNo: String(invoiceNo || ""),
        poNumber: String(poNumber || ""),
        quotationNo: String(quotationNo || ""),
        includeInvoiceNo: Boolean(includeInvoiceNo),
        includeChallanNo: Boolean(includeChallanNo),
        includeQuotationNo: Boolean(includeQuotationNo),
        includeRequisitionNo: Boolean(includeRequisitionNo),
        includePoNumber: Boolean(includePoNumber),
        rows: sanitizedRows,
        mergedRegions: sanitizedMergedRegions,
        cellFormats: { ...cellFormats },
        vatPercent: parseFloat(vatPercent) || 0,
        transportationFee: parseFloat(transportationFee) || 0
      };

      setSaveStatus("saving");
      try {
        await setDoc(doc(db, targetCollection, docId), docData);
        if (currentDocId !== docId) {
          setCurrentDocId(docId);
        }
        const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        setLastSavedTime(timeStr);
        setSaveStatus("saved");
        setTimeout(() => setSaveStatus("idle"), 3000);
      } catch (e) {
        console.error("Auto-save failed:", e);
        setSaveStatus("error");
        setTimeout(() => setSaveStatus("idle"), 3000);
        handleFirestoreError(e, OperationType.WRITE, `${targetCollection}/${docId}`);
      }
    }, 1500);

    autoSaveTimerRef.current = timer;
    return () => clearTimeout(timer);
  }, [
    activeCompany,
    docType,
    dateVal,
    messers,
    address,
    vesselName,
    portBerth,
    includeVesselName,
    includePortBerth,
    currency,
    includeDiscount,
    discountType,
    discountValue,
    challanNo,
    requisitionNo,
    invoiceNo,
    poNumber,
    quotationNo,
    includeInvoiceNo,
    includeChallanNo,
    includeQuotationNo,
    includeRequisitionNo,
    includePoNumber,
    rows,
    mergedRegions,
    cellFormats,
    autoSaveEnabled,
    currentDocId,
    vatPercent,
    transportationFee
  ]);

  // Active sheet draft persistence to prevent any loss of data on reload or refresh - strictly isolated by active company
  useEffect(() => {
    if (typeof window === "undefined") return;
    try {
      const draftPayload = {
        docType,
        dateVal,
        messers,
        address,
        vesselName,
        portBerth,
        includeVesselName,
        includePortBerth,
        currency,
        includeDiscount,
        discountType,
        discountValue,
        discountPercent: discountType === "percentage" ? (parseFloat(discountValue) || 0) : 0,
        challanNo,
        requisitionNo,
        invoiceNo,
        poNumber,
        quotationNo,
        includeInvoiceNo,
        includeChallanNo,
        includeQuotationNo,
        includeRequisitionNo,
        includePoNumber,
        vatPercent,
        transportationFee,
        currentDocId: currentDocId || null,
        rows,
        mergedRegions,
        cellFormats
      };
      localStorage.setItem(getDraftStorageKey(activeCompany), JSON.stringify(draftPayload));
    } catch (e) {
      // Ignore localStorage storage quota errors
    }
  }, [
    activeCompany,
    docType,
    dateVal,
    messers,
    address,
    vesselName,
    portBerth,
    includeVesselName,
    includePortBerth,
    currency,
    includeDiscount,
    discountType,
    discountValue,
    challanNo,
    requisitionNo,
    invoiceNo,
    poNumber,
    quotationNo,
    includeInvoiceNo,
    includeChallanNo,
    includeQuotationNo,
    includeRequisitionNo,
    includePoNumber,
    rows,
    mergedRegions,
    cellFormats,
    currentDocId,
    vatPercent,
    transportationFee
  ]);

  // Adjust textarea heights dynamically based on content
  useEffect(() => {
    const textareas = document.querySelectorAll("textarea[data-row]");
    textareas.forEach((ta: any) => {
      ta.style.height = "auto";
      ta.style.height = `${ta.scrollHeight}px`;
    });
  }, [rows]);

  useEffect(() => {
    const handleDismiss = () => setContextMenu(null);
    const handleGlobalMouseUp = () => setIsSelecting(false);
    document.addEventListener("click", handleDismiss);
    window.addEventListener("scroll", handleDismiss, true);
    window.addEventListener("mouseup", handleGlobalMouseUp);
    return () => {
      document.removeEventListener("click", handleDismiss);
      window.removeEventListener("scroll", handleDismiss, true);
      window.removeEventListener("mouseup", handleGlobalMouseUp);
    };
  }, []);

  const handleRowChange = (index: number, field: keyof QuotationRow, value: string) => {
    const now = Date.now();
    if (now - lastTypingTimeRef.current > 650) {
      recordChange();
    }
    lastTypingTimeRef.current = now;

    setRows((prevRows) => {
      const updated = [...prevRows];
      const targetRow = { ...updated[index] };
      
      if (field === "desc" || field === "unit" || field === "qty" || field === "price") {
        (targetRow as any)[field] = value;
      }

      const cleanQty = stripHtml(String(targetRow.qty || ""));
      const cleanPrice = stripHtml(String(targetRow.price || ""));
      const q = parseNumericInput(cleanQty);
      const p = parseNumericInput(cleanPrice);
      targetRow.amount = docType === "challan" ? 0 : q * p;

      updated[index] = targetRow;
      return updated;
    });
  };

  const addRow = () => {
    addRows(1);
  };

  const MAX_LINES_LIMIT = 1000;

  const addRows = (count: number) => {
    if (rows.length >= MAX_LINES_LIMIT) {
      showToast(`Maximum limit of ${MAX_LINES_LIMIT.toLocaleString()} lines reached`);
      return;
    }
    const availableSlots = MAX_LINES_LIMIT - rows.length;
    const requestedCount = Math.max(1, count);
    const qtyToAdd = Math.min(requestedCount, availableSlots);

    if (qtyToAdd <= 0) {
      showToast(`Maximum limit of ${MAX_LINES_LIMIT.toLocaleString()} lines reached`);
      return;
    }

    recordChange();
    setRows((prevRows) => {
      const newRows = [...prevRows];
      const startIdx = newRows.length;
      for (let i = 1; i <= qtyToAdd; i++) {
        newRows.push({
          sl: startIdx + i,
          desc: "",
          qty: "",
          unit: "",
          price: "",
          amount: 0,
        });
      }
      return newRows;
    });

    if (qtyToAdd < requestedCount) {
      showToast(`Added ${qtyToAdd} lines (reached ${MAX_LINES_LIMIT.toLocaleString()} line limit)`);
    } else {
      showToast(`Added ${qtyToAdd} line${qtyToAdd > 1 ? "s" : ""}`);
    }
  };

  const setExactTotalRows = (targetCount: number) => {
    const count = Math.max(1, Math.min(MAX_LINES_LIMIT, targetCount));
    recordChange();
    setRows((prevRows) => {
      if (count === prevRows.length) return prevRows;
      if (count > prevRows.length) {
        const needed = count - prevRows.length;
        const newRows = [...prevRows];
        const startIdx = newRows.length;
        for (let i = 1; i <= needed; i++) {
          newRows.push({
            sl: startIdx + i,
            desc: "",
            qty: "",
            unit: "",
            price: "",
            amount: 0,
          });
        }
        showToast(`Sheet resized to ${count} lines (+${needed} lines added)`);
        return newRows;
      } else {
        const removed = prevRows.slice(count);
        const hasData = removed.some(r => r.desc.trim() || r.qty.trim() || r.unit.trim() || r.price.trim());
        if (hasData) {
          const ok = window.confirm(`Trimming down to ${count} lines will delete rows containing data. Are you sure you want to proceed?`);
          if (!ok) return prevRows;
        }
        showToast(`Sheet resized to ${count} lines`);
        return prevRows.slice(0, count).map((r, i) => ({ ...r, sl: i + 1 }));
      }
    });
  };

  const trimTrailingBlankRows = () => {
    recordChange();
    setRows((prevRows) => {
      let lastNonEmptyIndex = -1;
      for (let i = prevRows.length - 1; i >= 0; i--) {
        const r = prevRows[i];
        if (r.desc.trim() !== "" || r.qty.trim() !== "" || r.unit.trim() !== "" || r.price.trim() !== "") {
          lastNonEmptyIndex = i;
          break;
        }
      }
      if (lastNonEmptyIndex === -1) {
        showToast("Sheet has no filled rows");
        return prevRows.slice(0, 10);
      }
      // Keep at least up to last content row, minimum 5 rows for aesthetics
      const newCount = Math.max(lastNonEmptyIndex + 1, 5);
      const trimmed = prevRows.slice(0, newCount);
      showToast(`Cleaned blank lines at bottom (${trimmed.length} total lines remaining)`);
      return trimmed.map((r, i) => ({ ...r, sl: i + 1 }));
    });
  };

  const removeRows = (count: number = 1) => {
    recordChange();
    setRows((prevRows) => {
      if (prevRows.length <= 1) return prevRows;
      const keepCount = Math.max(1, prevRows.length - count);
      return prevRows.slice(0, keepCount);
    });
  };

  const removeRow = () => {
    removeRows(1);
  };

  const insertRow = (index: number, position: 'above' | 'below') => {
    insertMultipleRows(index, position, 1);
  };

  const insertMultipleRows = (index: number, position: 'above' | 'below', count: number = 1) => {
    if (rows.length >= MAX_LINES_LIMIT) {
      showToast(`Maximum limit of ${MAX_LINES_LIMIT.toLocaleString()} lines reached`);
      return;
    }
    const availableSlots = MAX_LINES_LIMIT - rows.length;
    const qty = Math.min(Math.max(1, count), availableSlots);
    if (qty <= 0) {
      showToast(`Maximum limit of ${MAX_LINES_LIMIT.toLocaleString()} lines reached`);
      return;
    }
    const insertAt = position === 'above' ? index : index + 1;
    recordChange();
    setRows((prevRows) => {
      const updated = [...prevRows];
      const newItems: QuotationRow[] = [];
      for (let i = 0; i < qty; i++) {
        newItems.push({
          sl: 0,
          desc: "",
          qty: "",
          unit: "",
          price: "",
          amount: 0,
        });
      }
      updated.splice(insertAt, 0, ...newItems);
      return updated.slice(0, MAX_LINES_LIMIT).map((r, i) => ({
        ...r,
        sl: i + 1
      }));
    });

    setMergedRegions((prevRegions) =>
      prevRegions.map((region) => {
        if (region.startRow >= insertAt) {
          return { ...region, startRow: region.startRow + qty, endRow: region.endRow + qty };
        }
        if (region.endRow >= insertAt) {
          return { ...region, endRow: region.endRow + qty };
        }
        return region;
      })
    );

    showToast(`Inserted ${qty} line${qty > 1 ? "s" : ""} ${position} line ${index + 1}`);
  };

  const handleImportFromExcel = (newRows: QuotationRow[], mode: "replace" | "append" | "insert_at", insertIndex: number = 0) => {
    if (newRows.length === 0) return;
    recordChange();
    setRows((prevRows) => {
      let result: QuotationRow[] = [];
      if (mode === "replace") {
        result = newRows.slice(0, MAX_LINES_LIMIT);
        // Ensure at least 20 blank rows if fewer
        while (result.length < Math.min(20, MAX_LINES_LIMIT)) {
          result.push({
            sl: result.length + 1,
            desc: "",
            qty: "",
            unit: "",
            price: "",
            amount: 0,
          });
        }
        setMergedRegions([]);
      } else if (mode === "append") {
        const isCurrentSheetEmpty = prevRows.every(r => !r.desc.trim() && !r.qty.trim() && !r.unit.trim() && !r.price.trim());
        if (isCurrentSheetEmpty && prevRows.length <= 20) {
          result = newRows.slice(0, MAX_LINES_LIMIT);
          while (result.length < Math.min(20, MAX_LINES_LIMIT)) {
            result.push({
              sl: result.length + 1,
              desc: "",
              qty: "",
              unit: "",
              price: "",
              amount: 0,
            });
          }
        } else {
          result = [...prevRows, ...newRows].slice(0, MAX_LINES_LIMIT);
        }
      } else if (mode === "insert_at") {
        const updated = [...prevRows];
        updated.splice(insertIndex, 0, ...newRows);
        result = updated.slice(0, MAX_LINES_LIMIT);
      }
      showToast(`Imported ${Math.min(newRows.length, MAX_LINES_LIMIT)} lines from Excel successfully!`);
      return result.map((r, i) => ({ ...r, sl: i + 1 }));
    });
  };

  const deleteSpecificRow = (index: number) => {
    recordChange();
    setRows((prevRows) => {
      if (prevRows.length <= 1) {
        return [{
          sl: 1,
          desc: "",
          qty: "",
          unit: "",
          price: "",
          amount: 0,
        }];
      }
      const updated = prevRows.filter((_, i) => i !== index);
      return updated.map((r, i) => ({
        ...r,
        sl: i + 1
      }));
    });

    setMergedRegions((prevRegions) =>
      prevRegions
        .map((region) => {
          let { startRow, endRow } = region;
          if (startRow > index) startRow -= 1;
          if (endRow >= index) endRow -= 1;
          return { ...region, startRow, endRow };
        })
        .filter((region) => region.endRow >= region.startRow)
    );
  };

  const clearSpecificRow = (index: number) => {
    recordChange();
    setRows((prevRows) => {
      const updated = [...prevRows];
      updated[index] = {
        sl: index + 1,
        desc: "",
        qty: "",
        unit: "",
        price: "",
        amount: 0,
      };
      return updated;
    });
  };

  const COLUMN_FIELD_BY_INDEX: Record<number, "desc" | "qty" | "unit" | "price" | null> = {
    [-1]: null,
    0: "desc",
    1: "qty",
    2: "unit",
    3: "price",
    4: null,
  };

  const getMergeRegionAt = (rowIndex: number, colIndex: number): MergedRegion | undefined => {
    return mergedRegions.find(
      (m) =>
        rowIndex >= m.startRow &&
        rowIndex <= m.endRow &&
        colIndex >= m.startCol &&
        colIndex <= m.endCol
    );
  };

  const getMergeInfo = (rowIndex: number, colIndex: number) => {
    const region = getMergeRegionAt(rowIndex, colIndex);
    if (!region) return { region: undefined, isAnchor: false };
    const isAnchor = rowIndex === region.startRow && colIndex === region.startCol;
    return { region, isAnchor };
  };

  const rangesOverlap = (a: MergedRegion, b: { startRow: number; endRow: number; startCol: number; endCol: number }) => {
    return a.startRow <= b.endRow && a.endRow >= b.startRow && a.startCol <= b.endCol && a.endCol >= b.startCol;
  };

  const mergeSelectedRange = () => {
    if (!selectionStart || !selectionEnd) return;

    const startRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
    const endRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
    const startCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
    const endCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);

    if (startRow === endRow && startCol === endCol) return;

    const candidateRegion = { startRow, endRow, startCol, endCol };
    const overlapping = mergedRegions.find((m) => rangesOverlap(m, candidateRegion));
    if (overlapping) {
      window.alert("Part of this selection is already merged. Unmerge it first, then try again.");
      return;
    }

    recordChange();
    setRows((prevRows) => {
      const updated = prevRows.map((r) => ({ ...r }));
      const pieces: string[] = [];

      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const field = COLUMN_FIELD_BY_INDEX[c];
          if (field && updated[r]) {
            const val = String(updated[r][field] ?? "").trim();
            if (val !== "") pieces.push(val);
          }
        }
      }

      const combined = pieces.join(" ");

      for (let r = startRow; r <= endRow; r++) {
        if (!updated[r]) continue;
        for (let c = startCol; c <= endCol; c++) {
          const field = COLUMN_FIELD_BY_INDEX[c];
          if (!field) continue;
          if (r === startRow && c === startCol) {
            (updated[r] as any)[field] = combined;
          } else {
            (updated[r] as any)[field] = "";
          }
        }
        const q = parseNumericInput(updated[r].qty);
        const p = parseNumericInput(updated[r].price);
        updated[r].amount = docType === "challan" ? 0 : q * p;
      }
      return updated;
    });

    const newRegion: MergedRegion = {
      id: generateUUID(),
      startRow,
      endRow,
      startCol,
      endCol,
    };
    setMergedRegions((prev) => [...prev, newRegion]);

    setSelectionStart({ rowIndex: startRow, colIndex: startCol });
    setSelectionEnd({ rowIndex: endRow, colIndex: endCol });
    setSelectedCell({ rowIndex: startRow, colIndex: startCol });
    setSelectedRowIndex(startRow);
  };

  const unmergeRegionAt = (rowIndex: number, colIndex: number) => {
    const region = getMergeRegionAt(rowIndex, colIndex);
    if (!region) return;
    recordChange();
    setMergedRegions((prev) => prev.filter((m) => m.id !== region.id));
  };

  const hasRangeSelectionMatchingRegion = (region: MergedRegion) => {
    if (!selectionStart || !selectionEnd) return false;
    const startRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
    const endRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
    const startCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
    const endCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);
    return (
      startRow === region.startRow &&
      endRow === region.endRow &&
      startCol === region.startCol &&
      endCol === region.endCol
    );
  };

  const toggleMergeSelectedRange = () => {
    if (!selectionStart || !selectionEnd) return;
    const { rowIndex, colIndex } = selectionStart;
    const existing = getMergeRegionAt(rowIndex, colIndex);
    if (existing && hasRangeSelectionMatchingRegion(existing)) {
      unmergeRegionAt(rowIndex, colIndex);
    } else if (existing) {
      unmergeRegionAt(rowIndex, colIndex);
    } else {
      mergeSelectedRange();
    }
  };

  const moveRow = (index: number, direction: 'up' | 'down') => {
    recordChange();
    setRows((prevRows) => {
      if (direction === 'up' && index === 0) return prevRows;
      if (direction === 'down' && index === prevRows.length - 1) return prevRows;

      const updated = [...prevRows];
      const swapIndex = direction === 'up' ? index - 1 : index + 1;
      
      const temp = updated[index];
      updated[index] = updated[swapIndex];
      updated[swapIndex] = temp;

      return updated.map((r, i) => ({
        ...r,
        sl: i + 1
      }));
    });
  };

  const isCellSelected = (rowIndex: number, colIndex: number) => {
    if (!selectionStart || !selectionEnd) {
      return selectedCell?.rowIndex === rowIndex && selectedCell?.colIndex === colIndex;
    }
    const minRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
    const maxRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
    const minCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
    const maxCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);

    return rowIndex >= minRow && rowIndex <= maxRow && colIndex >= minCol && colIndex <= maxCol;
  };

  const hasRangeSelection = () => {
    if (!selectionStart || !selectionEnd) return false;
    return selectionStart.rowIndex !== selectionEnd.rowIndex || selectionStart.colIndex !== selectionEnd.colIndex;
  };

  const handleCellMouseDown = (e: React.MouseEvent, rowIndex: number, colIndex: number) => {
    if (e.button !== 0) return;

    // Check if user clicked an editable element or within one
    const target = e.target as HTMLElement;
    const isInsideEditable = target && (target.isContentEditable || !!target.closest("[contenteditable='true']") || target.tagName === "INPUT");

    setSelectedCell({ rowIndex, colIndex });
    setSelectedRowIndex(rowIndex);
    setSelectionStart({ rowIndex, colIndex });
    setSelectionEnd({ rowIndex, colIndex });

    if (isInsideEditable) {
      // User is selecting text or placing cursor in an editable cell!
      // Do NOT preventDefault, do NOT blur, and do not start multi-cell drag
      setIsSelecting(false);
      return;
    }

    const isActive = selectedCell?.rowIndex === rowIndex && selectedCell?.colIndex === colIndex;
    setIsSelecting(true);

    if (!isActive) {
      if (document.activeElement instanceof HTMLElement) {
        document.activeElement.blur();
      }
      e.preventDefault();
    }
  };

  const handleCellMouseEnter = (rowIndex: number, colIndex: number) => {
    if (isSelecting) {
      setSelectionEnd({ rowIndex, colIndex });
    }
  };

  const handleCellMouseUp = (e: React.MouseEvent, rowIndex: number, colIndex: number) => {
    setIsSelecting(false);

    const target = e.target as HTMLElement;
    const isInsideEditable = target && (target.isContentEditable || !!target.closest("[contenteditable='true']") || target.tagName === "INPUT");
    if (isInsideEditable) {
      // Preserve active text selection in editable cells
      return;
    }

    const isSingleCell = selectionStart && selectionStart.rowIndex === rowIndex && selectionStart.colIndex === colIndex;
    if (isSingleCell && !hasRangeSelection()) {
      if (colIndex >= 0 && colIndex <= 3) {
        const textarea = document.querySelector(`[data-row="${rowIndex}"][data-col="${colIndex}"]`) as HTMLTextAreaElement | null;
        if (textarea) {
          textarea.focus();
        }
      }
    }
  };

  const getCellClassName = (rowIndex: number, colIndex: number, baseClasses: string) => {
    const isSelected = isCellSelected(rowIndex, colIndex);
    const isActive = selectedCell?.rowIndex === rowIndex && selectedCell?.colIndex === colIndex;
    
    let highlightClass = "";
    if (isActive) {
      highlightClass = "outline outline-2 outline-indigo-600 outline-offset-[-2px] bg-indigo-50/15 print:!outline-none print:!bg-transparent print:!shadow-none z-10 print:z-auto relative";
    } else if (isSelected) {
      highlightClass = "outline outline-1 outline-indigo-400 outline-offset-[-1px] bg-indigo-50/25 print:!outline-none print:!bg-transparent print:!shadow-none z-10 print:z-auto relative shadow-3xs";
    }
    
    return `${baseClasses} ${highlightClass}`;
  };

  const COLUMN_NAMES: Record<number, string> = {
    [-1]: "SL",
    0: "Description",
    1: "Qty",
    2: "Unit",
    3: "Unit Price",
    4: "Amount",
  };

  const activeCellFormat: CellFormat = React.useMemo(() => {
    if (selectedCell) {
      const key = `${selectedCell.rowIndex}_${selectedCell.colIndex}`;
      return cellFormats[key] || {};
    }
    if (selectionStart) {
      const key = `${selectionStart.rowIndex}_${selectionStart.colIndex}`;
      return cellFormats[key] || {};
    }
    return {};
  }, [selectedCell, selectionStart, cellFormats]);

  const selectionSummary = React.useMemo(() => {
    if (selectionStart && selectionEnd && (selectionStart.rowIndex !== selectionEnd.rowIndex || selectionStart.colIndex !== selectionEnd.colIndex)) {
      const minRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
      const maxRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
      const minCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
      const maxCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);
      const totalCells = (maxRow - minRow + 1) * (maxCol - minCol + 1);
      return `Rows ${minRow + 1}–${maxRow + 1}, Cols ${COLUMN_NAMES[minCol] || minCol}–${COLUMN_NAMES[maxCol] || maxCol} (${totalCells} cells)`;
    }
    if (selectedCell) {
      return `Row ${selectedCell.rowIndex + 1}, ${COLUMN_NAMES[selectedCell.colIndex] || "Cell"}`;
    }
    if (selectedRowIndex >= 0) {
      return `Row ${selectedRowIndex + 1} (Entire Row)`;
    }
    return undefined;
  }, [selectedCell, selectionStart, selectionEnd, selectedRowIndex]);

  const handleApplyFormat = (formatUpdate: Partial<CellFormat>) => {
    // If the user has highlighted text or is editing an inline contenteditable element,
    // apply formatting strictly using document.execCommand to the selected text range,
    // preventing the styles from bleeding into the entire container!
    if (hasActiveSelectionInEditable()) {
      let handled = false;
      if (formatUpdate.bold !== undefined) {
        applyInlineFormatting("bold");
        handled = true;
      }
      if (formatUpdate.italic !== undefined) {
        applyInlineFormatting("italic");
        handled = true;
      }
      if (formatUpdate.underline !== undefined) {
        applyInlineFormatting("underline", formatUpdate.underline);
        handled = true;
      }
      if (formatUpdate.color !== undefined) {
        applyInlineFormatting("fontColor", formatUpdate.color);
        handled = true;
      }
      if (formatUpdate.bgColor !== undefined) {
        applyInlineFormatting("highlight", formatUpdate.bgColor);
        handled = true;
      }
      if (formatUpdate.fontSize !== undefined) {
        applyInlineFormatting("fontSize", formatUpdate.fontSize);
        handled = true;
      }
      if (formatUpdate.fontFamily !== undefined) {
        applyInlineFormatting("fontFamily", formatUpdate.fontFamily);
        handled = true;
      }
      if (handled) {
        return;
      }
    }

    recordChange();
    setCellFormats((prev) => {
      const next = { ...prev };

      if (selectionStart && selectionEnd) {
        const minRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
        const maxRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
        const minCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
        const maxCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);

        for (let r = minRow; r <= maxRow; r++) {
          for (let c = minCol; c <= maxCol; c++) {
            const key = `${r}_${c}`;
            next[key] = {
              ...(next[key] || {}),
              ...formatUpdate,
            };
          }
        }
      } else if (selectedCell) {
        const key = `${selectedCell.rowIndex}_${selectedCell.colIndex}`;
        next[key] = {
          ...(next[key] || {}),
          ...formatUpdate,
        };
      } else if (selectedRowIndex >= 0) {
        [-1, 0, 1, 2, 3, 4].forEach((c) => {
          const key = `${selectedRowIndex}_${c}`;
          next[key] = {
            ...(next[key] || {}),
            ...formatUpdate,
          };
        });
      }

      return next;
    });
  };

  const handleApplyBorderPreset = (preset: string) => {
    recordChange();
    setCellFormats((prev) => {
      const next = { ...prev };
      const minRow = selectionStart && selectionEnd ? Math.min(selectionStart.rowIndex, selectionEnd.rowIndex) : (selectedCell ? selectedCell.rowIndex : selectedRowIndex);
      const maxRow = selectionStart && selectionEnd ? Math.max(selectionStart.rowIndex, selectionEnd.rowIndex) : (selectedCell ? selectedCell.rowIndex : selectedRowIndex);
      const minCol = selectionStart && selectionEnd ? Math.min(selectionStart.colIndex, selectionEnd.colIndex) : (selectedCell ? selectedCell.colIndex : -1);
      const maxCol = selectionStart && selectionEnd ? Math.max(selectionStart.colIndex, selectionEnd.colIndex) : (selectedCell ? selectedCell.colIndex : 4);

      for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
          const key = `${r}_${c}`;
          const existing = next[key] || {};
          let borders: CellBorders = { ...(existing.borders || {}) };

          const isTop = r === minRow;
          const isBottom = r === maxRow;
          const isLeft = c === minCol;
          const isRight = c === maxCol;

          switch (preset) {
            case "all":
              borders = {
                top: "1px solid #94a3b8",
                bottom: "1px solid #94a3b8",
                left: "1px solid #94a3b8",
                right: "1px solid #94a3b8",
              };
              break;
            case "none":
              borders = {
                top: "none",
                bottom: "none",
                left: "none",
                right: "none",
              };
              break;
            case "outside":
              if (isTop) borders.top = "1px solid #94a3b8";
              if (isBottom) borders.bottom = "1px solid #94a3b8";
              if (isLeft) borders.left = "1px solid #94a3b8";
              if (isRight) borders.right = "1px solid #94a3b8";
              break;
            case "thick_outside":
              if (isTop) borders.top = "2px solid #94a3b8";
              if (isBottom) borders.bottom = "2px solid #94a3b8";
              if (isLeft) borders.left = "2px solid #94a3b8";
              if (isRight) borders.right = "2px solid #94a3b8";
              break;
            case "bottom":
              if (isBottom) borders.bottom = "1px solid #94a3b8";
              break;
            case "top":
              if (isTop) borders.top = "1px solid #94a3b8";
              break;
            case "left":
              if (isLeft) borders.left = "1px solid #94a3b8";
              break;
            case "right":
              if (isRight) borders.right = "1px solid #94a3b8";
              break;
            case "thick_bottom":
              if (isBottom) borders.bottom = "2px solid #94a3b8";
              break;
            case "bottom_double":
              if (isBottom) borders.bottom = "3px double #94a3b8";
              break;
            case "top_and_bottom":
              if (isTop) borders.top = "1px solid #94a3b8";
              if (isBottom) borders.bottom = "1px solid #94a3b8";
              break;
            case "top_and_thick_bottom":
              if (isTop) borders.top = "1px solid #94a3b8";
              if (isBottom) borders.bottom = "2px solid #94a3b8";
              break;
            case "top_and_double_bottom":
              if (isTop) borders.top = "1px solid #94a3b8";
              if (isBottom) borders.bottom = "3px double #94a3b8";
              break;
          }

          next[key] = {
            ...existing,
            borders,
          };
        }
      }

      return next;
    });
  };

  const getCellStyle = (rowIndex: number, colIndex: number): React.CSSProperties => {
    const key = `${rowIndex}_${colIndex}`;
    const fmt = cellFormats[key];
    if (!fmt) return {};

    const style: React.CSSProperties = {};

    if (fmt.fontFamily) style.fontFamily = fmt.fontFamily;
    if (fmt.fontSize) style.fontSize = `${fmt.fontSize}pt`;
    if (fmt.bold !== undefined) style.fontWeight = fmt.bold ? "bold" : "normal";
    if (fmt.italic !== undefined) style.fontStyle = fmt.italic ? "italic" : "normal";
    if (fmt.underline) {
      if (fmt.underline === "double") {
        style.textDecoration = "underline";
        style.textDecorationStyle = "double";
      } else if (fmt.underline === "single") {
        style.textDecoration = "underline";
      } else {
        style.textDecoration = "none";
      }
    }
    if (fmt.align) style.textAlign = fmt.align;
    style.verticalAlign = fmt.valign || "middle";
    if (fmt.color) style.color = fmt.color;
    if (fmt.bgColor) style.backgroundColor = fmt.bgColor;
    if (fmt.indent) {
      style.paddingLeft = `${fmt.indent * 8}px`;
      style.textIndent = `${fmt.indent * 8}px`;
    }
    if (fmt.orientation && fmt.orientation !== "horizontal") {
      switch (fmt.orientation) {
        case "angle-up":
          style.transform = "rotate(-45deg)";
          style.display = "inline-block";
          break;
        case "angle-down":
          style.transform = "rotate(45deg)";
          style.display = "inline-block";
          break;
        case "vertical":
          style.writingMode = "vertical-rl";
          break;
        case "rotate-up":
          style.transform = "rotate(-90deg)";
          style.display = "inline-block";
          break;
        case "rotate-down":
          style.transform = "rotate(90deg)";
          style.display = "inline-block";
          break;
      }
    }

    if (fmt.borders) {
      if (fmt.borders.top) style.borderTop = fmt.borders.top;
      if (fmt.borders.bottom) style.borderBottom = fmt.borders.bottom;
      if (fmt.borders.left) style.borderLeft = fmt.borders.left;
      if (fmt.borders.right) style.borderRight = fmt.borders.right;
    }

    return style;
  };

  useEffect(() => {
    const handleGlobalShortcuts = (e: KeyboardEvent) => {
      if (!e.ctrlKey && !e.metaKey) return;

      const key = e.key.toLowerCase();
      const isShift = e.shiftKey;

      // Undo shortcut: Ctrl+Z (or Cmd+Z)
      if (key === "z" && !isShift) {
        e.preventDefault();
        e.stopPropagation();
        handleUndo();
        return;
      }

      // Redo shortcut: Ctrl+Y or Ctrl+Shift+Z (or Cmd+Y / Cmd+Shift+Z)
      if (key === "y" || (key === "z" && isShift)) {
        e.preventDefault();
        e.stopPropagation();
        handleRedo();
        return;
      }

      // Copy shortcut: Ctrl+C for selected cells or range
      if (key === "c") {
        const textSelection = window.getSelection()?.toString();
        // If user is selecting specific text in an editable cell or input, let browser copy that text
        if (textSelection && textSelection.trim().length > 0) {
          return;
        }

        const minRow = selectionStart && selectionEnd ? Math.min(selectionStart.rowIndex, selectionEnd.rowIndex) : (selectedCell ? selectedCell.rowIndex : selectedRowIndex);
        const maxRow = selectionStart && selectionEnd ? Math.max(selectionStart.rowIndex, selectionEnd.rowIndex) : (selectedCell ? selectedCell.rowIndex : selectedRowIndex);
        const minCol = selectionStart && selectionEnd ? Math.min(selectionStart.colIndex, selectionEnd.colIndex) : (selectedCell ? selectedCell.colIndex : 0);
        const maxCol = selectionStart && selectionEnd ? Math.max(selectionStart.colIndex, selectionEnd.colIndex) : (selectedCell ? selectedCell.colIndex : 3);

        if (minRow >= 0 && maxRow < rows.length) {
          const tsvRows: string[] = [];
          for (let r = minRow; r <= maxRow; r++) {
            const rowData = rows[r];
            if (!rowData) continue;
            const cells: string[] = [];
            for (let c = minCol; c <= maxCol; c++) {
              if (c === -1) cells.push(String(rowData.sl || r + 1));
              else if (c === 0) cells.push(stripHtml(rowData.desc || ""));
              else if (c === 1) cells.push(stripHtml(rowData.qty || ""));
              else if (c === 2) cells.push(stripHtml(rowData.unit || ""));
              else if (c === 3) cells.push(stripHtml(rowData.price || ""));
              else if (c === 4) cells.push(String(rowData.amount || 0));
            }
            tsvRows.push(cells.join("\t"));
          }
          const tsvString = tsvRows.join("\r\n");
          if (tsvString && navigator.clipboard?.writeText) {
            e.preventDefault();
            navigator.clipboard.writeText(tsvString);
            showToast(`Copied ${tsvRows.length} row${tsvRows.length > 1 ? "s" : ""} to clipboard`);
            return;
          }
        }
      }

      const target = e.target as HTMLElement;
      const isContentEditable = target?.isContentEditable || !!target?.closest?.("[contenteditable='true']");
      const isSheetInput = target?.closest?.(".sheet") || target?.hasAttribute?.("data-row");

      if (!selectedCell && !selectionStart && !isSheetInput && !isContentEditable) return;

      if (key === "b") {
        e.preventDefault();
        recordChange();
        if (isContentEditable || hasActiveSelectionInEditable()) {
          applyInlineFormatting("bold");
        } else {
          handleApplyFormat({ bold: !activeCellFormat.bold });
        }
      } else if (key === "i") {
        e.preventDefault();
        recordChange();
        if (isContentEditable || hasActiveSelectionInEditable()) {
          applyInlineFormatting("italic");
        } else {
          handleApplyFormat({ italic: !activeCellFormat.italic });
        }
      } else if (key === "u") {
        e.preventDefault();
        recordChange();
        if (isContentEditable || hasActiveSelectionInEditable()) {
          applyInlineFormatting("underline");
        } else {
          handleApplyFormat({ underline: activeCellFormat.underline === "single" ? "none" : "single" });
        }
      } else if (key === "l" && !e.shiftKey) {
        e.preventDefault();
        recordChange();
        handleApplyFormat({ align: "left" });
      } else if (key === "e" && !e.shiftKey) {
        e.preventDefault();
        recordChange();
        handleApplyFormat({ align: "center" });
      } else if (key === "r" && !e.shiftKey) {
        e.preventDefault();
        recordChange();
        handleApplyFormat({ align: "right" });
      } else if (e.shiftKey && (key === ">" || key === ".")) {
        e.preventDefault();
        recordChange();
        const currentSize = activeCellFormat.fontSize || 8.5;
        const newSize = Math.min(72, currentSize + 1);
        if (isContentEditable || hasActiveSelectionInEditable()) {
          applyInlineFormatting("fontSize", newSize);
        } else {
          handleApplyFormat({ fontSize: newSize });
        }
      } else if (e.shiftKey && (key === "<" || key === ",")) {
        e.preventDefault();
        recordChange();
        const currentSize = activeCellFormat.fontSize || 8.5;
        const newSize = Math.max(5, currentSize - 1);
        if (isContentEditable || hasActiveSelectionInEditable()) {
          applyInlineFormatting("fontSize", newSize);
        } else {
          handleApplyFormat({ fontSize: newSize });
        }
      }
    };

    window.addEventListener("keydown", handleGlobalShortcuts, { capture: true });
    return () => window.removeEventListener("keydown", handleGlobalShortcuts, { capture: true });
  }, [activeCellFormat, selectedCell, selectionStart, selectionEnd, selectedRowIndex, rows, mergedRegions, cellFormats, docType]);

  const handleCellClick = (rowIndex: number, colIndex: number) => {
    setSelectedRowIndex(rowIndex);
    setSelectedCell({ rowIndex, colIndex });
    setSelectionStart({ rowIndex, colIndex });
    setSelectionEnd({ rowIndex, colIndex });
    const textarea = document.querySelector(`[data-row="${rowIndex}"][data-col="${colIndex}"]`) as HTMLTextAreaElement | null;
    if (textarea) {
      textarea.focus();
    }
  };

  const clearSpecificCell = (rowIndex: number, colIndex: number) => {
    let field: "desc" | "qty" | "unit" | "price" | null = null;
    if (colIndex === 0) field = "desc";
    else if (colIndex === 1) field = "qty";
    else if (colIndex === 2) field = "unit";
    else if (colIndex === 3) field = "price";
    
    if (field) {
      handleRowChange(rowIndex, field, "");
    }
  };

  const handleCellContextMenu = (e: React.MouseEvent, idx: number, colIdx: number) => {
    e.preventDefault();
    const clickedInsideRange = isCellSelected(idx, colIdx);
    if (!clickedInsideRange) {
      setSelectionStart({ rowIndex: idx, colIndex: colIdx });
      setSelectionEnd({ rowIndex: idx, colIndex: colIdx });
    }
    setSelectedRowIndex(idx);
    setSelectedCell({ rowIndex: idx, colIndex: colIdx });
    
    let x = e.clientX;
    let y = e.clientY;
    const menuWidth = 260;
    const menuHeight = 320;
    
    if (x + menuWidth > window.innerWidth) x = window.innerWidth - menuWidth - 10;
    if (y + menuHeight > window.innerHeight) y = window.innerHeight - menuHeight - 10;
    if (x < 0) x = 10;
    if (y < 0) y = 10;

    setContextMenu({ visible: true, x, y, rowIndex: idx, colIndex: colIdx });
  };

  const rowsTotal = rows.reduce((sum, r) => sum + r.amount, 0);
  const parsedVatPercent = parseNumericInput(vatPercent);
  const parsedTransportationFee = parseNumericInput(transportationFee);
  const parsedDiscountValue = parseNumericInput(discountValue);

  let discountAmount = 0;
  if (docType === "invoice" && includeDiscount && parsedDiscountValue > 0) {
    if (discountType === "percentage") {
      discountAmount = (rowsTotal * parsedDiscountValue) / 100;
    } else {
      discountAmount = parsedDiscountValue;
    }
  }
  discountAmount = Math.min(rowsTotal, Math.max(0, discountAmount));
  const netAfterDiscount = Math.max(0, rowsTotal - discountAmount);
  const vatAmount = docType === "invoice" ? (netAfterDiscount * parsedVatPercent) / 100 : 0;
  const grandTotal = docType === "invoice" 
    ? (netAfterDiscount + vatAmount + parsedTransportationFee) 
    : rowsTotal;
  const calculatedGrandTotal = docType === "challan" ? 0 : grandTotal;

  const handleKeyDown = (e: React.KeyboardEvent<HTMLElement>, rowIndex: number, colIndex: number) => {
    const { key } = e;
    let targetRow = rowIndex;
    let targetCol = colIndex;

    const targetEl = e.currentTarget;
    const isFormInput = targetEl instanceof HTMLInputElement || targetEl instanceof HTMLTextAreaElement;

    let cursorStart = 0;
    let cursorEnd = 0;
    let valueLength = 0;

    if (isFormInput) {
      cursorStart = targetEl.selectionStart ?? 0;
      cursorEnd = targetEl.selectionEnd ?? 0;
      valueLength = targetEl.value?.length ?? 0;
    } else {
      const sel = typeof window !== "undefined" ? window.getSelection() : null;
      if (sel && sel.rangeCount > 0) {
        const range = sel.getRangeAt(0);
        cursorStart = range.startOffset;
        cursorEnd = range.endOffset;
      }
      valueLength = (targetEl.textContent || "").length;
    }

    if (key === "ArrowUp") {
      targetRow = rowIndex - 1;
    } else if (key === "ArrowDown") {
      targetRow = rowIndex + 1;
    } else if (key === "ArrowLeft") {
      if (cursorStart === 0 && cursorEnd === 0) {
        targetCol = colIndex - 1;
      } else {
        return;
      }
    } else if (key === "ArrowRight") {
      if (cursorStart === valueLength && cursorEnd === valueLength) {
        targetCol = colIndex + 1;
      } else {
        return;
      }
    } else if (key === "Enter") {
      if (e.shiftKey) {
        return;
      }
      targetRow = rowIndex + 1;
    } else {
      return;
    }

    e.preventDefault();
    const targetElement = document.querySelector(
      `[data-row="${targetRow}"][data-col="${targetCol}"]`
    ) as HTMLElement | null;

    if (targetElement) {
      targetElement.focus();
      if (typeof (targetElement as any).select === "function") {
        (targetElement as any).select();
      }
    }
  };

  const performPaste = (
    clipboardData: DataTransfer | null,
    startRowIndex: number,
    startColIndex: number,
    eventToPrevent?: { preventDefault: () => void }
  ) => {
    if (!clipboardData) return;

    const plainText = clipboardData.getData("text/plain") || clipboardData.getData("text") || "";
    const htmlText = clipboardData.getData("text/html") || "";

    if (!plainText && !htmlText) return;

    // Explicitly disable CSV comma splitting so text with commas isn't incorrectly split into columns
    const result = parseClipboardData({ text: plainText, html: htmlText }, { allowCsv: false });
    const parsedGrid = result.grid;

    if (!parsedGrid || parsedGrid.length === 0) return;

    // Single-cell paste (1 line, 1 column)
    if (parsedGrid.length === 1 && parsedGrid[0].length <= 1) {
      if (startColIndex >= 0) {
        // If event was triggered inside an active contenteditable element, allow native paste at cursor position
        if (eventToPrevent) {
          return;
        }
        // Pasted into cell via cell selection without active cursor
        const fieldMap = ["desc", "qty", "unit", "price"] as const;
        const field = fieldMap[startColIndex];
        if (field) {
          recordChange();
          handleRowChange(startRowIndex, field, cleanCellText(parsedGrid[0][0]));
        }
        return;
      }
    }

    // Multi-cell or multi-line paste (e.g. copied from Excel)
    if (eventToPrevent) {
      eventToPrevent.preventDefault();
    }

    recordChange();

    // Never drop the first row of user data!
    // Only skip row 0 if it is strictly a column header with NO numeric data in Qty/Price
    let dataRows = parsedGrid;
    if (result.hasHeader && parsedGrid.length > 1 && parsedGrid[0].length >= 2) {
      const firstRow = parsedGrid[0];
      const hasNumericData = firstRow.some((cell, idx) => {
        if (idx === 0) return false;
        const clean = cell.replace(/[,$\s]/g, "").trim();
        return /^\d+(\.\d+)?$/.test(clean) && clean.length > 0;
      });
      if (!hasNumericData) {
        const headerNames = ["description", "desc", "qty", "quantity", "unit", "price", "rate", "amount", "total", "sl", "s.no", "item #", "#"];
        const matches = firstRow.filter((c) => headerNames.includes(c.toLowerCase().trim()));
        if (matches.length >= Math.min(2, firstRow.length)) {
          dataRows = parsedGrid.slice(1);
        }
      }
    }
    if (dataRows.length === 0) return;

    const sample = dataRows.slice(0, 15);
    const maxCols = Math.max(...dataRows.map((r) => r.length));

    // Helper to check if a specific column across sample rows is predominantly numeric
    const isColNumeric = (colIndex: number): boolean => {
      let count = 0;
      let numericCount = 0;
      sample.forEach((r) => {
        if (r[colIndex] !== undefined && r[colIndex].trim().length > 0) {
          count++;
          const val = r[colIndex].replace(/[$€£,\s]/g, "").trim();
          if (/^-?\d+(\.\d+)?$/.test(val)) {
            numericCount++;
          }
        }
      });
      return count > 0 && numericCount / count >= 0.7;
    };

    // Detect if column 0 represents Serial Number (1, 2, 3...)
    const isColZeroSerial = (() => {
      if (result.hasSerialColumn) return true;
      if (sample.length === 0 || maxCols < 2) return false;

      let serialCount = 0;
      let col1TextCount = 0;

      sample.forEach((r) => {
        if (r.length >= 2) {
          const val0 = r[0].replace(/[.\-#\s]/g, "").trim();
          if (/^\d+$/.test(val0) && val0.length <= 5) {
            serialCount++;
          }
          if (/[a-zA-Z]/.test(r[1])) {
            col1TextCount++;
          }
        }
      });

      return serialCount >= Math.max(1, Math.ceil(sample.length * 0.6)) && col1TextCount >= 1;
    })();

    setRows((prevRows) => {
      const updated = [...prevRows];
      const targetTotalRows = startRowIndex + dataRows.length;

      // Expand table to ensure every single copied line gets its own row (up to MAX_LINES_LIMIT)
      while (updated.length < targetTotalRows && updated.length < MAX_LINES_LIMIT) {
        updated.push({
          sl: updated.length + 1,
          desc: "",
          qty: "",
          unit: "",
          price: "",
          amount: 0,
        });
      }

      dataRows.forEach((cols, rOffset) => {
        const rIndex = startRowIndex + rOffset;
        if (rIndex >= MAX_LINES_LIMIT) return;

        const targetRow = { ...updated[rIndex] };

        // Sub-column paste (User clicked directly on Qty, Unit, or Price column)
        if (startColIndex > 0) {
          if (startColIndex === 1) {
            // Started at Qty (Col 1)
            targetRow.qty = cleanCellText(cols[0] ?? "");
            if (cols.length === 2) {
              if (isColNumeric(1) && docType !== "challan") {
                targetRow.price = cleanCellText(cols[1] ?? "");
                targetRow.unit = "";
              } else {
                targetRow.unit = cleanCellText(cols[1] ?? "");
                if (docType !== "challan") targetRow.price = "";
              }
            } else if (cols.length >= 3) {
              targetRow.unit = cleanCellText(cols[1] ?? "");
              if (docType !== "challan") targetRow.price = cleanCellText(cols[2] ?? "");
            }
          } else if (startColIndex === 2) {
            // Started at Unit (Col 2)
            targetRow.unit = cleanCellText(cols[0] ?? "");
            if (cols.length >= 2 && docType !== "challan") {
              targetRow.price = cleanCellText(cols[1] ?? "");
            }
          } else if (startColIndex === 3) {
            // Started at Price (Col 3)
            if (docType !== "challan") {
              targetRow.price = cleanCellText(cols[0] ?? "");
            }
          }
        }
        // Pasted starting at SL or Description column (startColIndex <= 0)
        else if (isColZeroSerial) {
          // Copied Col 0 is SL! The remaining columns start at index 1:
          // [SL (0), Description (1), ...]
          const itemCols = cols.slice(1);
          targetRow.desc = cleanCellText(itemCols[0] ?? "");

          if (itemCols.length === 1) {
            targetRow.qty = "";
            targetRow.unit = "";
            if (docType !== "challan") targetRow.price = "";
          } else if (itemCols.length === 2) {
            // [SL, Desc, Qty]
            targetRow.qty = cleanCellText(itemCols[1] ?? "");
            targetRow.unit = "";
            if (docType !== "challan") targetRow.price = "";
          } else if (itemCols.length === 3) {
            // [SL, Desc, Qty, Unit] OR [SL, Desc, Qty, Price]
            targetRow.qty = cleanCellText(itemCols[1] ?? "");
            if (isColNumeric(3) && docType !== "challan") {
              targetRow.unit = "";
              targetRow.price = cleanCellText(itemCols[2] ?? "");
            } else {
              targetRow.unit = cleanCellText(itemCols[2] ?? "");
              if (docType !== "challan") targetRow.price = "";
            }
          } else if (itemCols.length === 4) {
            // [SL, Desc, Qty, Unit, Price] OR [SL, Desc, Qty, Price, Amount]
            targetRow.qty = cleanCellText(itemCols[1] ?? "");
            if (isColNumeric(3) && isColNumeric(4)) {
              // [SL, Desc, Qty, Price, Amount]
              targetRow.unit = "";
              if (docType !== "challan") targetRow.price = cleanCellText(itemCols[2] ?? "");
            } else {
              // [SL, Desc, Qty, Unit, Price]
              targetRow.unit = cleanCellText(itemCols[2] ?? "");
              if (docType !== "challan") targetRow.price = cleanCellText(itemCols[3] ?? "");
            }
          } else {
            // [SL, Desc, Qty, Unit, Price, Amount...]
            targetRow.qty = cleanCellText(itemCols[1] ?? "");
            targetRow.unit = cleanCellText(itemCols[2] ?? "");
            if (docType !== "challan") targetRow.price = cleanCellText(itemCols[3] ?? "");
          }
        }
        // Copied from Excel WITHOUT Serial Column (Col 0 is Description)
        else {
          if (maxCols === 1) {
            // Single column of items
            targetRow.desc = cleanCellText(cols[0] ?? "");
          } else if (maxCols === 2) {
            // [Desc, Qty]
            targetRow.desc = cleanCellText(cols[0] ?? "");
            targetRow.qty = cleanCellText(cols[1] ?? "");
            targetRow.unit = "";
            if (docType !== "challan") targetRow.price = "";
          } else if (maxCols === 3) {
            // [Desc, Qty, Unit] OR [Desc, Qty, Price]
            targetRow.desc = cleanCellText(cols[0] ?? "");
            targetRow.qty = cleanCellText(cols[1] ?? "");
            if (isColNumeric(2) && docType !== "challan") {
              targetRow.unit = "";
              targetRow.price = cleanCellText(cols[2] ?? "");
            } else {
              targetRow.unit = cleanCellText(cols[2] ?? "");
              if (docType !== "challan") targetRow.price = "";
            }
          } else if (maxCols === 4) {
            // [Desc, Qty, Unit, Price] OR [Desc, Qty, Price, Amount]
            targetRow.desc = cleanCellText(cols[0] ?? "");
            targetRow.qty = cleanCellText(cols[1] ?? "");
            if (isColNumeric(2) && isColNumeric(3)) {
              targetRow.unit = "";
              if (docType !== "challan") targetRow.price = cleanCellText(cols[2] ?? "");
            } else {
              targetRow.unit = cleanCellText(cols[2] ?? "");
              if (docType !== "challan") targetRow.price = cleanCellText(cols[3] ?? "");
            }
          } else {
            // 5+ columns: [Desc, Qty, Unit, Price, Amount...]
            targetRow.desc = cleanCellText(cols[0] ?? "");
            targetRow.qty = cleanCellText(cols[1] ?? "");
            targetRow.unit = cleanCellText(cols[2] ?? "");
            if (docType !== "challan") targetRow.price = cleanCellText(cols[3] ?? "");
          }
        }

        const cleanQty = stripHtml(String(targetRow.qty || ""));
        const cleanPrice = stripHtml(String(targetRow.price || ""));
        const q = parseNumericInput(cleanQty);
        const p = parseNumericInput(cleanPrice);
        targetRow.amount = docType === "challan" ? 0 : q * p;
        updated[rIndex] = targetRow;
      });

      return updated;
    });

    showToast(`Pasted ${dataRows.length} lines from Excel`);
  };

  const handlePaste = (
    e: React.ClipboardEvent<HTMLElement>,
    startRowIndex: number,
    startColIndex: number
  ) => {
    performPaste(e.clipboardData, startRowIndex, startColIndex, e);
  };

  // Global window paste listener: allows Ctrl+V to paste into the sheet even when a cell or row is highlighted
  useEffect(() => {
    const handleWindowPaste = (e: ClipboardEvent) => {
      const target = e.target as HTMLElement;
      // If user is editing a standard input or textarea, let default native paste proceed
      if (target?.tagName === "INPUT" || target?.tagName === "TEXTAREA") {
        return;
      }
      // If inside an active contenteditable cell, RichTextCell's onPaste handles it
      const isInsideEditable = target?.hasAttribute("contenteditable") || !!target?.closest?.("[contenteditable='true']");
      if (isInsideEditable) {
        return;
      }

      const startRow = selectedCell?.rowIndex ?? selectedRowIndex ?? 0;
      const startCol = selectedCell?.colIndex ?? 0;
      if (startRow >= 0 && e.clipboardData) {
        performPaste(e.clipboardData, startRow, startCol, e);
      }
    };

    window.addEventListener("paste", handleWindowPaste);
    return () => window.removeEventListener("paste", handleWindowPaste);
  }, [selectedCell, selectedRowIndex, rows, docType]);

  const unwrapAllDescriptions = () => {
    let fixedCount = 0;
    setRows((prev) => {
      const updated = prev.map((r) => {
        const original = r.desc || "";
        const cleaned = cleanCellText(original);
        if (cleaned !== original) {
          fixedCount++;
        }
        return {
          ...r,
          desc: cleaned,
          qty: cleanCellText(r.qty || ""),
          unit: cleanCellText(r.unit || ""),
          price: cleanCellText(r.price || ""),
        };
      });
      return updated;
    });

    if (fixedCount > 0) {
      showToast(`Cleaned & unwrapped line breaks across ${fixedCount} item(s)`);
    } else {
      showToast(`All description lines are continuous and clean`);
    }
  };

  const syncAllEditableFields = () => {
    if (typeof document === "undefined") return;
    const editables = document.querySelectorAll<HTMLElement>("[data-sync-id]");
    editables.forEach((el) => {
      const syncId = el.getAttribute("data-sync-id");
      if (syncId) {
        const printEl = document.getElementById(`print-${syncId}`);
        if (printEl) {
          printEl.innerHTML = el.innerHTML || "&nbsp;";
        }
      }
    });
  };

  useEffect(() => {
    const handleBeforePrint = () => {
      if (document.activeElement instanceof HTMLElement) {
        document.activeElement.blur();
      }
      setSelectedCell(null);
      setSelectionStart(null);
      setSelectionEnd(null);
      syncAllEditableFields();
    };
    window.addEventListener("beforeprint", handleBeforePrint);
    return () => window.removeEventListener("beforeprint", handleBeforePrint);
  }, []);

  const handlePrint = () => {
    if (document.activeElement instanceof HTMLElement) {
      document.activeElement.blur();
    }
    syncAllEditableFields();
    setSelectedCell(null);
    setSelectionStart(null);
    setSelectionEnd(null);
    setTimeout(() => {
      syncAllEditableFields();
      window.print();
    }, 120);
  };

  const handleDownloadPDF = async () => {
    const container = document.querySelector(".quotation-container");
    if (!container) return;
    
    const element = document.querySelector(".sheet") as HTMLElement | null;
    if (!element) return;
    
    if (document.activeElement instanceof HTMLElement) {
      document.activeElement.blur();
    }
    syncAllEditableFields();
    
    setIsGeneratingPDF(true);
    container.classList.add("is-generating-pdf");
    document.body.classList.add("is-generating-pdf");

    const originalGetComputedStyle = window.getComputedStyle;
    window.getComputedStyle = function (el: Element, pseudoElt?: string | null) {
      const style = originalGetComputedStyle.call(this, el, pseudoElt);
      return new Proxy(style, {
        get(target, prop, receiver) {
          if (prop === 'getPropertyValue') {
            return function(propertyName: string) {
              const val = target.getPropertyValue(propertyName);
              return typeof val === 'string' ? replaceOklchInCss(val) : val;
            };
          }
          const val = Reflect.get(target, prop, receiver);
          if (typeof val === 'string') {
            return replaceOklchInCss(val);
          }
          if (typeof val === 'function') {
            return val.bind(target);
          }
          return val;
        }
      }) as any;
    };
    
    const sheets = Array.from(document.styleSheets);
    let concatenatedCss = "";
    const disabledSheets: { sheet: CSSStyleSheet; wasDisabled: boolean }[] = [];

    for (const sheet of sheets) {
      try {
        const rules = Array.from(sheet.cssRules || []);
        const sheetCss = rules.map(rule => rule.cssText).join("\n");
        concatenatedCss += sheetCss + "\n";
        
        disabledSheets.push({ sheet, wasDisabled: sheet.disabled });
        sheet.disabled = true;
      } catch (e) {
        console.warn("Could not read stylesheet rules (possibly cross-origin):", e);
      }
    }

    const translatedCss = replaceOklchInCss(concatenatedCss);
    const tempStyle = document.createElement("style");
    tempStyle.id = "temp-pdf-colors";
    tempStyle.textContent = translatedCss;
    document.head.appendChild(tempStyle);

    const elementsWithInlineStyle = element.querySelectorAll("[style]");
    const inlineStylesBackup = new Map<HTMLElement, string>();
    
    const rootStyle = element.getAttribute("style");
    if (rootStyle && rootStyle.includes("oklch")) {
      inlineStylesBackup.set(element, rootStyle);
      element.setAttribute("style", replaceOklchInCss(rootStyle));
    }
    
    elementsWithInlineStyle.forEach((el) => {
      const htmlEl = el as HTMLElement;
      const styleAttr = htmlEl.getAttribute("style");
      if (styleAttr && styleAttr.includes("oklch")) {
        inlineStylesBackup.set(htmlEl, styleAttr);
        htmlEl.setAttribute("style", replaceOklchInCss(styleAttr));
      }
    });

    const filePrefix = docType === "challan" ? "Challan" : docType === "invoice" ? "Invoice" : "Quotation";
    const identifier = docType === "challan" ? (challanNo || "NEW") : docType === "invoice" ? (invoiceNo || "NEW") : (requisitionNo || "NEW");
    const filename = `${filePrefix}_${identifier.replace(/[\/\\?%*:|"<>\s]/g, "_")}.pdf`;
    
    const opt = {
      margin:       [4, 4, 4, 4] as [number, number, number, number],
      filename:     filename,
      image:        { type: "jpeg" as const, quality: 0.98 },
      html2canvas:  { 
        scale: 2.5,
        useCORS: true,
        logging: false,
        scrollY: 0,
        scrollX: 0
      },
      jsPDF:        { unit: "mm", format: "a4", orientation: "portrait" as const }
    };
    
    const cleanUpAfterPdf = () => {
      window.getComputedStyle = originalGetComputedStyle;

      disabledSheets.forEach(({ sheet, wasDisabled }) => {
        sheet.disabled = wasDisabled;
      });
      
      const addedStyle = document.getElementById("temp-pdf-colors");
      if (addedStyle) {
        addedStyle.remove();
      }
      
      inlineStylesBackup.forEach((originalStyle, htmlEl) => {
        htmlEl.setAttribute("style", originalStyle);
      });
      
      container.classList.remove("is-generating-pdf");
      document.body.classList.remove("is-generating-pdf");
      setIsGeneratingPDF(false);
    };

    try {
      // @ts-ignore
      const html2pdfModule = await import("html2pdf.js");
      const html2pdf: any = (html2pdfModule as any).default || html2pdfModule;
      html2pdf()
        .from(element)
        .set(opt)
        .save()
        .then(() => {
          cleanUpAfterPdf();
        })
        .catch((err: any) => {
          console.error("PDF generation error:", err);
          cleanUpAfterPdf();
        });
    } catch (err: any) {
      console.error("PDF module loading error:", err);
      cleanUpAfterPdf();
    }
  };

  const handleDownloadExcel = async () => {
    if (document.activeElement instanceof HTMLElement) {
      document.activeElement.blur();
    }
    syncAllEditableFields();
    setIsGeneratingExcel(true);
    try {
      await generateExcelDocument({
        currentCompany: COMPANY_PROFILES[activeCompany],
        docType,
        messers,
        address,
        vesselName,
        portBerth,
        includeVesselName,
        includePortBerth,
        invoiceNo,
        challanNo,
        quotationNo,
        requisitionNo,
        poNumber,
        includeInvoiceNo,
        includeChallanNo,
        includeQuotationNo,
        includeRequisitionNo,
        includePoNumber,
        dateVal,
        rows,
        includeDiscount,
        discountType,
        discountValue,
        vatPercent,
        transportationFee,
        currency,
        currencySymbol: currency === "USD" ? "$" : "Tk",
      });
      showToast("Excel spreadsheet generated successfully!");
    } catch (err: any) {
      console.error("Excel generation error:", err);
      showToast("Failed to generate Excel spreadsheet", "info");
    } finally {
      setIsGeneratingExcel(false);
    }
  };

  const safeSelectedRowIndex = Math.max(0, Math.min(selectedRowIndex, rows.length - 1));
  const GRID_COLUMNS = docType === "challan" ? [-1, 0, 1, 2] : [-1, 0, 1, 2, 3, 4];

  const handleClearFormatting = () => {
    recordChange();
    if (selectionStart && selectionEnd) {
      const minRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
      const maxRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
      const minCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
      const maxCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);
      setCellFormats((prev) => {
        const next = { ...prev };
        for (let r = minRow; r <= maxRow; r++) {
          for (let c = minCol; c <= maxCol; c++) {
            delete next[`${r}_${c}`];
          }
        }
        return next;
      });
    } else if (selectedCell) {
      setCellFormats((prev) => {
        const next = { ...prev };
        delete next[`${selectedCell.rowIndex}_${selectedCell.colIndex}`];
        return next;
      });
    } else if (selectedRowIndex >= 0) {
      setCellFormats((prev) => {
        const next = { ...prev };
        [-1, 0, 1, 2, 3, 4].forEach((c) => {
          delete next[`${selectedRowIndex}_${c}`];
        });
        return next;
      });
    }
  };

  const renderSignatureBlock = (isPrintFixed: boolean = false) => {
    const isChallan = docType === "challan";

    return (
      <div className={`sig-block-wrapper w-full ${isPrintFixed ? "print-fixed-sig-inner" : "mt-4 pt-2 print:mt-0 print:pt-0"}`}>
        <div className={`sig-section w-full flex flex-row ${isChallan ? "justify-start" : "justify-between"} items-end gap-10 sm:gap-14`}>
          {/* Receiver's Signature Block */}
          <div className="sig-box w-[240px] max-w-full text-center flex flex-col justify-between h-[80px] print:h-[80px]">
            <div className="h-4 invisible select-none" aria-hidden="true">&nbsp;</div>
            <div className="sig-line border-t-2 border-slate-700 print:border-slate-600 pt-1.5 text-[10pt] print:text-[9.5pt] font-extrabold text-black tracking-wide">
              Receiver's Signature
            </div>
          </div>
          
          {/* Authorized Signature Block (Consistent across quotation & invoice; excluded for Challan format) */}
          {!isChallan && (
            <div className="sig-box w-full sm:w-[240px] print:w-[240px] text-center flex flex-col justify-between h-[80px] print:h-[80px] relative">
              <div className="sig-title text-[9.5pt] print:text-[9pt] font-black text-black uppercase tracking-wider leading-none">
                For {currentCompany.name}
              </div>
              
              {currentCompany.hasStamp && currentCompany.stampUrl && (
                <div className="absolute inset-0 flex items-center justify-center pointer-events-none z-10 select-none pb-2">
                  <img 
                    src={currentCompany.stampUrl}
                    alt={`${currentCompany.name} Stamp`}
                    referrerPolicy="no-referrer"
                    className="w-[96px] h-[96px] sm:w-[100px] sm:h-[100px] print:w-[94px] print:h-[94px] object-contain select-none opacity-90"
                    style={{ printColorAdjust: "exact" }}
                  />
                </div>
              )}

              <div className="sig-line border-t-2 border-slate-700 print:border-slate-600 pt-1.5 text-[10pt] print:text-[9.5pt] font-extrabold relative z-20 text-black tracking-wide leading-none">
                Authorized Signature
              </div>
            </div>
          )}
        </div>

        {/* Non-returnable & non-exchangeable notice */}
        <div className="doc-footer-notice text-center mt-3 pt-1.5 print:mt-1.5 print:pt-1 text-[9px] print:text-[8pt] leading-tight font-bold text-black uppercase tracking-wider">
          ITEMS ONCE SOLD ARE NON-RETURNABLE AND NON-EXCHANGEABLE.
        </div>
      </div>
    );
  };

  return (
    <div className="erp-app-shell flex min-h-screen bg-slate-100 print:bg-white print:p-0 print:m-0 print:block text-slate-900 font-sans antialiased w-full selection:bg-blue-600 selection:text-white">
      {/* Persistent Collapsible Sidebar with built-in responsive mobile drawer */}
      <ErpSidebar
        activeView={activeView}
        onSelectView={(v) => {
          setActiveView(v);
          setIsMobileSidebarOpen(false);
        }}
        onNewDoc={(type) => {
          startNewDoc(type, true);
          setActiveView("editor");
          setIsMobileSidebarOpen(false);
        }}
        isCollapsed={isSidebarCollapsed}
        onToggleCollapse={() => setIsSidebarCollapsed(!isSidebarCollapsed)}
        isMobileOpen={isMobileSidebarOpen}
        onCloseMobile={() => setIsMobileSidebarOpen(false)}
        savedDocsCount={savedDocs.length}
        activeCompany={activeCompany}
        onSelectCompany={handleSwitchCompany}
      />

      {/* Main ERP Content Area */}
      <div className="flex-1 flex flex-col min-w-0 w-full overflow-x-hidden print:overflow-visible print:p-0 print:m-0 print:bg-white print:block">
        {/* Clean Sticky Top Navigation */}
        <div className="no-print print:hidden sticky top-0 z-30 w-full">
          <ErpTopNav
            activeView={activeView}
            docType={docType}
            onSelectDocType={(type) => {
              recordChange();
              setDocType(type);
              setMergedRegions([]);
              setRows((prev) =>
                prev.map((r) => {
                  const q = parseNumericInput(stripHtml(String(r.qty || "")));
                  const p = parseNumericInput(stripHtml(String(r.price || "")));
                  return {
                    ...r,
                    amount: type === "challan" ? 0 : q * p,
                  };
                })
              );
            }}
            currentDocName={savedDocs.find((d) => d.id === currentDocId)?.name || (messers ? `${messers} (${dateVal})` : undefined)}
            saveStatus={saveStatus}
            lastSavedTime={lastSavedTime}
            onSaveDoc={() => saveCurrentDocToApp()}
            onPrint={handlePrint}
            onDownloadPDF={handleDownloadPDF}
            isGeneratingPDF={isGeneratingPDF}
            onDownloadExcel={handleDownloadExcel}
            isGeneratingExcel={isGeneratingExcel}
            onOpenExcelModal={() => setIsExcelModalOpen(true)}
            onToggleMobileSidebar={() => setIsMobileSidebarOpen(!isMobileSidebarOpen)}
            onNavigateToArchive={() => setActiveView(activeView === "saved-docs" ? "editor" : "saved-docs")}
            savedDocsCount={savedDocs.length}
            activeCompany={activeCompany}
            onSelectCompany={handleSwitchCompany}
          />
        </div>

        {/* View 1: ERP Dashboard */}
        {activeView === "dashboard" && (
          <div className="p-4 sm:p-6 lg:p-8 max-w-7xl w-full mx-auto animate-in fade-in duration-200">
            <ErpDashboard
              savedDocs={savedDocs}
              activeCompany={activeCompany}
              onSelectCompany={handleSwitchCompany}
              onOpenDoc={(doc) => {
                loadSavedDoc(doc);
                setActiveView("editor");
              }}
              onNewDoc={(type) => {
                startNewDoc(type, true);
                setActiveView("editor");
              }}
              onDeleteDoc={(id) => deleteSavedDoc(id)}
              onDuplicateDoc={(doc) => {
                loadSavedDoc(doc);
                duplicateCurrentDoc();
              }}
              onSwitchToArchive={() => setActiveView("saved-docs")}
            />
          </div>
        )}

        {/* View 2: Document Editor (centered A4 canvas with minimal surrounding UI & compact ribbon toolbar) */}
        <div
          id="document-editor-wrapper"
          className={activeView === "editor" ? "flex flex-col items-center py-4 px-2 sm:px-4 print:p-0 print:py-0 print:px-0 print:m-0 print:bg-white print:block w-full" : "hidden"}
        >
          {/* Compact Ribbon Toolbar */}
          <div className="w-full max-w-[210mm] no-print print:hidden mb-3">
            <ExcelRibbonToolbar
              onUndo={handleUndo}
              onRedo={handleRedo}
              canUndo={undoStackRef.current.length > 0}
              canRedo={redoStackRef.current.length > 0}
              activeFormat={activeCellFormat}
              onApplyFormat={handleApplyFormat}
              onApplyBorderPreset={handleApplyBorderPreset}
              selectionSummary={selectionSummary}
              canMerge={!!(selectionStart && selectionEnd && (selectionStart.rowIndex !== selectionEnd.rowIndex || selectionStart.colIndex !== selectionEnd.colIndex))}
              onToggleMerge={toggleMergeSelectedRange}
              onClearFormatting={handleClearFormatting}
              docType={docType}
              onSelectDocType={(type) => {
                recordChange();
                setDocType(type);
                setMergedRegions([]);
                setRows((prev) =>
                  prev.map((r) => {
                    const q = parseNumericInput(stripHtml(String(r.qty || "")));
                    const p = parseNumericInput(stripHtml(String(r.price || "")));
                    return {
                      ...r,
                      amount: type === "challan" ? 0 : q * p,
                    };
                  })
                );
              }}
              includeVesselName={includeVesselName}
              onToggleVesselName={setIncludeVesselName}
              includePortBerth={includePortBerth}
              onTogglePortBerth={setIncludePortBerth}
              autoSaveEnabled={autoSaveEnabled}
              lastSavedTime={lastSavedTime}
              currentDocId={currentDocId}
              currentDocName={savedDocs.find((d) => d.id === currentDocId)?.name}
              onCloseCurrentDoc={resetSheetFields}
              onNewDoc={() => startNewDoc()}
              onDuplicateDoc={currentDocId ? duplicateCurrentDoc : undefined}
              onDeleteDoc={currentDocId ? () => deleteSavedDoc(currentDocId) : undefined}
              onSaveDoc={() => saveCurrentDocToApp()}
              saveStatus={saveStatus}
              onOpenExcelModal={() => setIsExcelModalOpen(true)}
              onPrint={handlePrint}
              onDownloadPDF={handleDownloadPDF}
              isGeneratingPDF={isGeneratingPDF}
              onDownloadExcel={handleDownloadExcel}
              isGeneratingExcel={isGeneratingExcel}
              includeDiscount={includeDiscount}
              onToggleDiscount={(val) => setIncludeDiscount(val)}
            />
          </div>

          <div className="w-full flex flex-col items-center print:block print:p-0 print:m-0 print:bg-white">

            {/* A4 Standard-compliant visual grid container */}
            <div className="sheet relative w-full max-w-[210mm] print:max-w-full min-h-[297mm] print:min-h-0 bg-white p-2.5 sm:p-[6mm] print:p-0 shadow-xl print:shadow-none border border-slate-300 print:border-none rounded-xs print:rounded-none box-border z-10 mx-auto overflow-hidden print:overflow-visible">
          
          {/* Anti-slip Background Watermark Asset */}
          <div className="watermark-container absolute inset-0 pointer-events-none flex items-center justify-center overflow-hidden z-0 select-none">
            <img 
              src="https://i.ibb.co.com/3mNycQXx/1.png" 
              alt="Watermark background" 
              referrerPolicy="no-referrer"
              className="object-contain select-none max-w-[500px] w-[70%] opacity-[0.045]"
              style={{ printColorAdjust: "exact" }}
            />
          </div>

          {/* Outer Layout Table ensuring thead repeats company details on multi-page browser printing */}
          <table className="print-outer-layout-table w-full max-w-full table-fixed border-none p-0 m-0 relative z-10 box-border">
            <thead className="print:table-header-group">
              <tr>
                <td className="border-none p-0 m-0">
                  {/* Format of the Page: Written ONLY on the Top of the Page (No Underline) */}
                  <div className="page-format-header text-center mb-2 pb-0 border-none">
                    <span className="doc-title inline-block text-[12pt] sm:text-[14pt] print:text-[13pt] font-black uppercase tracking-[7px] text-black">
                      {docType === "challan" ? "Challan" : docType === "invoice" ? "Invoice" : "Quotation"}
                    </span>
                  </div>

                  {/* Enlarged Business Header without bottom border and with borderless logo */}
                  <div className="business-header border-b-0 print:border-b-0 pb-3.5 sm:pb-4 mb-4 sm:mb-5 print:pb-2.5 print:mb-3.5 flex flex-row items-center justify-between gap-3 sm:gap-4 text-black text-left w-full max-w-full overflow-hidden box-border">
                    <div className="flex items-center gap-3 sm:gap-4 min-w-0 flex-1">
                      <div className="logo-container h-20 w-20 sm:h-24 sm:w-24 print:h-[82px] print:w-[82px] shrink-0 border-none rounded-none shadow-none bg-transparent flex items-center justify-center">
                        <img
                          src={currentCompany.logoUrl}
                          alt={`${currentCompany.name} Logo`}
                          referrerPolicy="no-referrer"
                          className="w-full h-full object-contain"
                        />
                      </div>
                      <div className="min-w-0 flex-1">
                        <h1 className="text-[19pt] sm:text-[23pt] print:text-[21pt] font-black tracking-tight leading-none text-black uppercase truncate">
                          {currentCompany.name}
                        </h1>
                        <p className="company-tagline-1 text-[8pt] sm:text-[9pt] print:text-[8.2pt] font-extrabold text-[#1e40af] tracking-wide uppercase leading-tight mt-1 sm:mt-1.5">
                          {currentCompany.tagline1}
                        </p>
                        {currentCompany.tagline2 && (
                          <p className="company-tagline-2 text-[7.2pt] sm:text-[8pt] print:text-[7.2pt] font-bold text-slate-500 uppercase tracking-wide mt-0.5 leading-tight">
                            {currentCompany.tagline2}
                          </p>
                        )}
                      </div>
                    </div>

                    <div className="contact-details text-right text-[7.2pt] sm:text-[7.8pt] print:text-[7.2pt] text-slate-800 space-y-0.5 leading-tight sm:block hidden print:block shrink-0 max-w-[45%]">
                      <p className="font-bold">
                        Office: <span className="font-medium text-slate-900">{currentCompany.officeAddress}</span>
                      </p>
                      <p className="font-bold">
                        Helplines: <span className="font-medium font-mono text-slate-900">{currentCompany.helplines}</span>
                      </p>
                      <p className="font-bold">
                        Official Email: <span className="font-medium text-slate-900">{currentCompany.email}</span>
                      </p>
                      <p className="font-black text-[7.8pt] tracking-widest text-[#4f46e5] uppercase pt-0.5">
                        {currentCompany.locationCity}
                      </p>
                    </div>
                    
                    {/* Mobile contact information fallback */}
                    <div className="text-center text-[7.5pt] text-slate-800 space-y-0.5 leading-tight sm:hidden print:hidden shrink-0">
                      <p className="font-medium">{currentCompany.officeAddress} &bull; Helplines: {currentCompany.helplines}</p>
                      <p className="font-medium">{currentCompany.email}</p>
                      <p className="font-bold text-[#4f46e5]">{currentCompany.locationCity}</p>
                    </div>
                  </div>

                {/* ==================================================================== */}
                {/* INFORMATION TABLE (Unified Bordered View with Compact Font Size) */}
                {/* ==================================================================== */}
                {(() => {
                  const hasMessers = !!cleanHtmlText(messers);
                  const hasVessel = !!vesselName.trim();
                  const hasPort = !!portBerth.trim();
                  const hasAddress = !!cleanHtmlText(address);
                  const hasLeftMeta = hasMessers || hasVessel || hasPort || hasAddress;

                  const hasRequisition = !!requisitionNo.trim();
                  const hasDate = !!dateVal.trim();
                  const hasInvoiceNo = !!invoiceNo.trim();
                  const hasChallanNo = !!challanNo.trim();
                  const hasPoNumber = !!poNumber.trim();

                  const hasRightMeta =
                    docType === "quotation"
                      ? (hasRequisition || hasDate)
                      : docType === "challan"
                      ? (hasChallanNo || hasRequisition || hasDate)
                      : (hasInvoiceNo || hasChallanNo || hasDate || hasRequisition || hasPoNumber);

                  const hasAnyMeta = hasLeftMeta || hasRightMeta;

                  return (
                    <>
                      {/* Physical spacer between Business Header and Information Table to ensure clear breathing room */}
                      <div
                        className={`meta-top-spacer w-full select-none ${!hasAnyMeta ? 'print:hidden meta-box-empty' : ''}`}
                        aria-hidden="true"
                        style={{ height: '12px', minHeight: '12px', display: 'block', clear: 'both' }}
                      />

                      <div className={`meta-info-container w-full border border-slate-600 print:border-slate-600 rounded-lg print:rounded-none bg-slate-50/60 print:bg-white p-4 sm:p-5 print:p-3 shadow-2xs print:shadow-none box-border ${!hasAnyMeta ? 'print:hidden meta-box-empty' : ''}`}>
                        <div className="meta-grid-inner grid grid-cols-12 gap-0 text-left w-full">
                        {/* Left Column: Client & Vessel Information */}
                        <div className={`meta-left-col col-span-7 space-y-2 sm:space-y-2.5 print:space-y-1.5 pr-4 sm:pr-6 print:pr-4 ${!hasLeftMeta ? 'print:hidden meta-col-empty' : ''} ${!hasRightMeta ? 'print:col-span-12 print:pr-0' : ''}`}>
                          {/* Messers */}
                          <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 w-full ${!hasMessers ? 'print:hidden meta-row-empty' : ''}`}>
                            <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                              <User className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                              <span>Messers:</span>
                            </label>
                            <RichTextCell
                              value={messers}
                              syncId="messers"
                              onChange={(val) => setMessers(val)}
                              placeholder=""
                              className="print:hidden flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus:border-indigo-600 focus:border-solid font-bold text-[6.8pt] sm:text-[7.2pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal transition-colors break-words flex items-center"
                            />
                            <div
                              id="print-messers"
                              className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-bold text-[6.8pt] text-black leading-tight break-words px-1.5 min-h-[20px] print:min-h-[20px]"
                              dangerouslySetInnerHTML={{ __html: messers.trim() || '&nbsp;' }}
                            />
                          </div>

                          {/* Vessel Name: Single row, text directly after colon */}
                          <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 w-full ${!hasVessel ? 'print:hidden meta-row-empty' : ''}`}>
                            <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                              <Ship className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                              <span>Vessel Name:</span>
                            </label>
                            <input
                              type="text"
                              value={vesselName}
                              onChange={(e) => setVesselName(e.target.value)}
                              placeholder=""
                              className="print:hidden flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus:border-indigo-600 focus:border-solid font-semibold text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal transition-colors"
                            />
                            <div
                              id="print-vessel"
                              className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-semibold text-[6.8pt] text-black leading-tight break-words px-1.5 min-h-[20px] print:min-h-[20px]"
                            >
                              {vesselName.trim() || '\u00A0'}
                            </div>
                          </div>

                          {/* Port / Berth: Single row, text directly after colon */}
                          <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 w-full ${!hasPort ? 'print:hidden meta-row-empty' : ''}`}>
                            <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                              <Anchor className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                              <span>Port / Berth:</span>
                            </label>
                            <input
                              type="text"
                              value={portBerth}
                              onChange={(e) => setPortBerth(e.target.value)}
                              placeholder=""
                              className="print:hidden flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus:border-indigo-600 focus:border-solid text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal transition-colors"
                            />
                            <div
                              id="print-port"
                              className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-semibold text-[6.8pt] text-black leading-tight break-words px-1.5 min-h-[20px] print:min-h-[20px]"
                            >
                              {portBerth.trim() || '\u00A0'}
                            </div>
                          </div>

                          {/* Address: Single row, text directly after colon */}
                          <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 w-full ${!hasAddress ? 'print:hidden meta-row-empty' : ''}`}>
                            <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                              <MapPin className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                              <span>Address:</span>
                            </label>
                            <RichTextCell
                              value={address}
                              syncId="address"
                              onChange={(val) => setAddress(val)}
                              placeholder=""
                              className="print:hidden flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus:border-indigo-600 focus:border-solid text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal transition-colors break-words flex items-center"
                            />
                            <div
                              id="print-address"
                              className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none text-[6.8pt] text-black leading-tight break-words px-1.5 min-h-[20px] print:min-h-[20px]"
                              dangerouslySetInnerHTML={{ __html: address.trim() || '&nbsp;' }}
                            />
                          </div>
                        </div>

                        {/* Right Column: References & Date */}
                        <div className={`meta-right-col col-span-5 space-y-2 sm:space-y-2.5 print:space-y-1.5 pl-4 sm:pl-6 print:pl-4 ${!hasRightMeta ? 'print:hidden meta-col-empty' : ''} ${!hasLeftMeta ? 'print:col-span-12 print:pl-0' : ''}`}>
                          {/* Quotation format: ONLY Requisition No. with Date */}
                          {docType === "quotation" && (
                            <div className="space-y-2 sm:space-y-2.5 print:space-y-1.5">
                              <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 w-full ${!hasRequisition ? 'print:hidden meta-row-empty' : ''}`}>
                                <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                                  <FileText className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                                  <span>Requisition No.:</span>
                                </label>
                                <input
                                  type="text"
                                  value={requisitionNo}
                                  onChange={(e) => setRequisitionNo(e.target.value)}
                                  placeholder=""
                                  className="print:hidden flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus:border-indigo-600 focus:border-solid font-mono font-bold text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal transition-colors"
                                />
                                <div
                                  id="print-requisition-no"
                                  className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-mono font-bold text-[6.8pt] text-black leading-tight break-words px-1.5 min-h-[20px] print:min-h-[20px]"
                                >
                                  {requisitionNo.trim() || '\u00A0'}
                                </div>
                              </div>

                              {/* Date field */}
                              <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 relative w-full min-w-0 ${!hasDate ? 'print:hidden meta-row-empty' : ''}`}>
                                <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                                  <Calendar className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                                  <span>Date:</span>
                                </label>
                                <div className="flex items-center flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus-within:border-indigo-600 focus-within:border-solid transition-colors print:hidden">
                                  <input
                                    type="text"
                                    value={dateVal}
                                    onChange={(e) => setDateVal(e.target.value)}
                                    className="flex-1 min-w-0 h-full font-mono font-bold text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal"
                                  />
                                  <button
                                    type="button"
                                    onClick={triggerDatePicker}
                                    className="h-[20px] w-[20px] p-0 shrink-0 hover:bg-slate-200/80 rounded text-slate-600 transition-colors cursor-pointer flex items-center justify-center mr-0.5"
                                    title="Open Date Picker"
                                  >
                                    <Calendar className="h-2.5 w-2.5 text-indigo-600" />
                                  </button>
                                </div>
                                <div
                                  id="print-date"
                                  className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-mono font-bold text-[6.8pt] text-black leading-tight px-1.5 min-h-[20px] print:min-h-[20px]"
                                >
                                  {dateVal.trim() || '\u00A0'}
                                </div>
                                <input
                                  ref={dateRef}
                                  type="date"
                                  onChange={handleDatePickerChange}
                                  className="absolute right-0 bottom-0 w-[18px] h-[18px] opacity-0 pointer-events-none"
                                />
                              </div>
                            </div>
                          )}

                          {/* Invoice format: All single lines, no side-by-side fields */}
                          {docType === "invoice" && (
                            <div className="space-y-2 sm:space-y-2.5 print:space-y-1.5">
                              <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 w-full ${!hasInvoiceNo ? 'print:hidden meta-row-empty' : ''}`}>
                                <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                                  <FileText className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                                  <span>Invoice No.:</span>
                                </label>
                                <input
                                  type="text"
                                  value={invoiceNo}
                                  onChange={(e) => setInvoiceNo(e.target.value)}
                                  placeholder=""
                                  className="print:hidden flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus:border-indigo-600 focus:border-solid font-mono font-bold text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal transition-colors"
                                />
                                <div
                                  id="print-invoice-no"
                                  className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-mono font-bold text-[6.8pt] text-black leading-tight break-words px-1.5 min-h-[20px] print:min-h-[20px]"
                                >
                                  {invoiceNo.trim() || '\u00A0'}
                                </div>
                              </div>

                              <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 w-full ${!hasChallanNo ? 'print:hidden meta-row-empty' : ''}`}>
                                <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                                  <Hash className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                                  <span>Challan No.:</span>
                                </label>
                                <input
                                  type="text"
                                  value={challanNo}
                                  onChange={(e) => setChallanNo(e.target.value)}
                                  placeholder=""
                                  className="print:hidden flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus:border-indigo-600 focus:border-solid font-mono font-bold text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal transition-colors"
                                />
                                <div
                                  id="print-challan-no"
                                  className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-mono font-bold text-[6.8pt] text-black leading-tight break-words px-1.5 min-h-[20px] print:min-h-[20px]"
                                >
                                  {challanNo.trim() || '\u00A0'}
                                </div>
                              </div>

                              {/* Date field */}
                              <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 relative w-full min-w-0 ${!hasDate ? 'print:hidden meta-row-empty' : ''}`}>
                                <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                                  <Calendar className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                                  <span>Date:</span>
                                </label>
                                <div className="flex items-center flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus-within:border-indigo-600 focus-within:border-solid transition-colors print:hidden">
                                  <input
                                    type="text"
                                    value={dateVal}
                                    onChange={(e) => setDateVal(e.target.value)}
                                    className="flex-1 min-w-0 h-full font-mono font-bold text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal"
                                  />
                                  <button
                                    type="button"
                                    onClick={triggerDatePicker}
                                    className="h-[20px] w-[20px] p-0 shrink-0 hover:bg-slate-200/80 rounded text-slate-600 transition-colors cursor-pointer flex items-center justify-center mr-0.5"
                                    title="Open Date Picker"
                                  >
                                    <Calendar className="h-2.5 w-2.5 text-indigo-600" />
                                  </button>
                                </div>
                                <div
                                  id="print-date"
                                  className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-mono font-bold text-[6.8pt] text-black leading-tight px-1.5 min-h-[20px] print:min-h-[20px]"
                                >
                                  {dateVal.trim() || '\u00A0'}
                                </div>
                                <input
                                  ref={dateRef}
                                  type="date"
                                  onChange={handleDatePickerChange}
                                  className="absolute right-0 bottom-0 w-[18px] h-[18px] opacity-0 pointer-events-none"
                                />
                              </div>

                              <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 w-full ${!hasRequisition ? 'print:hidden meta-row-empty' : ''}`}>
                                <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                                  <FileText className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                                  <span>Requisition No.:</span>
                                </label>
                                <input
                                  type="text"
                                  value={requisitionNo}
                                  onChange={(e) => setRequisitionNo(e.target.value)}
                                  placeholder=""
                                  className="print:hidden flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus:border-indigo-600 focus:border-solid font-mono font-medium text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal transition-colors"
                                />
                                <div
                                  id="print-requisition-no"
                                  className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-mono font-medium text-[6.8pt] text-black leading-tight break-words px-1.5 min-h-[20px] print:min-h-[20px]"
                                >
                                  {requisitionNo.trim() || '\u00A0'}
                                </div>
                              </div>

                              <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 w-full ${!hasPoNumber ? 'print:hidden meta-row-empty' : ''}`}>
                                <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                                  <Hash className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                                  <span>PO Number:</span>
                                </label>
                                <input
                                  type="text"
                                  value={poNumber}
                                  onChange={(e) => setPoNumber(e.target.value)}
                                  placeholder=""
                                  className="print:hidden flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus:border-indigo-600 focus:border-solid font-mono font-medium text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal transition-colors"
                                />
                                <div
                                  id="print-po-number"
                                  className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-mono font-medium text-[6.8pt] text-black leading-tight break-words px-1.5 min-h-[20px] print:min-h-[20px]"
                                >
                                  {poNumber.trim() || '\u00A0'}
                                </div>
                              </div>
                            </div>
                          )}

                          {/* Challan format: All single lines, no side-by-side fields */}
                          {docType === "challan" && (
                            <div className="space-y-2 sm:space-y-2.5 print:space-y-1.5">
                              <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 w-full ${!hasChallanNo ? 'print:hidden meta-row-empty' : ''}`}>
                                <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                                  <FileText className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                                  <span>Challan No.:</span>
                                </label>
                                <input
                                  type="text"
                                  value={challanNo}
                                  onChange={(e) => setChallanNo(e.target.value)}
                                  placeholder=""
                                  className="print:hidden flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus:border-indigo-600 focus:border-solid font-mono font-bold text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal transition-colors"
                                />
                                <div
                                  id="print-challan-no"
                                  className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-mono font-bold text-[6.8pt] text-black leading-tight break-words px-1.5 min-h-[20px] print:min-h-[20px]"
                                >
                                  {challanNo.trim() || '\u00A0'}
                                </div>
                              </div>

                              <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 w-full ${!hasRequisition ? 'print:hidden meta-row-empty' : ''}`}>
                                <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                                  <Hash className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                                  <span>Requisition No.:</span>
                                </label>
                                <input
                                  type="text"
                                  value={requisitionNo}
                                  onChange={(e) => setRequisitionNo(e.target.value)}
                                  placeholder=""
                                  className="print:hidden flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus:border-indigo-600 focus:border-solid font-mono font-bold text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal transition-colors"
                                />
                                <div
                                  id="print-requisition-no"
                                  className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-mono font-bold text-[6.8pt] text-black leading-tight break-words px-1.5 min-h-[20px] print:min-h-[20px]"
                                >
                                  {requisitionNo.trim() || '\u00A0'}
                                </div>
                              </div>

                              {/* Date field */}
                              <div className={`meta-info-row min-h-[34px] sm:min-h-[38px] print:min-h-[22px] flex items-center gap-1.5 py-0.5 print:py-0 relative w-full min-w-0 ${!hasDate ? 'print:hidden meta-row-empty' : ''}`}>
                                <label className="text-[6.2pt] sm:text-[6.8pt] print:text-[6.2pt] font-extrabold text-slate-800 print:text-black uppercase tracking-wider flex items-center gap-1 shrink-0 whitespace-nowrap">
                                  <Calendar className="h-2.5 w-2.5 sm:h-3 sm:w-3 print:h-2.5 print:w-2.5 text-slate-600 print:text-black shrink-0" />
                                  <span>Date:</span>
                                </label>
                                <div className="flex items-center flex-1 min-w-0 h-[30px] min-h-[30px] sm:h-[34px] sm:min-h-[34px] box-border border-b border-dotted border-slate-400 focus-within:border-indigo-600 focus-within:border-solid transition-colors print:hidden">
                                  <input
                                    type="text"
                                    value={dateVal}
                                    onChange={(e) => setDateVal(e.target.value)}
                                    className="flex-1 min-w-0 h-full font-mono font-bold text-[6.5pt] sm:text-[7pt] text-slate-900 outline-none bg-transparent px-1.5 py-1 leading-normal"
                                  />
                                  <button
                                    type="button"
                                    onClick={triggerDatePicker}
                                    className="h-[20px] w-[20px] p-0 shrink-0 hover:bg-slate-200/80 rounded text-slate-600 transition-colors cursor-pointer flex items-center justify-center mr-0.5"
                                    title="Open Date Picker"
                                  >
                                    <Calendar className="h-2.5 w-2.5 text-indigo-600" />
                                  </button>
                                </div>
                                <div
                                  id="print-date"
                                  className="hidden print:flex flex-1 min-w-0 items-center border-b border-dotted border-slate-400 print:border-none font-mono font-bold text-[6.8pt] text-black leading-tight px-1.5 min-h-[20px] print:min-h-[20px]"
                                >
                                  {dateVal.trim() || '\u00A0'}
                                </div>
                                <input
                                  ref={dateRef}
                                  type="date"
                                  onChange={handleDatePickerChange}
                                  className="absolute right-0 bottom-0 w-[18px] h-[18px] opacity-0 pointer-events-none"
                                />
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                      </div>

                      {/* Physical spacer between Information Table and Items Table ensuring they are NEVER connected */}
                      <div
                        className="meta-bottom-spacer w-full select-none"
                        aria-hidden="true"
                        style={{ height: '12px', minHeight: '12px', display: 'block', clear: 'both' }}
                      />
                    </>
                  );
                })()}


              </td>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td className="border-none p-0 m-0">

                {/* Main Data Sheet Table */}
                <div className="items-table-wrapper w-full overflow-x-auto no-scrollbar">
                  <table className="main-table w-full min-w-full border-collapse border border-slate-600 print:border-slate-600 table-fixed text-[8pt]">
                    <thead>
                      <tr className="bg-slate-100/90 print:bg-transparent text-black text-[8.5pt] sm:text-[9pt]">
                        <th className={`${docType === 'challan' ? 'w-[7%]' : 'w-[6%]'} border border-slate-600 print:border-slate-600 py-1.5 px-1 text-center font-bold bg-slate-100/90 print:bg-transparent text-black text-[8.5pt] sm:text-[9pt] uppercase tracking-wider`}>SL</th>
                        <th className={`${docType === 'challan' ? 'w-[75%]' : 'w-[56%]'} border border-slate-600 print:border-slate-600 py-1.5 px-2 text-left font-bold bg-slate-100/90 print:bg-transparent text-black text-[8.5pt] sm:text-[9pt] uppercase tracking-wider`}>Description of Marine Items / Spare Parts</th>
                        <th className={`${docType === 'challan' ? 'w-[9%]' : 'w-[7%]'} border border-slate-600 print:border-slate-600 py-1.5 px-1 text-center font-bold bg-slate-100/90 print:bg-transparent text-black text-[8.5pt] sm:text-[9pt] uppercase tracking-wider`}>Qty</th>
                        <th className={`${docType === 'challan' ? 'w-[9%]' : 'w-[8%]'} border border-slate-600 print:border-slate-600 py-1.5 px-1 text-center font-bold bg-slate-100/90 print:bg-transparent text-black text-[8.5pt] sm:text-[9pt] uppercase tracking-wider`}>Unit</th>
                        {docType !== "challan" && (
                          <>
                            <th className="w-[11%] border border-slate-600 print:border-slate-600 py-1.5 px-1 text-center font-bold bg-slate-100/90 print:bg-transparent text-black text-[8.5pt] sm:text-[9pt] uppercase tracking-wider">Price</th>
                            <th className="w-[12%] border border-slate-600 print:border-slate-600 py-1.5 px-1 text-center font-bold bg-slate-100/90 print:bg-transparent text-black text-[8.5pt] sm:text-[9pt] uppercase tracking-wider">Amount</th>
                          </>
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((row, idx) => (
                        <tr 
                          key={idx} 
                          className={`group hover:bg-slate-50/50 transition-colors ${
                            idx === safeSelectedRowIndex ? "bg-indigo-50/10" : ""
                          }`}
                        >
                          {GRID_COLUMNS.map((colIndex) => {
                            const { region, isAnchor } = getMergeInfo(idx, colIndex);
                            if (region && !isAnchor) return null;

                            const colSpan = region ? region.endCol - region.startCol + 1 : 1;
                            const rowSpan = region ? region.endRow - region.startRow + 1 : 1;

                            if (colIndex === -1) {
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  style={getCellStyle(idx, -1)}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, -1)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, -1)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, -1)}
                                  onClick={() => handleCellClick(idx, -1)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, -1)}
                                  className={getCellClassName(idx, -1, `border border-slate-400 print:border-slate-600 text-center font-mono text-[7.5pt] sm:text-[8pt] align-middle py-0.5 px-0.5 whitespace-nowrap leading-tight transition-all cursor-pointer select-none bg-slate-50/30 text-slate-800`)}
                                >
                                  {idx + 1}
                                </td>
                              );
                            }

                            if (colIndex === 0) {
                              const cellStyle = getCellStyle(idx, 0);
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  style={cellStyle}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 0)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 0)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 0)}
                                  onClick={() => handleCellClick(idx, 0)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 0)}
                                  className={getCellClassName(idx, 0, `border border-slate-400 print:border-slate-600 text-left px-1.5 py-0.5 text-[8pt] sm:text-[8.5pt] align-middle whitespace-normal transition-all cursor-text ${region ? "bg-slate-50/40" : ""}`)}
                                >
                                  <RichTextCell
                                    value={row.desc}
                                    syncId={`desc-${idx}`}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 0 });
                                    }}
                                    onChange={(val) => handleRowChange(idx, "desc", val)}
                                    onKeyDown={(e) => {
                                      if (e.key === "Enter" && !e.shiftKey) {
                                        e.preventDefault();
                                        const nextArea = document.querySelector(`[data-row="${idx + 1}"][data-col="0"]`) as HTMLElement | null;
                                        if (nextArea) nextArea.focus();
                                      } else {
                                        handleKeyDown(e, idx, 0);
                                      }
                                    }}
                                    onPaste={(e) => handlePaste(e, idx, 0)}
                                    dataRow={idx}
                                    dataCol={0}
                                    style={{ ...cellStyle, wordBreak: "break-word", overflowWrap: "anywhere", whiteSpace: "pre-wrap" }}
                                    className="w-full min-w-full text-left border-none outline-none bg-transparent p-0 text-slate-800 text-[8pt] sm:text-[8.5pt] leading-[1.25] block overflow-visible py-[1px] whitespace-pre-wrap break-words no-print print:hidden font-normal"
                                  />
                                  <div 
                                    id={`print-desc-${idx}`}
                                    style={{ ...cellStyle, wordBreak: "break-word", overflowWrap: "anywhere", whiteSpace: "pre-wrap" }} 
                                    className="hidden print:block whitespace-pre-wrap break-words text-left text-slate-900 leading-[1.25] py-[1px] text-[8pt]"
                                    dangerouslySetInnerHTML={{ __html: row.desc || "&nbsp;" }}
                                  />
                                </td>
                              );
                            }

                            if (colIndex === 1) {
                              const cellStyle = getCellStyle(idx, 1);
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  style={cellStyle}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 1)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 1)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 1)}
                                  onClick={() => handleCellClick(idx, 1)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 1)}
                                  className={getCellClassName(idx, 1, "border border-slate-400 print:border-slate-600 text-center font-mono text-[8pt] sm:text-[8.5pt] align-middle py-0.5 px-1 transition-all cursor-text")}
                                >
                                  <RichTextCell
                                    value={row.qty}
                                    syncId={`qty-${idx}`}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 1 });
                                    }}
                                    onChange={(val) => handleRowChange(idx, "qty", val)}
                                    onKeyDown={(e) => handleKeyDown(e, idx, 1)}
                                    onPaste={(e) => handlePaste(e, idx, 1)}
                                    dataRow={idx}
                                    dataCol={1}
                                    style={cellStyle}
                                    className="w-full text-center border-none outline-none bg-transparent px-0 font-mono text-slate-800 align-middle overflow-visible py-[1px] leading-[1.25] whitespace-nowrap break-normal no-print print:hidden text-[8pt] sm:text-[8.5pt]"
                                  />
                                  <div 
                                    id={`print-qty-${idx}`}
                                    style={cellStyle} 
                                    className="hidden print:block whitespace-nowrap break-normal text-center font-mono text-slate-900 py-[1px] leading-[1.25] text-[8pt]"
                                    dangerouslySetInnerHTML={{ __html: row.qty || "&nbsp;" }}
                                  />
                                </td>
                              );
                            }

                            if (colIndex === 2) {
                              const cellStyle = getCellStyle(idx, 2);
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  style={cellStyle}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 2)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 2)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 2)}
                                  onClick={() => handleCellClick(idx, 2)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 2)}
                                  className={getCellClassName(idx, 2, "border border-slate-400 print:border-slate-600 text-center text-[8pt] sm:text-[8.5pt] align-middle py-0.5 px-1 transition-all cursor-text")}
                                >
                                  <RichTextCell
                                    value={row.unit}
                                    syncId={`unit-${idx}`}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 2 });
                                    }}
                                    onChange={(val) => handleRowChange(idx, "unit", val)}
                                    onKeyDown={(e) => handleKeyDown(e, idx, 2)}
                                    onPaste={(e) => handlePaste(e, idx, 2)}
                                    dataRow={idx}
                                    dataCol={2}
                                    style={cellStyle}
                                    className="w-full text-center border-none outline-none bg-transparent px-0 text-slate-800 align-middle overflow-visible py-[1px] leading-[1.25] whitespace-nowrap break-normal no-print print:hidden text-[8pt] sm:text-[8.5pt]"
                                  />
                                  <div 
                                    id={`print-unit-${idx}`}
                                    style={cellStyle} 
                                    className="hidden print:block whitespace-nowrap break-normal text-center text-slate-900 py-[1px] leading-[1.25] text-[8pt]"
                                    dangerouslySetInnerHTML={{ __html: row.unit || "&nbsp;" }}
                                  />
                                </td>
                              );
                            }

                            if (colIndex === 3) {
                              const cellStyle = getCellStyle(idx, 3);
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  style={cellStyle}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 3)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 3)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 3)}
                                  onClick={() => handleCellClick(idx, 3)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 3)}
                                  className={getCellClassName(idx, 3, "border border-slate-400 print:border-slate-600 text-center font-mono text-[8pt] sm:text-[8.5pt] align-middle py-0.5 px-1 transition-all cursor-text")}
                                >
                                  <RichTextCell
                                    value={row.price}
                                    syncId={`price-${idx}`}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 3 });
                                    }}
                                    onChange={(val) => handleRowChange(idx, "price", val)}
                                    onKeyDown={(e) => handleKeyDown(e, idx, 3)}
                                    onPaste={(e) => handlePaste(e, idx, 3)}
                                    dataRow={idx}
                                    dataCol={3}
                                    style={cellStyle}
                                    className="w-full text-center border-none outline-none bg-transparent px-0 font-mono text-slate-800 align-middle overflow-visible py-[1px] leading-[1.25] whitespace-nowrap break-normal no-print print:hidden text-[8pt] sm:text-[8.5pt]"
                                  />
                                  <div 
                                    id={`print-price-${idx}`}
                                    style={cellStyle} 
                                    className="hidden print:block whitespace-nowrap break-normal text-center font-mono text-slate-900 py-[1px] leading-[1.25] text-[8pt]"
                                    dangerouslySetInnerHTML={{ __html: row.price || "&nbsp;" }}
                                  />
                                </td>
                              );
                            }

                            const cellStyle = getCellStyle(idx, 4);
                            return (
                              <td
                                key={colIndex}
                                colSpan={colSpan}
                                rowSpan={rowSpan}
                                style={cellStyle}
                                onMouseDown={(e) => handleCellMouseDown(e, idx, 4)}
                                onMouseEnter={() => handleCellMouseEnter(idx, 4)}
                                onMouseUp={(e) => handleCellMouseUp(e, idx, 4)}
                                onClick={() => handleCellClick(idx, 4)}
                                onContextMenu={(e) => handleCellContextMenu(e, idx, 4)}
                                className={getCellClassName(idx, 4, "border border-slate-400 print:border-slate-600 text-right pr-1.5 pl-1 font-mono text-[8pt] sm:text-[8.5pt] font-semibold text-slate-800 align-middle py-0.5 transition-all cursor-pointer")}
                              >
                                <div style={cellStyle} className="whitespace-nowrap overflow-visible leading-[1.25] text-[8pt] sm:text-[8.5pt]">
                                  {row.amount !== 0 ? row.amount.toLocaleString("en-US", { minimumFractionDigits: 2 }) : "0.00"}
                                </div>
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Grid Line Actions & Dynamic Row Controller */}
                <div className="no-print print:hidden my-2.5 px-3 py-2 bg-gradient-to-r from-slate-50 via-indigo-50/30 to-slate-50 rounded-xl border border-slate-200/80 shadow-xs flex flex-wrap items-center justify-between gap-3">
                  {/* Plus/Minus & Batch Line Adders */}
                  <div className="flex flex-wrap items-center gap-2">
                    {/* Add 1 Line */}
                    <button
                      type="button"
                      onClick={() => addRows(1)}
                      className="bg-indigo-600 hover:bg-indigo-700 active:scale-95 text-white font-bold text-xs h-8 w-8 rounded-lg shadow-xs transition-all cursor-pointer flex items-center justify-center"
                      title="Add line (+)"
                    >
                      <Plus className="h-4 w-4 stroke-[2.5]" />
                    </button>

                    {/* Subtract 1 Line */}
                    <button
                      type="button"
                      onClick={removeRow}
                      disabled={rows.length <= 1}
                      className="bg-white hover:bg-rose-50 active:scale-95 border border-rose-200 text-rose-600 font-bold text-xs h-8 w-8 rounded-lg shadow-xs transition-all cursor-pointer flex items-center justify-center disabled:opacity-40 disabled:cursor-not-allowed"
                      title="Subtract line (-)"
                    >
                      <Minus className="h-4 w-4 stroke-[2.5]" />
                    </button>

                    <div className="h-5 w-[1px] bg-slate-200 mx-0.5 hidden sm:block" />

                    {/* Minimal quick add multiple lines */}
                    <form
                      onSubmit={(e) => {
                        e.preventDefault();
                        const num = parseInt(customRowCountInput, 10);
                        if (!isNaN(num) && num > 0) {
                          addRows(num);
                        }
                      }}
                      className="flex items-center gap-1 bg-white border border-slate-300 hover:border-indigo-400 focus-within:border-indigo-500 focus-within:ring-2 focus-within:ring-indigo-100 rounded-lg p-0.5 transition-all shadow-2xs"
                    >
                      <span className="text-xs font-black text-indigo-600 pl-2 pr-0.5">+</span>
                      <input
                        type="number"
                        min="1"
                        max="1000"
                        value={customRowCountInput}
                        onChange={(e) => setCustomRowCountInput(e.target.value)}
                        placeholder="10"
                        className="w-12 h-6.5 text-center text-xs font-mono font-bold text-slate-800 border-none outline-none focus:outline-none bg-slate-50 rounded px-1"
                        title="Enter number of lines to add (max 1,000 total)"
                      />
                      <button
                        type="submit"
                        disabled={rows.length >= 1000}
                        className="bg-indigo-50 hover:bg-indigo-600 text-indigo-700 hover:text-white disabled:opacity-40 disabled:hover:bg-indigo-50 disabled:hover:text-indigo-700 font-bold text-xs h-6.5 w-7 rounded transition-all cursor-pointer flex items-center justify-center"
                        title="Add lines (+)"
                      >
                        <Plus className="h-3.5 w-3.5 stroke-[2.5]" />
                      </button>
                    </form>

                    {/* Minimal quick subtract multiple lines */}
                    <form
                      onSubmit={(e) => {
                        e.preventDefault();
                        const num = parseInt(customSubtractCountInput, 10);
                        if (!isNaN(num) && num > 0) {
                          removeRows(num);
                        }
                      }}
                      className="flex items-center gap-1 bg-white border border-slate-300 hover:border-rose-400 focus-within:border-rose-500 focus-within:ring-2 focus-within:ring-rose-100 rounded-lg p-0.5 transition-all shadow-2xs"
                    >
                      <span className="text-xs font-black text-rose-600 pl-2 pr-0.5">-</span>
                      <input
                        type="number"
                        min="1"
                        max="1000"
                        value={customSubtractCountInput}
                        onChange={(e) => setCustomSubtractCountInput(e.target.value)}
                        placeholder="10"
                        className="w-12 h-6.5 text-center text-xs font-mono font-bold text-slate-800 border-none outline-none focus:outline-none bg-slate-50 rounded px-1"
                        title="Enter number of lines to subtract"
                      />
                      <button
                        type="submit"
                        disabled={rows.length <= 1}
                        className="bg-rose-50 hover:bg-rose-600 text-rose-700 hover:text-white disabled:opacity-40 disabled:hover:bg-rose-50 disabled:hover:text-rose-700 font-bold text-xs h-6.5 w-7 rounded transition-all cursor-pointer flex items-center justify-center"
                        title="Subtract lines (-)"
                      >
                        <Minus className="h-3.5 w-3.5 stroke-[2.5]" />
                      </button>
                    </form>
                  </div>

                  {/* Status counter */}
                  <div className="flex items-center gap-1.5 text-[10px] text-slate-500 font-semibold shrink-0">
                    <span className="bg-slate-200/80 text-slate-700 px-2.5 py-1 rounded-md font-mono text-[11px]" title="Lines / 1,000 Max Limit">
                      <strong className="text-slate-900">{rows.length}</strong> / 1000
                    </span>
                  </div>
                </div>

                {/* Bottom closing wraps, sums, signatures */}
                <div className="closing-wrap mt-1.5">
                  {docType !== "challan" && (
                    <table className="closing-row w-full max-w-full table-fixed border-collapse border border-slate-600 print:border-slate-600 mt-1.5 bg-white text-black z-10 relative">
                      <tbody>
                        {docType === "invoice" ? (
                          <>
                            <tr className="align-stretch">
                              <td rowSpan={includeDiscount ? 5 : 4} className="amount-words-container w-1/2 border-r border-slate-400 print:border-slate-600 p-1.5 bg-slate-50/50 text-left align-middle">
                                <span className="font-extrabold text-[6.5pt] text-slate-700 uppercase tracking-wider block mb-0.5">
                                  Amount in Words:
                                </span>
                                <span className="text-[8pt] font-mono italic text-black font-black uppercase leading-tight">
                                  {numberToWords(calculatedGrandTotal, currency || "Taka")}
                                </span>
                              </td>
                              <td className="w-1/2 p-0 border-b border-slate-400 print:border-slate-600 align-stretch">
                                <div className="flex flex-row items-stretch h-full min-h-[20px] w-full">
                                  <div className="total-lbl bg-slate-50 w-[145px] sm:w-[155px] shrink-0 pr-1.5 text-right text-[7.5pt] font-bold uppercase flex items-center justify-end tracking-wider">
                                    <div className="flex items-center justify-end gap-1 w-full pl-1">
                                      <span>SUBTOTAL</span>
                                      <span className="font-bold text-[8pt] shrink-0">=</span>
                                      {!includeDiscount && (
                                        <button
                                          type="button"
                                          id="btn-add-invoice-discount"
                                          onClick={() => {
                                            setIncludeDiscount(true);
                                            if (!discountValue || discountValue === "0") {
                                              setDiscountValue("");
                                            }
                                          }}
                                          className="no-print print:hidden px-1 py-0.2 text-[6pt] font-bold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded cursor-pointer transition-colors flex items-center gap-0.5 shrink-0"
                                          title="Add discount option to invoice"
                                        >
                                          <Plus className="w-1.5 h-1.5" />
                                          <span>DISC.</span>
                                        </button>
                                      )}
                                    </div>
                                  </div>
                                  <div className="total-val flex-grow text-right pr-2.5 text-[8.5pt] font-mono font-black flex items-center justify-end px-1.5 py-0.5 leading-tight">
                                    {rowsTotal.toLocaleString("en-US", { minimumFractionDigits: 2 })}
                                  </div>
                                </div>
                              </td>
                            </tr>
                            {includeDiscount && (
                              <tr className="align-stretch">
                                <td className="w-1/2 p-0 border-b border-slate-400 print:border-slate-600 align-stretch">
                                  <div className="flex flex-row items-stretch h-full min-h-[20px] w-full">
                                    <div className="total-lbl bg-slate-50 w-[145px] sm:w-[155px] shrink-0 pr-1.5 text-right text-[7.5pt] font-bold uppercase flex items-center justify-end tracking-wider">
                                      <div className="flex items-center justify-end gap-1 w-full pl-0.5 overflow-visible">
                                        <div className="flex items-center gap-1 no-print print:hidden shrink-0">
                                          {/* Fast 1-click toggle between % and Fixed */}
                                          <div className="inline-flex rounded border border-slate-300 bg-slate-100 p-0.5 shrink-0" title="Switch discount type (% or 123)">
                                            <button
                                              type="button"
                                              onClick={() => setDiscountType("percentage")}
                                              className={`px-1 py-0.2 text-[6.5pt] font-bold rounded-xs transition-colors cursor-pointer ${
                                                discountType === "percentage"
                                                  ? "bg-blue-600 text-white shadow-2xs"
                                                  : "text-slate-600 hover:text-slate-900"
                                              }`}
                                              title="Percentage discount (%)"
                                            >
                                              %
                                            </button>
                                            <button
                                              type="button"
                                              onClick={() => setDiscountType("fixed")}
                                              className={`px-1 py-0.2 text-[6.5pt] font-bold rounded-xs transition-colors cursor-pointer ${
                                                discountType === "fixed"
                                                  ? "bg-blue-600 text-white shadow-2xs"
                                                  : "text-slate-600 hover:text-slate-900"
                                              }`}
                                              title="Fixed amount discount (123)"
                                            >
                                              123
                                            </button>
                                          </div>
                                          <select
                                            id="invoice-discount-type-select"
                                            value={discountType}
                                            onChange={(e) => setDiscountType(e.target.value as "percentage" | "fixed")}
                                            className="sr-only"
                                            aria-label="Discount Type"
                                          >
                                            <option value="percentage">%</option>
                                            <option value="fixed">123</option>
                                          </select>
                                          <div className="flex items-center gap-0.5 shrink-0">
                                            <input
                                              type="text"
                                              id="invoice-discount-value-input"
                                              value={discountValue}
                                              onChange={(e) => {
                                                const val = e.target.value;
                                                if (val === "" || /^-?\d*[.,]?\d*$/.test(val)) {
                                                  setDiscountValue(val);
                                                }
                                              }}
                                              placeholder="0"
                                              className={`h-[18px] text-center border border-slate-300 rounded font-mono text-[7.5pt] font-bold bg-white text-slate-900 px-1 focus:outline-none focus:ring-1 focus:ring-blue-500 shrink-0 ${
                                                discountType === "fixed" ? "w-13" : "w-9"
                                              }`}
                                              title={`Enter ${discountType === "percentage" ? "discount percentage" : "fixed discount amount"}`}
                                            />
                                            <span className="text-[7pt] font-black text-slate-700 font-mono shrink-0">
                                              {discountType === "percentage" ? "%" : (currency === "USD" ? "$" : "Tk")}
                                            </span>
                                          </div>
                                          <button
                                            type="button"
                                            id="btn-remove-invoice-discount"
                                            onClick={() => setIncludeDiscount(false)}
                                            title="Remove discount from invoice"
                                            className="text-slate-400 hover:text-rose-600 hover:bg-rose-50 p-0.5 rounded transition-colors cursor-pointer shrink-0"
                                            aria-label="Remove discount"
                                          >
                                            <X className="w-3 h-3 stroke-[2.5]" />
                                          </button>
                                        </div>
                                        {/* Print: Whole word "DISCOUNT" without disclosing percentage */}
                                        <span className="hidden print:inline font-bold uppercase tracking-wider text-[7.5pt] text-black">
                                          DISCOUNT
                                        </span>
                                        <span className="font-bold text-[8pt] text-slate-900 ml-0.5 shrink-0">=</span>
                                      </div>
                                    </div>
                                    <div className="total-val flex-grow text-right pr-2.5 text-[8.5pt] font-mono font-bold flex items-center justify-end px-1.5 py-0.5 leading-tight text-rose-700">
                                      {discountAmount > 0 ? `-${discountAmount.toLocaleString("en-US", { minimumFractionDigits: 2 })}` : "0.00"}
                                    </div>
                                  </div>
                                </td>
                              </tr>
                            )}
                            <tr className="align-stretch">
                              <td className="w-1/2 p-0 border-b border-slate-400 print:border-slate-600 align-stretch">
                                <div className="flex flex-row items-stretch h-full min-h-[20px] w-full">
                                  <div className="total-lbl bg-slate-50 w-[145px] sm:w-[155px] shrink-0 pr-1.5 text-right text-[7.5pt] font-bold uppercase flex items-center justify-end tracking-wider">
                                    <div className="flex items-center justify-end gap-1 w-full pl-1">
                                      <span>VAT</span>
                                      <div className="flex items-center gap-0.5 no-print print:hidden shrink-0">
                                        <input
                                          type="text"
                                          value={vatPercent}
                                          onChange={(e) => {
                                            const val = e.target.value;
                                            if (val === "" || /^-?\d*[.,]?\d*$/.test(val)) {
                                              setVatPercent(val);
                                            }
                                          }}
                                          className="h-[18px] w-8 text-center border border-slate-300 rounded font-mono text-[7pt] bg-white text-slate-800 px-0.5"
                                        />
                                        <span className="text-[7pt] font-bold text-slate-700">%</span>
                                      </div>
                                      <span className="font-bold text-[8pt] shrink-0">=</span>
                                    </div>
                                  </div>
                                  <div className="total-val flex-grow text-right pr-2.5 text-[8.5pt] font-mono font-semibold flex items-center justify-end px-1.5 py-0.5 leading-tight">
                                    {vatAmount.toLocaleString("en-US", { minimumFractionDigits: 2 })}
                                  </div>
                                </div>
                              </td>
                            </tr>
                            <tr className="align-stretch">
                              <td className="w-1/2 p-0 border-b border-slate-400 print:border-slate-600 align-stretch">
                                <div className="flex flex-row items-stretch h-full min-h-[20px] w-full">
                                  <div className="total-lbl bg-slate-50 w-[145px] sm:w-[155px] shrink-0 pr-1.5 text-right text-[7.5pt] font-bold uppercase flex items-center justify-end tracking-wider">
                                    <div className="flex items-center justify-end gap-1 w-full pl-1">
                                      <span className="inline-block w-[90px] text-[9px]">Transportation</span>
                                      <div className="flex items-center no-print print:hidden shrink-0">
                                        <input
                                          type="text"
                                          value={transportationFee}
                                          onChange={(e) => {
                                            const val = e.target.value;
                                            if (val === "" || /^-?\d*[.,]?\d*$/.test(val)) {
                                              setTransportationFee(val);
                                            }
                                          }}
                                          placeholder="0"
                                          className="h-[18px] w-12 text-center border border-slate-300 rounded font-mono text-[7pt] bg-white text-slate-800 px-0.5"
                                        />
                                      </div>
                                      <span className="font-bold text-[8pt] shrink-0">=</span>
                                    </div>
                                  </div>
                                  <div className="total-val flex-grow text-right pr-2.5 text-[8.5pt] font-mono font-semibold flex items-center justify-end px-1.5 py-0.5 leading-tight">
                                    {parsedTransportationFee.toLocaleString("en-US", { minimumFractionDigits: 2 })}
                                  </div>
                                </div>
                              </td>
                            </tr>
                            <tr className="align-stretch">
                              <td className="w-1/2 p-0 align-stretch">
                                <div className="flex flex-row items-stretch h-full min-h-[22px] w-full">
                                  <div className="total-lbl bg-indigo-50/40 w-[145px] sm:w-[155px] shrink-0 pr-1.5 text-right text-[8pt] font-black uppercase flex items-center justify-end gap-1 tracking-wider text-indigo-950">
                                    <span>Net Payable</span>
                                    <span className="font-black text-[8.5pt] shrink-0">=</span>
                                  </div>
                                  <div className="total-val flex-grow text-right pr-2.5 text-[9.5pt] font-mono font-black flex items-center justify-end px-1.5 py-0.5 leading-tight text-indigo-950">
                                    {grandTotal.toLocaleString("en-US", { minimumFractionDigits: 2 })}
                                  </div>
                                </div>
                              </td>
                            </tr>
                          </>
                        ) : (
                          <>
                            <tr className="align-stretch">
                              <td className="amount-words-container w-1/2 border-r border-slate-400 print:border-slate-600 p-1 bg-slate-50/50 text-left align-middle">
                                <span className="font-extrabold text-[6.5pt] text-slate-700 uppercase tracking-wider block mb-0.5">
                                  Amount in Words:
                                </span>
                                <span className="text-[7.5pt] font-mono italic text-black font-black uppercase leading-tight">
                                  {numberToWords(calculatedGrandTotal, currency || "Taka")}
                                </span>
                              </td>
                              <td className="w-1/2 p-0 align-stretch">
                                <div className="flex flex-row items-stretch h-full min-h-[22px] w-full">
                                  <div className="total-lbl bg-slate-50 w-[145px] sm:w-[155px] shrink-0 pr-1.5 text-right text-[7.5pt] font-bold uppercase flex items-center justify-end gap-1">
                                    <span>TOTAL</span>
                                    <span className="font-bold text-[8pt] shrink-0">=</span>
                                  </div>
                                  <div className="total-val flex-grow text-right pr-2.5 text-[8.5pt] font-mono font-black flex items-center justify-end px-1.5 py-0.5 leading-tight min-h-[22px]">
                                    {grandTotal.toLocaleString("en-US", { minimumFractionDigits: 2 })}
                                  </div>
                                </div>
                              </td>
                            </tr>
                          </>
                        )}
                      </tbody>
                    </table>
                  )}
                </div>
              </td>
            </tr>
          </tbody>
          <tfoot className="table-footer-group print:table-footer-group">
            <tr>
              <td className="border-none p-0 m-0">
                {/* Print spacer in tfoot reserves vertical space on the table so rows and totals never collide with the repeating fixed signature footer */}
                <div className="print-footer-spacer hidden print:block h-[115px] w-full" aria-hidden="true" />

                {/* Signature block removed from screen UI; preserved for printing only */}
                <div className="screen-sig-wrapper hidden print:hidden">
                  {renderSignatureBlock(false)}
                </div>
              </td>
            </tr>
          </tfoot>
        </table>

        {/* Repeating signature block on EVERY page while printing */}
        <div className="print-repeating-signature-footer hidden print:block pointer-events-none">
          {renderSignatureBlock(true)}
        </div>
      </div>
        </div>
      </div>

        {/* View 3: Records Archive */}
        {activeView === "saved-docs" && (
          <div className="p-4 sm:p-6 lg:p-8 max-w-5xl w-full mx-auto animate-in fade-in duration-200">
            <React.Suspense fallback={<div className="w-full max-w-[210mm] mx-auto bg-white rounded border border-slate-200 p-4 mt-4 animate-pulse h-24" />}>
              <SavedDocumentsPanel
                savedDocs={savedDocs}
                currentDocId={currentDocId}
                searchQuery={searchQuery}
                setSearchQuery={setSearchQuery}
                selectedTypeFilter={selectedTypeFilter}
                setSelectedTypeFilter={setSelectedTypeFilter}
                activeCompany={activeCompany}
                onSelectCompany={handleSwitchCompany}
                loadSavedDoc={(doc) => {
                  loadSavedDoc(doc);
                  setActiveView("editor");
                }}
                deleteSavedDoc={deleteSavedDoc}
                renameSavedDoc={renameSavedDoc}
                isPageMode={true}
                onSwitchPage={(page) => setActiveView(page === "saved-docs" ? "saved-docs" : "editor")}
              />
            </React.Suspense>
          </div>
        )}
      </div>

      {/* Cell right-click Menu context */}
      {contextMenu && contextMenu.visible && (
        <div 
          className="fixed bg-white border border-slate-200 rounded-lg shadow-xl py-1.5 w-64 z-[9999] text-xs text-slate-700"
          style={{ top: `${contextMenu.y}px`, left: `${contextMenu.x}px` }}
          onClick={(e) => e.stopPropagation()}
        >
          {contextMenu.colIndex !== undefined && contextMenu.colIndex >= 0 && contextMenu.colIndex <= 3 && (
            <>
              <button 
                onClick={() => {
                  clearSpecificCell(contextMenu.rowIndex, contextMenu.colIndex!);
                  setContextMenu(null);
                }}
                className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold text-slate-900"
              >
                Clear Cell Content
              </button>
              <div className="my-1 border-t border-slate-100"></div>
            </>
          )}

          <button 
            onClick={() => {
              toggleMergeSelectedRange();
              setContextMenu(null);
            }}
            disabled={!hasRangeSelection() && !getMergeRegionAt(contextMenu.rowIndex, contextMenu.colIndex)}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 disabled:opacity-40 font-bold text-slate-900"
          >
            {getMergeRegionAt(contextMenu.rowIndex, contextMenu.colIndex) ? "Unmerge Cells" : "Merge Selected Cells"}
          </button>

          <div className="my-1 border-t border-slate-100"></div>

          <button 
            onClick={() => {
              setIsExcelModalOpen(true);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-emerald-50 text-emerald-800 font-bold flex items-center gap-1.5"
          >
            <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-600" />
            <span>Paste / Import from Excel...</span>
          </button>

          <button 
            onClick={() => {
              unwrapAllDescriptions();
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-indigo-50 text-indigo-800 font-bold flex items-center gap-1.5"
          >
            <WrapText className="h-3.5 w-3.5 text-indigo-600" />
            <span>Unwrap Line Breaks (Continuous Flow)</span>
          </button>

          <div className="my-1 border-t border-slate-100"></div>

          <button 
            onClick={() => {
              insertMultipleRows(contextMenu.rowIndex, 'above', 1);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold"
          >
            Insert 1 Row Above
          </button>
          <button 
            onClick={() => {
              insertMultipleRows(contextMenu.rowIndex, 'below', 1);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold"
          >
            Insert 1 Row Below
          </button>
          <button 
            onClick={() => {
              insertMultipleRows(contextMenu.rowIndex, 'below', 5);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold text-indigo-700"
          >
            Insert +5 Rows Below
          </button>

          <div className="my-1 border-t border-slate-100"></div>

          <button 
            onClick={() => {
              if (contextMenu.rowIndex > 0) {
                moveRow(contextMenu.rowIndex, 'up');
                setContextMenu(null);
              }
            }}
            disabled={contextMenu.rowIndex === 0}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 disabled:opacity-40 font-bold"
          >
            Move Row Up
          </button>
          <button 
            onClick={() => {
              if (contextMenu.rowIndex < rows.length - 1) {
                moveRow(contextMenu.rowIndex, 'down');
                setContextMenu(null);
              }
            }}
            disabled={contextMenu.rowIndex === rows.length - 1}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 disabled:opacity-40 font-bold"
          >
            Move Row Down
          </button>

          <div className="my-1 border-t border-slate-100"></div>

          <button 
            onClick={() => {
              clearSpecificRow(contextMenu.rowIndex);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold"
          >
            Clear Row Content
          </button>
          <button 
            onClick={() => {
              deleteSpecificRow(contextMenu.rowIndex);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-rose-50 text-rose-600 font-bold border-t border-rose-50 mt-1"
          >
            Delete Row
          </button>
        </div>
      )}

      {/* Excel Smart Importer Modal */}
      {isExcelModalOpen && (
        <React.Suspense fallback={null}>
          <ExcelPasteModal
            isOpen={isExcelModalOpen}
            onClose={() => setIsExcelModalOpen(false)}
            onImportRows={handleImportFromExcel}
            selectedRowIndex={safeSelectedRowIndex}
            totalCurrentRows={rows.length}
            docType={docType}
          />
        </React.Suspense>
      )}

      {/* Floating Selection Formatting Toolbar */}
      <FloatingTextToolbar />

      {/* Floating Status Toast */}
      {toastMessage && (
        <div className="fixed bottom-20 right-6 z-[999999] bg-slate-900/95 backdrop-blur-xs text-white text-xs font-bold px-4 py-2.5 rounded-xl shadow-2xl border border-slate-700/80 flex items-center gap-2 animate-in fade-in slide-in-from-bottom-4 duration-250">
          <div className="bg-emerald-500 text-white rounded-full p-1">
            <CheckCheck className="h-3.5 w-3.5" />
          </div>
          <span>{toastMessage.text}</span>
        </div>
      )}
    </div>
  );
}
