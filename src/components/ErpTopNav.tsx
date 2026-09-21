import React from "react";
import { 
  Menu, 
  Save, 
  Printer, 
  FolderKanban,
  Laptop,
  FileSpreadsheet
} from "lucide-react";
import { CompanyId } from "../types";

export interface ErpTopNavProps {
  activeView: "dashboard" | "editor" | "saved-docs";
  docType?: "quotation" | "challan" | "invoice";
  onSelectDocType?: (type: "quotation" | "challan" | "invoice") => void;
  currentDocName?: string;
  saveStatus?: "idle" | "saving" | "saved" | "error";
  lastSavedTime?: string | null;
  onSaveDoc: () => void;
  onPrint: () => void;
  onDownloadPDF?: () => void;
  isGeneratingPDF?: boolean;
  onDownloadExcel?: () => void;
  isGeneratingExcel?: boolean;
  onOpenExcelModal?: () => void;
  onToggleMobileSidebar: () => void;
  isInstallable?: boolean;
  onInstallClick?: () => void;
  onNavigateToArchive?: () => void;
  savedDocsCount?: number;
  activeCompany?: CompanyId;
  onSelectCompany?: (company: CompanyId) => void;
}

export default function ErpTopNav({
  activeView,
  docType,
  onSelectDocType,
  currentDocName,
  saveStatus,
  lastSavedTime,
  onSaveDoc,
  onPrint,
  onDownloadPDF,
  isGeneratingPDF = false,
  onDownloadExcel,
  isGeneratingExcel = false,
  onOpenExcelModal,
  onToggleMobileSidebar,
  isInstallable,
  onInstallClick,
  onNavigateToArchive,
  savedDocsCount,
  activeCompany = "zainee",
  onSelectCompany,
}: ErpTopNavProps) {
  return (
    <header
      id="erp-top-nav"
      className="sticky top-0 z-30 h-14 min-h-[56px] bg-white border-b border-slate-300 px-4 flex items-center justify-between no-print print:hidden shadow-2xs"
    >
      {/* Left: Mobile Menu Button + Breadcrumb + Business Switcher */}
      <div className="flex items-center gap-3 min-w-0">
        <button
          type="button"
          onClick={onToggleMobileSidebar}
          className="lg:hidden p-1.5 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-md cursor-pointer"
          title="Toggle Navigation Menu"
        >
          <Menu className="h-5 w-5" />
        </button>

        <div className="flex items-center gap-2 text-xs truncate">
          <span className="font-semibold text-slate-500 hidden sm:inline">
            Zainee Enterprise
          </span>
          <span className="text-slate-300 hidden sm:inline">/</span>
          <span className="font-bold text-slate-900 truncate">
            {activeView === "dashboard"
              ? "Operations Dashboard"
              : activeView === "saved-docs"
              ? "Records Archive"
              : currentDocName || "Document Canvas"}
          </span>
        </div>
      </div>

      {/* Right: Actions */}
      <div className="flex items-center gap-2 shrink-0">
        {/* Editor Action Buttons */}
        {activeView === "editor" && (
          <div className="flex items-center gap-1">
            {/* Download Excel (.xlsx) */}
            {onDownloadExcel && (
              <button
                type="button"
                id="topnav-btn-excel"
                onClick={onDownloadExcel}
                disabled={isGeneratingExcel}
                className="h-8 px-2.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 hover:text-emerald-900 border border-emerald-300 rounded-md text-xs font-semibold flex items-center gap-1.5 shadow-2xs transition-colors cursor-pointer disabled:opacity-50"
                title="Generate and download Excel (.xlsx) exactly as PDF"
              >
                <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-700" />
                <span>{isGeneratingExcel ? "Exporting..." : "Excel"}</span>
              </button>
            )}

            {/* Print / Save PDF */}
            <button
              type="button"
              id="topnav-btn-print"
              onClick={onPrint}
              className="h-8 px-2.5 bg-white hover:bg-slate-50 text-slate-700 hover:text-slate-900 border border-slate-300 rounded-md text-xs font-semibold flex items-center gap-1.5 shadow-2xs transition-colors cursor-pointer"
              title="Print Document or Save as PDF (Ctrl+P)"
            >
              <Printer className="h-3.5 w-3.5 text-slate-700" />
              <span className="hidden sm:inline">Print</span>
            </button>

            {/* Save to Cloud Button */}
            <button
              type="button"
              id="topnav-btn-save"
              onClick={onSaveDoc}
              disabled={saveStatus === "saving"}
              className="h-8 px-3 bg-blue-600 hover:bg-blue-700 text-white rounded-md text-xs font-semibold flex items-center gap-1.5 shadow-2xs transition-colors cursor-pointer disabled:opacity-60"
              title="Save Document Record"
            >
              <Save className="h-3.5 w-3.5" />
              <span>Save</span>
            </button>
          </div>
        )}

        {/* Records Archive Floating Action Button placed at the bottom right */}
        {onNavigateToArchive && (
          <button
            type="button"
            id="topnav-btn-records-archive"
            onClick={onNavigateToArchive}
            className={`fixed bottom-6 right-6 z-40 w-12 h-12 rounded-full flex items-center justify-center shadow-lg transition-all duration-200 cursor-pointer border no-print print:hidden ${
              activeView === "saved-docs"
                ? "bg-slate-900 hover:bg-slate-800 text-white border-slate-700 shadow-slate-950/30 ring-2 ring-blue-500"
                : "bg-blue-600 hover:bg-blue-700 text-white border-blue-500/50 shadow-blue-600/30 hover:scale-105 active:scale-95"
            }`}
            title={activeView === "saved-docs" ? "Back to Editor" : "Records Archive"}
            aria-label={activeView === "saved-docs" ? "Back to Editor" : "Records Archive"}
          >
            <FolderKanban className="h-5 w-5" />
          </button>
        )}

        {/* Install PWA Button if available */}
        {isInstallable && onInstallClick && (
          <button
            type="button"
            onClick={onInstallClick}
            className="h-8 px-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-md text-xs font-semibold flex items-center gap-1.5 transition-colors cursor-pointer"
            title="Install Standalone Desktop App"
          >
            <Laptop className="h-3.5 w-3.5 text-slate-600" />
            <span className="hidden sm:inline">Install</span>
          </button>
        )}

      </div>
    </header>
  );
}
