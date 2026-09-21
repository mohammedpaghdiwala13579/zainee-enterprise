import React, { useState } from "react";
import { 
  FolderOpen, 
  Search, 
  X, 
  FileText, 
  Trash2, 
  Pencil, 
  Layers, 
  Check,
  FileEdit,
  FileSpreadsheet
} from "lucide-react";
import { SavedDocument, CompanyId } from "../types";
import { COMPANY_PROFILES } from "../utils/companyProfiles";
import { generateExcelDocument } from "../utils/excelGenerator";

export interface SavedDocumentsPanelProps {
  savedDocs: SavedDocument[];
  currentDocId: string | null;
  searchQuery: string;
  setSearchQuery: (val: string) => void;
  selectedTypeFilter: "all" | "quotation" | "challan" | "invoice";
  setSelectedTypeFilter: (val: "all" | "quotation" | "challan" | "invoice") => void;
  loadSavedDoc: (doc: SavedDocument) => void;
  deleteSavedDoc: (id: string, e?: React.MouseEvent) => void;
  renameSavedDoc: (id: string, e: React.MouseEvent) => void;
  // Multi-page navigation
  isPageMode?: boolean;
  onSwitchPage?: (page: "editor" | "saved-docs") => void;
  activeCompany?: CompanyId;
  onSelectCompany?: (company: CompanyId) => void;
}

export function calculateDocGrandTotal(doc: SavedDocument): { formatted: string; numeric: number } {
  if (doc.docType === "challan") {
    return { formatted: "—", numeric: 0 };
  }
  const rowsTotal = (doc.rows || []).reduce((sum: number, r: any) => sum + (Number(r.amount) || 0), 0);
  let discountAmount = 0;
  if (doc.includeDiscount) {
    const val = Number(doc.discountValue ?? doc.discountPercent) || 0;
    if (doc.discountType === "fixed") {
      discountAmount = Math.max(0, val);
    } else {
      discountAmount = (rowsTotal * Math.max(0, val)) / 100;
    }
  }
  discountAmount = Math.min(rowsTotal, Math.max(0, discountAmount));
  const netAfterDiscount = Math.max(0, rowsTotal - discountAmount);
  const vatPercent = Number(doc.vatPercent) || 0;
  const vatAmount = doc.docType === "invoice" ? (netAfterDiscount * vatPercent) / 100 : 0;
  const transFee = doc.docType === "invoice" ? (Number(doc.transportationFee) || 0) : 0;
  const grandTotal = doc.docType === "invoice" ? (netAfterDiscount + vatAmount + transFee) : rowsTotal;

  const rawCurrency = (doc.currency || "").trim().toUpperCase();
  // In the grand total column, write Taka in its currency symbol only (৳)
  const isTaka = !rawCurrency || rawCurrency === "TAKA" || rawCurrency === "BDT" || rawCurrency === "TK" || rawCurrency === "USD";
  const currencySymbol = isTaka ? "৳" : (doc.currency?.trim() || "৳");
  const currencyPrefix = `${currencySymbol} `;
  return {
    numeric: grandTotal,
    formatted: `${currencyPrefix}${grandTotal.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
  };
}

export default function SavedDocumentsPanel({
  savedDocs,
  currentDocId,
  searchQuery,
  setSearchQuery,
  selectedTypeFilter,
  setSelectedTypeFilter,
  loadSavedDoc,
  deleteSavedDoc,
  renameSavedDoc,
  isPageMode = false,
  onSwitchPage,
  activeCompany,
  onSelectCompany,
}: SavedDocumentsPanelProps) {
  const [exportingDocId, setExportingDocId] = useState<string | null>(null);

  const handleExportDocToExcel = async (doc: SavedDocument, e: React.MouseEvent) => {
    e.stopPropagation();
    setExportingDocId(doc.id);
    try {
      const company = COMPANY_PROFILES.zainee;
      await generateExcelDocument({
        currentCompany: company,
        docType: doc.docType,
        messers: doc.messers || "",
        address: doc.address || "",
        vesselName: doc.vesselName || "",
        portBerth: doc.portBerth || "",
        includeVesselName: doc.includeVesselName !== undefined ? doc.includeVesselName : true,
        includePortBerth: doc.includePortBerth !== undefined ? doc.includePortBerth : true,
        invoiceNo: doc.invoiceNo || "",
        challanNo: doc.challanNo || "",
        quotationNo: doc.quotationNo || "",
        requisitionNo: doc.requisitionNo || "",
        poNumber: doc.poNumber || "",
        includeInvoiceNo: doc.includeInvoiceNo !== undefined ? doc.includeInvoiceNo : true,
        includeChallanNo: doc.includeChallanNo !== undefined ? doc.includeChallanNo : true,
        includeQuotationNo: doc.includeQuotationNo !== undefined ? doc.includeQuotationNo : true,
        includeRequisitionNo: doc.includeRequisitionNo !== undefined ? doc.includeRequisitionNo : true,
        includePoNumber: doc.includePoNumber !== undefined ? doc.includePoNumber : true,
        dateVal: doc.dateVal || "",
        rows: doc.rows || [],
        includeDiscount: doc.includeDiscount,
        discountType: doc.discountType,
        discountValue: doc.discountValue ?? doc.discountPercent ?? 0,
        vatPercent: doc.vatPercent ?? 0,
        transportationFee: doc.transportationFee ?? 0,
        currency: doc.currency || "BDT",
        currencySymbol: doc.currencySymbol || "Tk",
      });
    } catch (err) {
      console.error("Direct Excel export error:", err);
    } finally {
      setExportingDocId(null);
    }
  };

  // Filter documents by search and document type
  const filteredDocs = savedDocs.filter((doc) => {
    if (selectedTypeFilter !== "all" && doc.docType !== selectedTypeFilter) {
      return false;
    }
    if (searchQuery.trim() !== "") {
      const q = searchQuery.toLowerCase().trim();
      const nameMatch = doc.name?.toLowerCase().includes(q);
      const messersMatch = doc.messers?.toLowerCase().includes(q);
      const addressMatch = doc.address?.toLowerCase().includes(q);
      const challanNoMatch = doc.challanNo?.toLowerCase().includes(q);
      const requisitionNoMatch = doc.requisitionNo?.toLowerCase().includes(q);
      const invoiceNoMatch = doc.invoiceNo?.toLowerCase().includes(q);
      const poNumberMatch = doc.poNumber?.toLowerCase().includes(q);
      const rowMatch = doc.rows?.some((r) => r.desc?.toLowerCase().includes(q));
      return nameMatch || messersMatch || addressMatch || challanNoMatch || requisitionNoMatch || invoiceNoMatch || poNumberMatch || rowMatch;
    }
    return true;
  });

  return (
    <div
      id="saved-documents-panel"
      className={`saved-docs-panel no-print print:hidden w-full bg-white text-slate-900 border border-slate-200/90 rounded-xl shadow-xs transition-all ${
        isPageMode
          ? "p-4 sm:p-5 max-w-[210mm] mx-auto mt-2"
          : "p-3 sm:p-4 max-w-[210mm] mx-auto mt-4"
      }`}
    >
      {/* Compact Top Header Bar */}
      <div className="flex flex-wrap items-center justify-between gap-2 pb-2.5 mb-2.5 border-b border-slate-100">
        <div className="flex items-center gap-2">
          <div className="p-1.5 bg-blue-50 text-blue-600 rounded-md shrink-0">
            <FolderOpen className="h-4 w-4" />
          </div>
          <div className="flex items-baseline gap-2">
            <h2 className="text-xs sm:text-sm font-semibold text-slate-900 tracking-tight">
              {isPageMode ? "Documents Archive" : "Saved Documents"}
            </h2>
            <span className="bg-slate-100 border border-slate-200 text-slate-700 text-[10px] font-mono font-medium px-2 py-0.5 rounded">
              {filteredDocs.length} {filteredDocs.length === 1 ? "record" : "records"}
            </span>
          </div>
        </div>

        {/* Minimal Actions & Page Navigation Controls */}
        <div className="flex items-center gap-1.5 text-xs" />
      </div>

      {/* Ultra-compact Search Bar & Filters */}
      <div className="flex flex-col gap-2 mb-3">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
          {/* Search Input */}
          <div className="relative flex-1">
            <Search className="absolute left-2.5 top-2.5 h-3.5 w-3.5 text-slate-400" />
            <input
              type="text"
              id="input-saved-docs-search"
              placeholder="Search by name, client, vessel, reference..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-8 pr-7 py-1.5 text-xs bg-slate-50 hover:bg-slate-100/70 focus:bg-white border border-slate-200 rounded text-slate-800 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-blue-500 transition-colors"
            />
            {searchQuery && (
              <button
                type="button"
                onClick={() => setSearchQuery("")}
                className="absolute right-2 top-2 text-slate-400 hover:text-slate-600 focus:outline-none cursor-pointer"
              >
                <X className="h-3.5 w-3.5" />
              </button>
            )}
          </div>

          {/* Minimal Type Filters */}
          <div className="flex items-center gap-1 bg-slate-100 p-0.5 rounded border border-slate-200 shrink-0">
            {(["all", "quotation", "challan", "invoice"] as const).map((type) => {
              const count =
                type === "all"
                  ? savedDocs.length
                  : savedDocs.filter((d) => d.docType === type).length;
              const isSelected = selectedTypeFilter === type;
              return (
                <button
                  key={type}
                  type="button"
                  id={`btn-filter-${type}`}
                  onClick={() => setSelectedTypeFilter(type)}
                  className={`px-2.5 py-1 rounded text-[10.5px] font-medium tracking-wide transition-colors cursor-pointer capitalize ${
                    isSelected
                      ? "bg-white text-blue-700 shadow-2xs font-semibold"
                      : "text-slate-600 hover:text-slate-900"
                  }`}
                >
                  {type === "all" ? "All" : type} ({count})
                </button>
              );
            })}
          </div>
        </div>
      </div>

      {/* Strictly 3-Column Documents Table (Name | Grand Total | Actions) */}
      {(() => {
        if (savedDocs.length === 0) {
          return (
            <div className="flex flex-col items-center justify-center py-10 px-4 text-center border border-dashed border-slate-200 rounded bg-slate-50/50">
              <FileText className="h-7 w-7 text-slate-300 stroke-[1.5]" />
              <h3 className="font-semibold text-xs text-slate-700 uppercase tracking-wider mt-2">No Saved Documents</h3>
              <p className="text-[11px] text-slate-400 mt-0.5 max-w-xs">
                Save your active sheet from the toolbar to access it here anytime.
              </p>
            </div>
          );
        }

        if (filteredDocs.length === 0) {
          return (
            <div className="flex flex-col items-center justify-center py-10 px-4 text-center border border-dashed border-slate-200 rounded bg-slate-50/50">
              <Search className="h-6 w-6 text-slate-300 stroke-[1.5]" />
              <h3 className="font-semibold text-xs text-slate-700 uppercase tracking-wider mt-2">No Matching Files</h3>
              <p className="text-[11px] text-slate-400 mt-0.5">
                No documents match the search string or filter selection.
              </p>
              <button
                type="button"
                onClick={() => {
                  setSearchQuery("");
                  setSelectedTypeFilter("all");
                }}
                className="mt-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium text-[11px] py-1 px-3 rounded border border-slate-200 transition-colors cursor-pointer"
              >
                Reset Filters
              </button>
            </div>
          );
        }

        return (
          <div
            className={`overflow-x-auto rounded border border-slate-200 bg-white ${
              !isPageMode ? "max-h-72 overflow-y-auto" : ""
            }`}
          >
            <table className="w-full text-left text-xs border-collapse">
              {/* Only 3 Columns: Name, Grand Total, Actions */}
              <thead className="bg-slate-50 text-[11px] font-semibold text-slate-600 uppercase tracking-wider border-b border-slate-200 sticky top-0 z-10">
                <tr>
                  <th className="px-3 py-2 text-left">Document & Client</th>
                  <th className="px-3 py-2 text-right w-36 sm:w-44">Grand Total</th>
                  <th className="px-3 py-2 text-right w-44 sm:w-56 pr-3">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredDocs.map((doc) => {
                  const isActive = currentDocId === doc.id;
                  const { formatted: totalDisplay } = calculateDocGrandTotal(doc);

                  return (
                    <tr
                      key={doc.id}
                      id={`saved-doc-row-${doc.id}`}
                      onClick={() => loadSavedDoc(doc)}
                      className={`hover:bg-slate-50/80 transition-colors cursor-pointer group ${
                        isActive ? "bg-blue-50/40 hover:bg-blue-50/60" : ""
                      }`}
                    >
                      {/* 1. Name Column */}
                      <td className="px-3 py-2.5 align-middle">
                        <div className="flex flex-col gap-0.5">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            {isActive && (
                              <span className="inline-flex items-center gap-1 bg-blue-100 text-blue-800 text-[9px] font-bold px-1.5 py-0.5 rounded tracking-wide">
                                <span className="h-1.5 w-1.5 rounded-full bg-blue-600 animate-pulse" />
                                ACTIVE
                              </span>
                            )}
                            <span
                              className="font-semibold text-slate-900 text-xs truncate max-w-[180px] sm:max-w-[280px]"
                              title={doc.name}
                            >
                              {doc.name}
                            </span>
                            <span
                              className={`inline-flex items-center rounded px-1.5 py-0.5 text-[9.5px] font-semibold uppercase tracking-wider border ${
                                doc.docType === "quotation"
                                  ? "bg-blue-50 border-blue-200 text-blue-700"
                                  : doc.docType === "challan"
                                  ? "bg-emerald-50 border-emerald-200 text-emerald-700"
                                  : "bg-slate-100 border-slate-200 text-slate-700"
                              }`}
                            >
                              {doc.docType}
                            </span>
                          </div>

                          {/* Client/date metadata in Name column */}
                          {(doc.messers || doc.vesselName || doc.dateVal) && (
                            <div className="flex items-center gap-2 text-[10.5px] text-slate-500 truncate max-w-[280px] sm:max-w-[380px]">
                              <span>{[doc.messers, doc.vesselName, doc.dateVal].filter(Boolean).join(" • ")}</span>
                            </div>
                          )}
                        </div>
                      </td>

                      {/* 2. Grand Total Column */}
                      <td className="px-3 py-2.5 text-right align-middle font-mono font-semibold text-xs text-slate-900">
                        {totalDisplay}
                      </td>

                      {/* 3. Actions Column */}
                      <td
                        className="px-3 py-2.5 text-right align-middle space-x-1 no-print whitespace-nowrap"
                        onClick={(e) => e.stopPropagation()}
                      >
                        {/* Open in Editor */}
                        <button
                          type="button"
                          id={`btn-open-doc-${doc.id}`}
                          onClick={() => loadSavedDoc(doc)}
                          className="inline-flex items-center gap-1 text-[11px] font-medium text-blue-600 hover:text-blue-800 hover:bg-blue-50 px-2 py-1 rounded transition-colors cursor-pointer"
                          title="Open in Document Editor"
                        >
                          <FileEdit className="h-3 w-3" />
                          <span>Open</span>
                        </button>

                        {/* Export to Excel */}
                        <button
                          type="button"
                          id={`btn-export-excel-${doc.id}`}
                          onClick={(e) => handleExportDocToExcel(doc, e)}
                          disabled={exportingDocId === doc.id}
                          className="inline-flex items-center gap-1 text-[11px] font-medium text-emerald-700 hover:text-emerald-900 hover:bg-emerald-50 px-2 py-1 rounded transition-colors cursor-pointer disabled:opacity-50"
                          title="Export directly to Excel (.xlsx)"
                        >
                          <FileSpreadsheet className="h-3 w-3 text-emerald-600" />
                          <span>{exportingDocId === doc.id ? "Exporting..." : "Excel"}</span>
                        </button>

                        {/* Rename */}
                        <button
                          type="button"
                          id={`btn-rename-doc-${doc.id}`}
                          onClick={(e) => renameSavedDoc(doc.id, e)}
                          className="inline-flex items-center gap-1 text-[11px] font-medium text-slate-500 hover:text-slate-800 hover:bg-slate-100 px-1.5 py-1 rounded transition-colors cursor-pointer"
                          title="Rename Document"
                        >
                          <Pencil className="h-3 w-3" />
                        </button>

                        {/* Delete */}
                        <button
                          type="button"
                          id={`btn-delete-doc-${doc.id}`}
                          onClick={(e) => deleteSavedDoc(doc.id, e)}
                          className="inline-flex items-center gap-1 text-[11px] font-medium text-rose-500 hover:text-rose-700 hover:bg-rose-50 px-1.5 py-1 rounded transition-colors cursor-pointer"
                          title="Delete Document"
                        >
                          <Trash2 className="h-3 w-3" />
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        );
      })()}
    </div>
  );
}
