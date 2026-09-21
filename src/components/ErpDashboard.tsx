import React, { useMemo, useState } from "react";
import { 
  FileText, 
  FileSpreadsheet, 
  Truck, 
  Ship, 
  Plus, 
  Search, 
  ArrowUpRight, 
  TrendingUp, 
  Clock, 
  CheckCircle2, 
  AlertCircle,
  Copy,
  Trash2,
  Printer,
  ExternalLink,
  DollarSign,
  FolderKanban,
  Building2
} from "lucide-react";
import { SavedDocument, CompanyId } from "../types";
import { COMPANY_PROFILES } from "../utils/companyProfiles";
import { calculateDocGrandTotal } from "./SavedDocumentsPanel";

export interface ErpDashboardProps {
  savedDocs: SavedDocument[];
  onOpenDoc: (doc: SavedDocument) => void;
  onNewDoc: (type?: "quotation" | "challan" | "invoice") => void;
  onDeleteDoc: (id: string, e?: React.MouseEvent) => void;
  onDuplicateDoc: (doc: SavedDocument) => void;
  onSwitchToArchive: () => void;
  activeCompany?: CompanyId;
  onSelectCompany?: (company: CompanyId) => void;
}

export default function ErpDashboard({
  savedDocs,
  onOpenDoc,
  onNewDoc,
  onDeleteDoc,
  onDuplicateDoc,
  onSwitchToArchive,
  activeCompany = "zainee",
  onSelectCompany,
}: ErpDashboardProps) {
  const displayedDocs = savedDocs;

  // Compute Operations KPIs
  const stats = useMemo(() => {
    let totalQuotations = 0;
    let totalChallans = 0;
    let totalInvoices = 0;
    let quotationValue = 0;
    let invoiceValue = 0;
    const vesselSet = new Set<string>();

    displayedDocs.forEach((doc) => {
      const { numeric } = calculateDocGrandTotal(doc);
      if (doc.vesselName && doc.vesselName.trim()) {
        vesselSet.add(doc.vesselName.trim());
      }

      if (doc.docType === "quotation") {
        totalQuotations += 1;
        quotationValue += numeric;
      } else if (doc.docType === "challan") {
        totalChallans += 1;
      } else if (doc.docType === "invoice") {
        totalInvoices += 1;
        invoiceValue += numeric;
      }
    });

    const totalDocs = displayedDocs.length;
    const activeVessels = vesselSet.size;

    return {
      totalDocs,
      totalQuotations,
      totalChallans,
      totalInvoices,
      quotationValue,
      invoiceValue,
      activeVessels,
    };
  }, [displayedDocs]);

  // Recent 8 documents
  const recentDocs = useMemo(() => {
    return [...displayedDocs]
      .sort((a, b) => new Date(b.updatedAt || 0).getTime() - new Date(a.updatedAt || 0).getTime())
      .slice(0, 8);
  }, [displayedDocs]);

  // Document percentage breakdown
  const docBreakdown = useMemo(() => {
    const total = stats.totalDocs || 1;
    const quotationPct = Math.round((stats.totalQuotations / total) * 100);
    const invoicePct = Math.round((stats.totalInvoices / total) * 100);
    const challanPct = Math.max(0, 100 - quotationPct - invoicePct);
    return { quotationPct, invoicePct, challanPct };
  }, [stats]);

  return (
    <div id="erp-dashboard-view" className="w-full max-w-7xl mx-auto space-y-6 animate-in fade-in duration-200">
      
      {/* 1. Header Section */}
      <div className="flex flex-col gap-3 pb-2 border-b border-slate-200">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight flex items-center gap-2">
              Operations & Documentation Overview
            </h1>
            <p className="text-xs text-slate-500 font-medium mt-0.5">
              Zainee Enterprise • Hardware, Tools &amp; General Order Supplier Hub
            </p>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              id="btn-dashboard-records-archive"
              onClick={onSwitchToArchive}
              className="h-8 px-2.5 bg-white hover:bg-slate-50 text-slate-700 hover:text-slate-900 border border-slate-300 rounded-md text-xs font-semibold flex items-center gap-1.5 shadow-2xs transition-colors cursor-pointer"
              title="Go to Records Archive"
              aria-label="Records Archive"
            >
              <FolderKanban className="h-3.5 w-3.5 text-blue-600" />
              <span className="hidden sm:inline">Records Archive</span>
              <span className="text-[10px] font-mono px-1.5 py-0.2 rounded bg-slate-100 text-slate-700 font-bold">
                {stats.totalDocs}
              </span>
            </button>

            <button
              type="button"
              id="btn-quick-new-quotation"
              onClick={() => onNewDoc("quotation")}
              className="h-8 px-3 bg-blue-600 hover:bg-blue-700 text-white rounded-md text-xs font-semibold flex items-center gap-1.5 shadow-xs transition-colors cursor-pointer"
            >
              <Plus className="h-3.5 w-3.5" />
              <span>New Document</span>
            </button>
          </div>
        </div>
      </div>

      {/* 2. Elegant KPI Cards (4-column grid) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        
        {/* KPI 1: Total Documents */}
        <div className="bg-white p-4 rounded-lg border border-slate-200 shadow-xs hover:border-slate-300 transition-all">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold text-slate-500 tracking-wide uppercase">Total Records</span>
            <div className="h-7 w-7 rounded-md bg-slate-100 flex items-center justify-center text-slate-600">
              <FileText className="h-3.5 w-3.5" />
            </div>
          </div>
          <div className="mt-3 flex items-baseline gap-2">
            <span className="text-2xl font-bold text-slate-900 tracking-tight font-mono">{stats.totalDocs}</span>
            <span className="text-[11px] text-slate-500">files archived</span>
          </div>
          <div className="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-[11px] text-slate-500">
            <span>{stats.totalQuotations} Quotations</span>
            <span>•</span>
            <span>{stats.totalInvoices} Invoices</span>
            <span>•</span>
            <span>{stats.totalChallans} Challans</span>
          </div>
        </div>

        {/* KPI 2: Total Invoiced Value */}
        <div className="bg-white p-4 rounded-lg border border-slate-200 shadow-xs hover:border-slate-300 transition-all">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold text-slate-500 tracking-wide uppercase">Invoices</span>
            <div className="h-7 w-7 rounded-md bg-emerald-50 text-emerald-700 flex items-center justify-center">
              <CheckCircle2 className="h-3.5 w-3.5" />
            </div>
          </div>
          <div className="mt-3 flex items-baseline gap-1.5">
            <span className="text-2xl font-bold text-slate-900 tracking-tight font-mono">
              ৳{stats.invoiceValue.toLocaleString("en-US", { maximumFractionDigits: 0 })}
            </span>
          </div>
          <div className="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500">{stats.totalInvoices} Invoices</span>
            <span className="text-emerald-700 font-medium">Invoiced Value</span>
          </div>
        </div>

        {/* KPI 3: Total Quotations Pipeline */}
        <div className="bg-white p-4 rounded-lg border border-slate-200 shadow-xs hover:border-slate-300 transition-all">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold text-slate-500 tracking-wide uppercase">Quotation Pipeline</span>
            <div className="h-7 w-7 rounded-md bg-blue-50 text-blue-600 flex items-center justify-center">
              <TrendingUp className="h-3.5 w-3.5" />
            </div>
          </div>
          <div className="mt-3 flex items-baseline gap-1.5">
            <span className="text-2xl font-bold text-slate-900 tracking-tight font-mono">
              ৳{stats.quotationValue.toLocaleString("en-US", { maximumFractionDigits: 0 })}
            </span>
          </div>
          <div className="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500">{stats.totalQuotations} active quotes</span>
            <span className="text-blue-700 font-medium">Pending Supply</span>
          </div>
        </div>

        {/* KPI 4: Active Vessels */}
        <div className="bg-white p-4 rounded-lg border border-slate-200 shadow-xs hover:border-slate-300 transition-all">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold text-slate-500 tracking-wide uppercase">Vessels Serviced</span>
            <div className="h-7 w-7 rounded-md bg-slate-100 text-slate-700 flex items-center justify-center">
              <Ship className="h-3.5 w-3.5" />
            </div>
          </div>
          <div className="mt-3 flex items-baseline gap-2">
            <span className="text-2xl font-bold text-slate-900 tracking-tight font-mono">{stats.activeVessels}</span>
            <span className="text-[11px] text-slate-500">merchant ships</span>
          </div>
          <div className="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500">{stats.totalChallans} Challans</span>
            <span className="text-slate-700 font-medium">Port &amp; Anchorage</span>
          </div>
        </div>

      </div>

      {/* 3. Distribution Bar & Quick Action Cards */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        
        {/* Document Breakdown Bar (2 cols) */}
        <div className="lg:col-span-2 bg-white p-4 rounded-lg border border-slate-200 shadow-xs">
          <div className="flex items-center justify-between mb-3">
            <div>
              <h2 className="text-sm font-bold text-slate-900">Document Volume Distribution</h2>
              <p className="text-xs text-slate-500">Distribution across active maritime paperwork types</p>
            </div>
            <span className="text-xs font-mono text-slate-500">{stats.totalDocs} total</span>
          </div>

          {/* Clean Segmented Bar */}
          <div className="w-full h-3 bg-slate-100 rounded-full overflow-hidden flex">
            <div 
              style={{ width: `${docBreakdown.quotationPct}%` }} 
              className="bg-blue-600 transition-all"
              title={`Quotations: ${docBreakdown.quotationPct}%`}
            />
            <div 
              style={{ width: `${docBreakdown.invoicePct}%` }} 
              className="bg-emerald-600 transition-all"
              title={`Invoices: ${docBreakdown.invoicePct}%`}
            />
            <div 
              style={{ width: `${docBreakdown.challanPct}%` }} 
              className="bg-slate-400 transition-all"
              title={`Challans: ${docBreakdown.challanPct}%`}
            />
          </div>

          {/* Legend */}
          <div className="mt-3 grid grid-cols-3 gap-2 text-xs">
            <div className="flex items-center gap-2">
              <span className="h-2.5 w-2.5 rounded-full bg-blue-600 shrink-0" />
              <div>
                <span className="text-slate-700 font-medium">Quotations: </span>
                <span className="text-slate-900 font-bold font-mono">{stats.totalQuotations}</span>
                <span className="text-slate-400 text-[10px] ml-1">({docBreakdown.quotationPct}%)</span>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <span className="h-2.5 w-2.5 rounded-full bg-emerald-600 shrink-0" />
              <div>
                <span className="text-slate-700 font-medium">Invoices: </span>
                <span className="text-slate-900 font-bold font-mono">{stats.totalInvoices}</span>
                <span className="text-slate-400 text-[10px] ml-1">({docBreakdown.invoicePct}%)</span>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <span className="h-2.5 w-2.5 rounded-full bg-slate-400 shrink-0" />
              <div>
                <span className="text-slate-700 font-medium">Challans: </span>
                <span className="text-slate-900 font-bold font-mono">{stats.totalChallans}</span>
                <span className="text-slate-400 text-[10px] ml-1">({docBreakdown.challanPct}%)</span>
              </div>
            </div>
          </div>
        </div>

        {/* Quick Launch Panel (1 col) */}
        <div className="bg-white p-4 rounded-lg border border-slate-200 shadow-xs flex flex-col justify-between">
          <div>
            <h2 className="text-sm font-bold text-slate-900 mb-1">Quick Document Generation</h2>
            <p className="text-xs text-slate-500 mb-3">Launch formal pre-configured templates</p>
            
            <div className="space-y-2">
              <button
                type="button"
                id="btn-quick-create-quotation"
                onClick={() => onNewDoc("quotation")}
                className="w-full text-left px-3 py-2 rounded-md border border-slate-200 hover:border-blue-300 hover:bg-blue-50/40 flex items-center justify-between text-xs font-semibold text-slate-800 transition-all cursor-pointer group"
              >
                <div className="flex items-center gap-2">
                  <FileText className="h-3.5 w-3.5 text-blue-600" />
                  <span>Quotation</span>
                </div>
                <ArrowUpRight className="h-3 w-3 text-slate-400 group-hover:text-blue-600 transition-colors" />
              </button>

              <button
                type="button"
                id="btn-quick-create-challan"
                onClick={() => onNewDoc("challan")}
                className="w-full text-left px-3 py-2 rounded-md border border-slate-200 hover:border-slate-300 hover:bg-slate-50 flex items-center justify-between text-xs font-semibold text-slate-800 transition-all cursor-pointer group"
              >
                <div className="flex items-center gap-2">
                  <Truck className="h-3.5 w-3.5 text-slate-600" />
                  <span>Challan</span>
                </div>
                <ArrowUpRight className="h-3 w-3 text-slate-400 group-hover:text-slate-900 transition-colors" />
              </button>

              <button
                type="button"
                id="btn-quick-create-invoice"
                onClick={() => onNewDoc("invoice")}
                className="w-full text-left px-3 py-2 rounded-md border border-slate-200 hover:border-emerald-300 hover:bg-emerald-50/40 flex items-center justify-between text-xs font-semibold text-slate-800 transition-all cursor-pointer group"
              >
                <div className="flex items-center gap-2">
                  <DollarSign className="h-3.5 w-3.5 text-emerald-600" />
                  <span>Invoice</span>
                </div>
                <ArrowUpRight className="h-3 w-3 text-slate-400 group-hover:text-emerald-600 transition-colors" />
              </button>
            </div>
          </div>
        </div>

      </div>

      {/* 4. Recent Documents Table */}
      <div className="bg-white rounded-lg border border-slate-200 shadow-xs overflow-hidden">
        <div className="p-4 border-b border-slate-200 flex items-center justify-between">
          <div>
            <h2 className="text-sm font-bold text-slate-900">Recent Documentation Records</h2>
            <p className="text-xs text-slate-500">Latest active documents saved across cloud database</p>
          </div>

          <button
            type="button"
            onClick={onSwitchToArchive}
            className="text-xs font-semibold text-blue-600 hover:text-blue-700 flex items-center gap-1 cursor-pointer transition-colors"
          >
            <span>View All Records ({stats.totalDocs})</span>
            <ArrowUpRight className="h-3 w-3" />
          </button>
        </div>

        {recentDocs.length === 0 ? (
          <div className="py-12 text-center text-slate-500 text-xs">
            <FileText className="h-8 w-8 text-slate-300 mx-auto mb-2" />
            <p className="font-semibold text-slate-700">No documents saved in database yet</p>
            <p className="text-slate-400 mt-1">Start by drafting a new quotation, challan, or invoice.</p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-[11px] font-semibold text-slate-500 uppercase tracking-wider">
                  <th className="py-2.5 px-4">Document / Title</th>
                  <th className="py-2.5 px-3">Type</th>
                  <th className="py-2.5 px-3">Client (Messers)</th>
                  <th className="py-2.5 px-3">Vessel</th>
                  <th className="py-2.5 px-3">Date</th>
                  <th className="py-2.5 px-4 text-right">Grand Total</th>
                  <th className="py-2.5 px-4 text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-xs text-slate-700">
                {recentDocs.map((doc) => {
                  const { formatted } = calculateDocGrandTotal(doc);
                  const isQuotation = doc.docType === "quotation";
                  const isInvoice = doc.docType === "invoice";
                  const isChallan = doc.docType === "challan";

                  return (
                    <tr 
                      key={doc.id}
                      className="hover:bg-slate-50/80 transition-colors group cursor-pointer"
                      onClick={() => onOpenDoc(doc)}
                    >
                      <td className="py-2.5 px-4 font-semibold text-slate-900">
                        <div className="flex items-center gap-1.5">
                          <FileText className="h-3.5 w-3.5 text-slate-400 shrink-0" />
                          <span className="truncate max-w-[200px]" title={doc.name}>
                            {doc.name || "Untitled Document"}
                          </span>
                        </div>
                      </td>

                      <td className="py-2.5 px-3">
                        <span
                          className={`inline-flex items-center px-2 py-0.5 rounded text-[10px] font-semibold uppercase tracking-wider ${
                            isQuotation
                              ? "bg-blue-50 text-blue-700 border border-blue-200"
                              : isInvoice
                              ? "bg-emerald-50 text-emerald-700 border border-emerald-200"
                              : "bg-slate-100 text-slate-700 border border-slate-300"
                          }`}
                        >
                          {doc.docType}
                        </span>
                      </td>

                      <td className="py-2.5 px-3 text-slate-600 truncate max-w-[180px]">
                        {doc.messers || "—"}
                      </td>

                      <td className="py-2.5 px-3 text-slate-600 font-mono text-[11px] truncate max-w-[140px]">
                        {doc.vesselName ? `M/V ${doc.vesselName}` : "—"}
                      </td>

                      <td className="py-2.5 px-3 text-slate-500 font-mono text-[11px]">
                        {doc.dateVal || "—"}
                      </td>

                      <td className="py-2.5 px-4 text-right font-mono font-semibold text-slate-900">
                        {isChallan ? <span className="text-slate-400 font-normal">Challan (No Total)</span> : formatted}
                      </td>

                      <td className="py-2.5 px-4 text-right" onClick={(e) => e.stopPropagation()}>
                        <div className="flex items-center justify-end gap-1">
                          <button
                            type="button"
                            onClick={() => onOpenDoc(doc)}
                            className="p-1 text-slate-400 hover:text-blue-600 rounded hover:bg-slate-100 transition-colors"
                            title="Open in Document Canvas"
                          >
                            <ExternalLink className="h-3.5 w-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={() => onDuplicateDoc(doc)}
                            className="p-1 text-slate-400 hover:text-slate-700 rounded hover:bg-slate-100 transition-colors"
                            title="Duplicate Document"
                          >
                            <Copy className="h-3.5 w-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={(e) => onDeleteDoc(doc.id, e)}
                            className="p-1 text-slate-400 hover:text-rose-600 rounded hover:bg-rose-50 transition-colors"
                            title="Delete Record"
                          >
                            <Trash2 className="h-3.5 w-3.5" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

    </div>
  );
}
