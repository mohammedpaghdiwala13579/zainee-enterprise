import React, { useState, useRef, useEffect } from "react";
import { 
  ChevronDown, 
  Check, 
  Plus, 
  Palette,
  X,
  Undo2,
  Redo2
} from "lucide-react";
import { CellFormat, CellBorders } from "../types";
import { applyInlineFormatting } from "../utils/textFormatter";

// Standard Excel Font Sizes Scale
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

// Standard Office Theme & Preset Fonts
const FONT_FAMILIES = [
  { name: "Calibri", family: "Calibri, Candara, Segoe, Segoe UI, Optima, Arial, sans-serif" },
  { name: "Arial", family: "Arial, Helvetica, sans-serif" },
  { name: "Times New Roman", family: "'Times New Roman', Times, serif" },
  { name: "Segoe UI", family: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif" },
  { name: "Georgia", family: "Georgia, serif" },
  { name: "Verdana", family: "Verdana, Geneva, sans-serif" },
  { name: "Tahoma", family: "Tahoma, Verdana, Segoe, sans-serif" },
  { name: "Courier New", family: "'Courier New', Courier, monospace" },
  { name: "Impact", family: "Impact, Haettenschweiler, 'Arial Narrow Bold', sans-serif" },
  { name: "Trebuchet MS", family: "'Trebuchet MS', 'Lucida Sans Unicode', sans-serif" },
  { name: "JetBrains Mono", family: "'JetBrains Mono', monospace" },
  { name: "Inter", family: "Inter, sans-serif" },
  { name: "Space Grotesk", family: "'Space Grotesk', sans-serif" }
];

// Excel Theme Colors (10 Base columns with 5 shade tints each)
const THEME_COLORS = [
  ["#FFFFFF", "#F2F2F2", "#D9D9D9", "#BFBFBF", "#A6A6A6", "#7F7F7F"], // White / Gray
  ["#000000", "#7F7F7F", "#595959", "#3F3F3F", "#262626", "#0C0C0C"], // Black
  ["#E7E6E6", "#D0CECE", "#AEAAAA", "#757171", "#3A3838", "#161616"], // Gray
  ["#44546A", "#D6DCE4", "#ACB9CA", "#8497B0", "#333F48", "#222A35"], // Dark Blue-Gray
  ["#4472C4", "#D9E1F2", "#B4C6E7", "#8EA9DB", "#305496", "#203764"], // Blue
  ["#ED7D31", "#FCE4D6", "#F8CBAD", "#F4B084", "#C65911", "#833C0C"], // Orange
  ["#A5A5A5", "#EDEDED", "#DBDBDB", "#C9C9C9", "#7B7B7B", "#525252"], // Gray 2
  ["#FFC000", "#FFF2CC", "#FFE699", "#FFD966", "#BF8F00", "#806000"], // Gold/Yellow
  ["#5B9BD5", "#DDEBF7", "#BDD7EE", "#9BC2E6", "#2E75B6", "#1F4E78"], // Light Blue
  ["#70AD47", "#E2EFDA", "#C6E0B4", "#A9D08E", "#548235", "#375623"], // Green
];

// Standard Row Colors
const STANDARD_COLORS = [
  "#C00000", // Dark Red
  "#FF0000", // Red
  "#FFC000", // Orange
  "#FFFF00", // Yellow
  "#92D050", // Light Green
  "#00B050", // Green
  "#00B0F0", // Light Blue
  "#0070C0", // Blue
  "#002060", // Dark Blue
  "#7030A0", // Purple
];

export interface ExcelRibbonToolbarProps {
  // Formatting
  currentFormat?: CellFormat;
  activeFormat?: CellFormat;
  onApplyFormat: (format: Partial<CellFormat>) => void;
  onApplyBorderPreset: (preset: string) => void;
  selectionSummary?: string;
  hasSelection?: boolean;
  canMerge?: boolean;
  onToggleMerge?: () => void;
  onClearFormatting?: () => void;

  // History Undo / Redo
  onUndo?: () => void;
  onRedo?: () => void;
  canUndo?: boolean;
  canRedo?: boolean;

  // Document Management
  docType: "quotation" | "challan" | "invoice";
  onSelectDocType: (type: "quotation" | "challan" | "invoice") => void;
  includeVesselName?: boolean;
  onToggleVesselName?: (val: boolean) => void;
  includePortBerth?: boolean;
  onTogglePortBerth?: (val: boolean) => void;
  includeDiscount?: boolean;
  onToggleDiscount?: (val: boolean) => void;
  autoSaveEnabled?: boolean;
  onToggleAutoSave?: (val: boolean) => void;
  lastSavedTime: string | null;
  currentDocId: string | null;
  currentDocName?: string;
  onCloseCurrentDoc?: () => void;

  // Operations
  onNewDoc: () => void;
  onDuplicateDoc?: () => void;
  onDeleteDoc?: () => void;
  onSaveDoc: () => void;
  saveStatus: "idle" | "saving" | "saved" | "error";

  // Actions
  onOpenExcelModal: () => void;
  onPrint: () => void;
  onDownloadPDF?: () => void;
  isGeneratingPDF?: boolean;
  onDownloadExcel?: () => void;
  isGeneratingExcel?: boolean;
}

export default function ExcelRibbonToolbar({
  currentFormat: propCurrentFormat,
  activeFormat,
  onApplyFormat,
  onApplyBorderPreset,
  selectionSummary,
  hasSelection: propHasSelection,
  canMerge,
  onToggleMerge,
  onClearFormatting,
  onUndo,
  onRedo,
  canUndo = false,
  canRedo = false,
  docType,
  onSelectDocType,
  includeVesselName = true,
  onToggleVesselName,
  includePortBerth = true,
  onTogglePortBerth,
  includeDiscount = false,
  onToggleDiscount,
  autoSaveEnabled,
  onToggleAutoSave,
  lastSavedTime,
  currentDocId,
  currentDocName,
  onCloseCurrentDoc,
  onNewDoc,
  onDuplicateDoc,
  onDeleteDoc,
  onSaveDoc,
  saveStatus,
  onOpenExcelModal,
  onPrint,
  onDownloadPDF,
  isGeneratingPDF,
  onDownloadExcel,
  isGeneratingExcel,
}: ExcelRibbonToolbarProps) {
  const currentFormat = activeFormat || propCurrentFormat || {};
  const hasSelection = propHasSelection !== undefined ? propHasSelection : true;

  // Dropdown visibility states
  const [openDropdown, setOpenDropdown] = useState<
    "fontFamily" | "fontSize" | "underline" | "borders" | "highlightColor" | "fontColor" | null
  >(null);

  // Active color memory (for single-click application)
  const [activeHighlightColor, setActiveHighlightColor] = useState<string>("#FFFF00");
  const [activeFontColor, setActiveFontColor] = useState<string>("#FF0000");
  const [lastBorderPreset, setLastBorderPreset] = useState<string>("all");

  // Custom size input state
  const [sizeInputVal, setSizeInputVal] = useState<string>(String(currentFormat.fontSize || 11));

  // Custom Color Pickers
  const highlightColorInputRef = useRef<HTMLInputElement>(null);
  const fontColorInputRef = useRef<HTMLInputElement>(null);
  const toolbarRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setSizeInputVal(String(currentFormat.fontSize || 11));
  }, [currentFormat.fontSize]);

  // Click outside to close any open dropdown
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (toolbarRef.current && !toolbarRef.current.contains(e.target as Node)) {
        setOpenDropdown(null);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const toggleDropdown = (name: typeof openDropdown) => {
    setOpenDropdown((prev) => (prev === name ? null : name));
  };

  // Font size step helpers
  const handleIncreaseFontSize = () => {
    const current = currentFormat.fontSize || 11;
    const next = FONT_SIZES.find((s) => s > current) || current + 2;
    const applied = applyInlineFormatting("fontSize", next);
    if (!applied) {
      onApplyFormat({ fontSize: next });
    }
    setSizeInputVal(String(next));
  };

  const handleDecreaseFontSize = () => {
    const current = currentFormat.fontSize || 11;
    const reversed = [...FONT_SIZES].reverse();
    const prev = reversed.find((s) => s < current) || Math.max(6, current - 2);
    const applied = applyInlineFormatting("fontSize", prev);
    if (!applied) {
      onApplyFormat({ fontSize: prev });
    }
    setSizeInputVal(String(prev));
  };

  const handleSizeInputSubmit = () => {
    const val = parseFloat(sizeInputVal);
    if (!isNaN(val) && val > 0 && val <= 144) {
      const applied = applyInlineFormatting("fontSize", val);
      if (!applied) {
        onApplyFormat({ fontSize: val });
      }
    } else {
      setSizeInputVal(String(currentFormat.fontSize || 11));
    }
  };

  // Font Family selector
  const activeFontFamilyObj = FONT_FAMILIES.find((f) => f.name.toLowerCase() === (currentFormat.fontFamily || "Calibri").toLowerCase()) || FONT_FAMILIES[0];

  // Borders handler
  const handleBorderSelect = (preset: string) => {
    setLastBorderPreset(preset);
    onApplyBorderPreset(preset);
    setOpenDropdown(null);
  };

  // Highlight / Fill Color handler
  const handleHighlightColorSelect = (color: string) => {
    setActiveHighlightColor(color);
    const applied = applyInlineFormatting("highlight", color);
    if (!applied) {
      // Only apply to cell if no editable text field is active
      onApplyFormat({ bgColor: color === "transparent" ? "" : color });
    }
    setOpenDropdown(null);
  };

  // Font Color handler
  const handleFontColorSelect = (color: string) => {
    setActiveFontColor(color);
    const applied = applyInlineFormatting("fontColor", color);
    if (!applied) {
      onApplyFormat({ color: color });
    }
    setOpenDropdown(null);
  };

  // Bold / Italic / Underline toggle handlers with inline text selection support
  const handleToggleBold = () => {
    const applied = applyInlineFormatting("bold");
    if (!applied) {
      onApplyFormat({ bold: !currentFormat.bold });
    }
  };

  const handleToggleItalic = () => {
    const applied = applyInlineFormatting("italic");
    if (!applied) {
      onApplyFormat({ italic: !currentFormat.italic });
    }
  };

  const handleToggleUnderline = (val?: "none" | "single" | "double") => {
    const applied = applyInlineFormatting("underline", val);
    if (!applied) {
      onApplyFormat({
        underline: val !== undefined ? val : (currentFormat.underline === "single" ? "none" : "single"),
      });
    }
  };

  return (
    <div
      ref={toolbarRef}
      id="excel-ribbon-toolbar"
      className="no-print w-full max-w-[210mm] bg-white text-slate-800 border border-slate-300 shadow-xs rounded-lg select-none font-sans text-xs px-2.5 py-1.5 mb-2 z-30 transition-all mx-auto"
    >
      {/* ========================================================================= */}
      {/* ROW 1: DOCUMENT TYPE & ACTION BUTTONS                                     */}
      {/* ========================================================================= */}
      <div className="flex flex-wrap items-center justify-between gap-1.5 pb-1.5 mb-1.5 border-b border-slate-300">
        
        {/* Left: Document Types & Status */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Document Type Tabs */}
          <div className="flex items-center bg-slate-100 p-0.5 rounded-md border border-slate-300">
            {(["quotation", "challan", "invoice"] as const).map((type) => (
              <button
                key={type}
                type="button"
                id={`btn-doctype-${type}`}
                onClick={() => onSelectDocType(type)}
                className={`px-2.5 py-1 rounded text-[10px] font-bold uppercase tracking-wider transition-all cursor-pointer ${
                  docType === type
                    ? "bg-white text-blue-700 shadow-2xs border border-slate-200/80"
                    : "text-slate-600 hover:text-slate-900 hover:bg-slate-200/50"
                }`}
              >
                {type}
              </button>
            ))}
          </div>

          {/* Active Document Indicator */}
          {currentDocId && (
            <div className="flex items-center gap-1.5 bg-blue-50/60 border border-blue-200/80 rounded-md px-2 py-1 text-[11px]">
              <span className="h-1.5 w-1.5 rounded-full bg-blue-600 animate-pulse" />
              <span className="text-blue-900 font-medium truncate max-w-[130px] sm:max-w-[200px]">
                {currentDocName || "Active Sheet"}
              </span>
              {onCloseCurrentDoc && (
                <button
                  type="button"
                  onClick={onCloseCurrentDoc}
                  className="text-blue-500 hover:text-blue-800 p-0.5 hover:bg-blue-100 rounded cursor-pointer"
                  title="Close current document"
                >
                  <X className="h-3 w-3" />
                </button>
              )}
            </div>
          )}
        </div>

        {/* Right: Document Actions */}
        <div className="flex flex-wrap items-center gap-1.5">
          {/* New Sheet */}
          <button
            type="button"
            id="btn-new-sheet"
            onClick={onNewDoc}
            className="h-7 px-2.5 bg-white hover:bg-slate-50 text-slate-700 hover:text-slate-900 rounded-md border border-slate-300 shadow-2xs transition-colors cursor-pointer flex items-center gap-1 text-xs font-semibold"
            title="Create New Blank Sheet"
          >
            <Plus className="h-3.5 w-3.5 text-blue-600" />
            <span className="hidden sm:inline">New</span>
          </button>
        </div>
      </div>

      {/* ========================================================================= */}
      {/* ROW 2: SPREADSHEET FORMATTING CONTROLS                                    */}
      {/* ========================================================================= */}
      <div className="flex flex-wrap items-center justify-between gap-x-1.5 gap-y-1">
        {/* Formatting Tools Left Section */}
        <div className="flex flex-wrap items-center gap-x-1 gap-y-0.5">

          {/* GROUP 0: QUICK ACCESS UNDO & REDO */}
          {(onUndo || onRedo) && (
            <div className="flex items-center gap-0.5 border-r border-slate-200 pr-1 mr-0.5">
              <button
                type="button"
                id="btn-toolbar-undo"
                onClick={onUndo}
                disabled={!canUndo}
                className={`h-6 w-6 rounded flex items-center justify-center transition-colors cursor-pointer ${
                  canUndo
                    ? "bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200"
                    : "text-slate-400 cursor-not-allowed border border-transparent opacity-50"
                }`}
                title="Undo (Ctrl+Z)"
              >
                <Undo2 className="h-3 w-3 text-slate-600" />
              </button>
              <button
                type="button"
                id="btn-toolbar-redo"
                onClick={onRedo}
                disabled={!canRedo}
                className={`h-6 w-6 rounded flex items-center justify-center transition-colors cursor-pointer ${
                  canRedo
                    ? "bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200"
                    : "text-slate-400 cursor-not-allowed border border-transparent opacity-50"
                }`}
                title="Redo (Ctrl+Y or Ctrl+Shift+Z)"
              >
                <Redo2 className="h-3 w-3 text-slate-600" />
              </button>
            </div>
          )}
          
          {/* GROUP 1: FONT FAMILY & FONT SIZE */}
          <div className="flex items-center gap-0.5">
            {/* Font Family Dropdown */}
            <div className="relative">
              <button
                type="button"
                id="btn-font-family"
                onClick={() => toggleDropdown("fontFamily")}
                className="h-6 w-[94px] bg-slate-50 hover:bg-slate-100 border border-slate-300 rounded px-1.5 flex items-center justify-between text-slate-800 text-[10.5px] font-normal transition-colors cursor-pointer"
                title="Font Family"
              >
                <span className="truncate" style={{ fontFamily: activeFontFamilyObj.family }}>
                  {activeFontFamilyObj.name}
                </span>
                <ChevronDown className="h-2.5 w-2.5 text-slate-400 ml-0.5 shrink-0" />
              </button>

              {openDropdown === "fontFamily" && (
                <div
                  onMouseDown={(e) => e.preventDefault()}
                  className="absolute left-0 top-[26px] w-[170px] max-h-[240px] overflow-y-auto bg-white border border-slate-200 shadow-xl rounded-md py-1 z-50 text-slate-800"
                >
                  <div className="px-2 py-0.5 text-[9px] font-bold text-slate-400 uppercase tracking-wider">
                    Available Fonts
                  </div>
                  {FONT_FAMILIES.map((font) => (
                    <button
                      key={font.name}
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => {
                        const applied = applyInlineFormatting("fontFamily", font.name);
                        if (!applied) {
                          onApplyFormat({ fontFamily: font.name });
                        }
                        setOpenDropdown(null);
                      }}
                      className={`w-full px-2 py-1 text-left text-xs flex items-center justify-between hover:bg-slate-50 transition-colors cursor-pointer ${
                        (currentFormat.fontFamily || "Calibri").toLowerCase() === font.name.toLowerCase()
                          ? "bg-blue-50 text-blue-700 font-bold"
                          : "text-slate-700"
                      }`}
                      style={{ fontFamily: font.family }}
                    >
                      <span className="truncate">{font.name}</span>
                      {(currentFormat.fontFamily || "Calibri").toLowerCase() === font.name.toLowerCase() && (
                        <Check className="h-3 w-3 text-blue-600 shrink-0" />
                      )}
                    </button>
                  ))}
                </div>
              )}
            </div>

            {/* Font Size Selector */}
            <div className="relative">
              <div className="h-6 w-[38px] bg-slate-50 border border-slate-300 rounded flex items-center justify-between overflow-hidden">
                <input
                  id="input-font-size"
                  type="text"
                  value={sizeInputVal}
                  onChange={(e) => setSizeInputVal(e.target.value)}
                  onBlur={handleSizeInputSubmit}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();
                      handleSizeInputSubmit();
                    }
                  }}
                  className="w-[22px] h-full bg-transparent text-center text-slate-800 text-[10px] font-mono outline-none p-0"
                  title="Font Size"
                />
                <button
                  type="button"
                  id="btn-font-size-dropdown"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => toggleDropdown("fontSize")}
                  className="h-full px-0.5 hover:bg-slate-200 text-slate-500 flex items-center justify-center cursor-pointer"
                  title="Select Font Size"
                >
                  <ChevronDown className="h-2 w-2" />
                </button>
              </div>

              {openDropdown === "fontSize" && (
                <div
                  onMouseDown={(e) => e.preventDefault()}
                  className="absolute left-0 top-[26px] w-[52px] max-h-[200px] overflow-y-auto bg-white border border-slate-200 shadow-xl rounded-md py-1 z-50 text-center"
                >
                  {FONT_SIZES.map((size) => (
                    <button
                      key={size}
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => {
                        const applied = applyInlineFormatting("fontSize", size);
                        if (!applied) {
                          onApplyFormat({ fontSize: size });
                        }
                        setSizeInputVal(String(size));
                        setOpenDropdown(null);
                      }}
                      className={`w-full py-0.5 text-xs hover:bg-slate-50 transition-colors cursor-pointer font-mono ${
                        (currentFormat.fontSize || 11) === size
                          ? "bg-blue-50 text-blue-700 font-bold"
                          : "text-slate-700"
                      }`}
                    >
                      {size}
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>

          {/* Divider */}
          <div className="h-3.5 w-[1px] bg-slate-200 mx-0.5" />

          {/* GROUP 2: BOLD, ITALIC, UNDERLINE, BORDERS, FILL & FONT COLOR */}
          <div className="flex items-center gap-0.5">
            {/* Bold */}
            <button
              type="button"
              id="btn-bold"
              onMouseDown={(e) => e.preventDefault()}
              onClick={handleToggleBold}
              className={`h-6 w-6 rounded flex items-center justify-center text-[10.5px] font-bold font-serif transition-colors cursor-pointer ${
                currentFormat.bold
                  ? "bg-blue-50 text-blue-700 border border-blue-300 shadow-2xs"
                  : "hover:bg-slate-100 text-slate-700 border border-transparent"
              }`}
              title="Bold (Ctrl+B)"
            >
              B
            </button>

            {/* Italic */}
            <button
              type="button"
              id="btn-italic"
              onMouseDown={(e) => e.preventDefault()}
              onClick={handleToggleItalic}
              className={`h-6 w-6 rounded flex items-center justify-center text-[10.5px] italic font-serif transition-colors cursor-pointer ${
                currentFormat.italic
                  ? "bg-blue-50 text-blue-700 border border-blue-300 shadow-2xs"
                  : "hover:bg-slate-100 text-slate-700 border border-transparent"
              }`}
              title="Italic (Ctrl+I)"
            >
              I
            </button>

            {/* Underline Split Button */}
            <div className="relative flex items-center">
              <button
                type="button"
                id="btn-underline"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => handleToggleUnderline()}
                className={`h-6 px-1 rounded-l flex items-center justify-center text-[10.5px] underline transition-colors cursor-pointer ${
                  currentFormat.underline && currentFormat.underline !== "none"
                    ? "bg-blue-50 text-blue-700 border-l border-t border-b border-blue-300 shadow-2xs"
                    : "hover:bg-slate-100 text-slate-700 border border-transparent"
                }`}
                title="Underline (Ctrl+U)"
              >
                <span className={currentFormat.underline === "double" ? "underline decoration-double" : "underline"}>
                  U
                </span>
              </button>
              <button
                type="button"
                id="btn-underline-dropdown"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => toggleDropdown("underline")}
                className={`h-6 px-0.5 rounded-r flex items-center justify-center transition-colors cursor-pointer ${
                  currentFormat.underline && currentFormat.underline !== "none"
                    ? "bg-blue-50 text-blue-700 border-r border-t border-b border-blue-300"
                    : "hover:bg-slate-100 text-slate-500"
                }`}
                title="Underline Options"
              >
                <ChevronDown className="h-2 w-2" />
              </button>

              {openDropdown === "underline" && (
                <div className="absolute left-0 top-[26px] w-[140px] bg-white border border-slate-200 shadow-xl rounded-md py-1 z-50">
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => {
                      handleToggleUnderline("single");
                      setOpenDropdown(null);
                    }}
                    className="w-full px-2.5 py-1 text-left text-xs hover:bg-slate-50 flex items-center justify-between text-slate-800 cursor-pointer"
                  >
                    <span className="underline">Underline</span>
                    {currentFormat.underline === "single" && <Check className="h-3 w-3 text-blue-600" />}
                  </button>
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => {
                      handleToggleUnderline("double");
                      setOpenDropdown(null);
                    }}
                    className="w-full px-2.5 py-1 text-left text-xs hover:bg-slate-50 flex items-center justify-between text-slate-800 cursor-pointer"
                  >
                    <span className="underline decoration-double">Double Underline</span>
                    {currentFormat.underline === "double" && <Check className="h-3 w-3 text-blue-600" />}
                  </button>
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => {
                      handleToggleUnderline("none");
                      setOpenDropdown(null);
                    }}
                    className="w-full px-2.5 py-1 text-left text-xs hover:bg-slate-50 flex items-center justify-between text-slate-500 cursor-pointer"
                  >
                    <span>None</span>
                    {(!currentFormat.underline || currentFormat.underline === "none") && <Check className="h-3 w-3 text-blue-600" />}
                  </button>
                </div>
              )}
            </div>

            {/* Borders Preset Dropdown */}
            <div className="relative flex items-center">
              <button
                type="button"
                id="btn-borders-quick"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => handleBorderSelect(lastBorderPreset)}
                className="h-6 px-1 rounded-l hover:bg-slate-100 flex items-center justify-center text-slate-700 transition-colors cursor-pointer"
                title={`Apply ${lastBorderPreset.replace(/_/g, " ")} Border`}
              >
                <svg className="h-3 w-3" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5">
                  <rect x="2" y="2" width="12" height="12" rx="0.5" strokeDasharray="2 2" />
                  <line x1="2" y1="14" x2="14" y2="14" stroke="currentColor" strokeWidth="2" />
                  <line x1="2" y1="8" x2="14" y2="8" strokeDasharray="2 2" />
                  <line x1="8" y1="2" x2="8" y2="14" strokeDasharray="2 2" />
                </svg>
              </button>
              <button
                type="button"
                id="btn-borders-dropdown"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => toggleDropdown("borders")}
                className="h-6 px-0.5 rounded-r hover:bg-slate-100 text-slate-500 flex items-center justify-center cursor-pointer"
                title="Borders"
              >
                <ChevronDown className="h-2 w-2" />
              </button>

              {openDropdown === "borders" && (
                <div className="absolute left-0 top-[26px] w-[200px] max-h-[280px] overflow-y-auto bg-white border border-slate-200 shadow-xl rounded-md py-1 z-50 text-xs">
                  <div className="px-3 py-1 text-[9px] font-bold text-slate-400 uppercase tracking-wider">
                    Border Presets
                  </div>
                  {[
                    { id: "bottom", label: "Bottom Border", desc: "Single bottom line" },
                    { id: "top", label: "Top Border", desc: "Single top line" },
                    { id: "left", label: "Left Border", desc: "Single left line" },
                    { id: "right", label: "Right Border", desc: "Single right line" },
                    { id: "none", label: "No Border", desc: "Clear all cell borders" },
                    { id: "all", label: "All Borders", desc: "Grid borders around all cells" },
                    { id: "outside", label: "Outside Borders", desc: "Border around perimeter" },
                    { id: "thick_outside", label: "Thick Outside Borders", desc: "Heavy outer border" },
                    { id: "bottom_double", label: "Bottom Double Border", desc: "Accounting style" },
                    { id: "thick_bottom", label: "Thick Bottom Border", desc: "Bold underline" },
                    { id: "top_and_bottom", label: "Top and Bottom Border", desc: "Header & sum borders" },
                    { id: "top_and_thick_bottom", label: "Top and Thick Bottom", desc: "Total accent" },
                    { id: "top_and_double_bottom", label: "Top and Double Bottom", desc: "Financial total double" },
                  ].map((preset) => (
                    <button
                      key={preset.id}
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => handleBorderSelect(preset.id)}
                      className="w-full px-2.5 py-1 text-left hover:bg-slate-50 flex items-center justify-between text-slate-800 cursor-pointer"
                    >
                      <div className="flex flex-col">
                        <span className="font-medium text-[11px]">{preset.label}</span>
                        <span className="text-[9px] text-slate-400">{preset.desc}</span>
                      </div>
                      {lastBorderPreset === preset.id && <Check className="h-3 w-3 text-blue-600 shrink-0" />}
                    </button>
                  ))}
                </div>
              )}
            </div>

            {/* Fill / Highlight Color Split Button */}
            <div className="relative flex items-center">
              <button
                type="button"
                id="btn-highlight-quick"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => handleHighlightColorSelect(activeHighlightColor)}
                className="h-6 px-1 rounded-l hover:bg-slate-100 flex flex-col items-center justify-center transition-colors cursor-pointer"
                title="Highlight / Fill Color"
              >
                <svg className="h-2.5 w-2.5 text-slate-700" viewBox="0 0 16 16" fill="currentColor">
                  <path d="M12.854.146a.5.5 0 0 0-.707 0L10.5 1.793 14.207 5.5l1.647-1.646a.5.5 0 0 0 0-.708l-3-3zm.646 6.061L9.793 2.5 3.293 9H3.5a.5.5 0 0 1 .5.5v.5h.5a.5.5 0 0 1 .5.5v.5h.5a.5.5 0 0 1 .5.5v.5h.5a.5.5 0 0 1 .5.5v.207l6.5-6.5zm-7.468 7.468A.5.5 0 0 1 6 13.5V13h-.5a.5.5 0 0 1-.5-.5V12h-.5a.5.5 0 0 1-.5-.5V11h-.5a.5.5 0 0 1-.5-.5V10h-.5a.499.499 0 0 1-.175-.032l-.179.178a.5.5 0 0 0-.11.168l-2 5a.5.5 0 0 0 .65.65l5-2a.5.5 0 0 0 .168-.11l.178-.178z" />
                </svg>
                <div
                  className="h-[2px] w-3 rounded-xs mt-0.5 shadow-xs"
                  style={{ backgroundColor: currentFormat.bgColor || activeHighlightColor }}
                />
              </button>
              <button
                type="button"
                id="btn-highlight-dropdown"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => toggleDropdown("highlightColor")}
                className="h-6 px-0.5 rounded-r hover:bg-slate-100 text-slate-500 flex items-center justify-center cursor-pointer"
                title="Highlight / Fill Color Palette"
              >
                <ChevronDown className="h-2 w-2" />
              </button>

              {openDropdown === "highlightColor" && (
                <div
                  onMouseDown={(e) => e.preventDefault()}
                  className="absolute left-0 top-[26px] w-[200px] bg-white border border-slate-200 shadow-xl rounded-md p-2 z-50 text-slate-800"
                >
                  <div className="flex items-center justify-between mb-1 pb-1 border-b border-slate-100">
                    <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Highlight Colors</span>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => handleHighlightColorSelect("transparent")}
                      className="text-[9px] text-blue-600 hover:underline cursor-pointer font-semibold"
                      title="Reset highlight for selected word or text"
                    >
                      No Fill (Reset)
                    </button>
                  </div>
                  <div className="grid grid-cols-10 gap-0.5 mb-1.5">
                    {THEME_COLORS.map((colGroup, colIdx) => (
                      <div key={colIdx} className="flex flex-col gap-0.5">
                        {colGroup.map((color, rowIdx) => (
                          <button
                            key={rowIdx}
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleHighlightColorSelect(color)}
                            className="h-3.5 w-3.5 rounded-xs border border-slate-300/80 hover:scale-125 transition-transform cursor-pointer"
                            style={{ backgroundColor: color }}
                            title={color}
                          />
                        ))}
                      </div>
                    ))}
                  </div>
                  <div className="text-[9px] font-bold text-slate-400 uppercase tracking-wider mb-1">
                    Standard Colors
                  </div>
                  <div className="flex items-center justify-between gap-0.5 mb-1.5">
                    {STANDARD_COLORS.map((color) => (
                      <button
                        key={color}
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={() => handleHighlightColorSelect(color)}
                        className="h-3.5 w-3.5 rounded-xs border border-slate-300/80 hover:scale-125 transition-transform cursor-pointer"
                        style={{ backgroundColor: color }}
                        title={color}
                      />
                    ))}
                  </div>
                  <div className="pt-1 border-t border-slate-100 flex items-center justify-between">
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => highlightColorInputRef.current?.click()}
                      className="text-[9px] text-slate-600 hover:text-slate-900 flex items-center gap-1 cursor-pointer font-medium"
                    >
                      <Palette className="h-2.5 w-2.5 text-blue-600" />
                      <span>Custom Color...</span>
                    </button>
                    <input
                      ref={highlightColorInputRef}
                      type="color"
                      className="sr-only"
                      onChange={(e) => handleHighlightColorSelect(e.target.value)}
                    />
                  </div>
                </div>
              )}
            </div>

            {/* Font / Text Color Split Button */}
            <div className="relative flex items-center">
              <button
                type="button"
                id="btn-fontcolor-quick"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => handleFontColorSelect(activeFontColor)}
                className="h-6 px-1 rounded-l hover:bg-slate-100 flex flex-col items-center justify-center transition-colors cursor-pointer"
                title="Font / Text Color"
              >
                <span className="text-[10px] font-bold leading-none text-slate-800">A</span>
                <div
                  className="h-[2px] w-3 rounded-xs mt-0.5 shadow-xs"
                  style={{ backgroundColor: currentFormat.color || activeFontColor }}
                />
              </button>
              <button
                type="button"
                id="btn-fontcolor-dropdown"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => toggleDropdown("fontColor")}
                className="h-6 px-0.5 rounded-r hover:bg-slate-100 text-slate-500 flex items-center justify-center cursor-pointer"
                title="Font Color Palette"
              >
                <ChevronDown className="h-2 w-2" />
              </button>

              {openDropdown === "fontColor" && (
                <div
                  onMouseDown={(e) => e.preventDefault()}
                  className="absolute left-0 top-[26px] w-[200px] bg-white border border-slate-200 shadow-xl rounded-md p-2 z-50 text-slate-800"
                >
                  <div className="flex items-center justify-between mb-1 pb-1 border-b border-slate-100">
                    <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Theme Colors</span>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => handleFontColorSelect("#000000")}
                      className="text-[9px] text-blue-600 hover:underline cursor-pointer font-semibold"
                      title="Reset font color to default"
                    >
                      Automatic (Reset)
                    </button>
                  </div>
                  <div className="grid grid-cols-10 gap-0.5 mb-1.5">
                    {THEME_COLORS.map((colGroup, colIdx) => (
                      <div key={colIdx} className="flex flex-col gap-0.5">
                        {colGroup.map((color, rowIdx) => (
                          <button
                            key={rowIdx}
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleFontColorSelect(color)}
                            className="h-3.5 w-3.5 rounded-xs border border-slate-300/80 hover:scale-125 transition-transform cursor-pointer"
                            style={{ backgroundColor: color }}
                            title={color}
                          />
                        ))}
                      </div>
                    ))}
                  </div>
                  <div className="text-[9px] font-bold text-slate-400 uppercase tracking-wider mb-1">
                    Standard Colors
                  </div>
                  <div className="flex items-center justify-between gap-0.5 mb-1.5">
                    {STANDARD_COLORS.map((color) => (
                      <button
                        key={color}
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={() => handleFontColorSelect(color)}
                        className="h-3.5 w-3.5 rounded-xs border border-slate-300/80 hover:scale-125 transition-transform cursor-pointer"
                        style={{ backgroundColor: color }}
                        title={color}
                      />
                    ))}
                  </div>
                  <div className="pt-1 border-t border-slate-100 flex items-center justify-between">
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => fontColorInputRef.current?.click()}
                      className="text-[9px] text-slate-600 hover:text-slate-900 flex items-center gap-1 cursor-pointer font-medium"
                    >
                      <Palette className="h-2.5 w-2.5 text-blue-600" />
                      <span>Custom Color...</span>
                    </button>
                    <input
                      ref={fontColorInputRef}
                      type="color"
                      className="sr-only"
                      onChange={(e) => handleFontColorSelect(e.target.value)}
                    />
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Divider */}
          <div className="h-3.5 w-[1px] bg-slate-200 mx-0.5" />

          {/* GROUP 3: ALIGNMENT (VALIGN & HALIGN) */}
          <div className="flex items-center gap-0.5">
            {/* Top Align */}
            <button
              type="button"
              id="btn-valign-top"
              onClick={() => onApplyFormat({ valign: "top" })}
              className={`h-6 w-6 rounded flex items-center justify-center transition-colors cursor-pointer ${
                currentFormat.valign === "top"
                  ? "bg-blue-50 text-blue-700 border border-blue-300 shadow-2xs"
                  : "hover:bg-slate-100 text-slate-600 border border-transparent"
              }`}
              title="Top Align"
            >
              <svg className="h-3 w-3" viewBox="0 0 16 16" fill="currentColor">
                <path d="M1 2h14v1.5H1V2zm2 3.5h10V7H3V5.5zm0 3.5h7v1.5H3V9z" />
              </svg>
            </button>

            {/* Middle Align */}
            <button
              type="button"
              id="btn-valign-middle"
              onClick={() => onApplyFormat({ valign: "middle" })}
              className={`h-6 w-6 rounded flex items-center justify-center transition-colors cursor-pointer ${
                currentFormat.valign === "middle" || !currentFormat.valign
                  ? "bg-blue-50 text-blue-700 border border-blue-300 shadow-2xs"
                  : "hover:bg-slate-100 text-slate-600 border border-transparent"
              }`}
              title="Middle Align"
            >
              <svg className="h-3 w-3" viewBox="0 0 16 16" fill="currentColor">
                <path d="M1 7h14v1.5H1V7zm3-3.5h8V5H4V3.5zm0 7h8V12H4v-1.5z" />
              </svg>
            </button>

            {/* Bottom Align */}
            <button
              type="button"
              id="btn-valign-bottom"
              onClick={() => onApplyFormat({ valign: "bottom" })}
              className={`h-6 w-6 rounded flex items-center justify-center transition-colors cursor-pointer ${
                currentFormat.valign === "bottom"
                  ? "bg-blue-50 text-blue-700 border border-blue-300 shadow-2xs"
                  : "hover:bg-slate-100 text-slate-600 border border-transparent"
              }`}
              title="Bottom Align"
            >
              <svg className="h-3 w-3" viewBox="0 0 16 16" fill="currentColor">
                <path d="M1 12.5h14V14H1v-1.5zm2-3.5h10V10.5H3V9zm0-3.5h7V7H3V5.5z" />
              </svg>
            </button>

            {/* Sub-divider */}
            <div className="h-3 w-[1px] bg-slate-200 mx-0.5" />

            {/* Align Left */}
            <button
              type="button"
              id="btn-align-left"
              onClick={() => onApplyFormat({ align: "left" })}
              className={`h-6 w-6 rounded flex items-center justify-center transition-colors cursor-pointer ${
                currentFormat.align === "left" || (!currentFormat.align && true)
                  ? "bg-blue-50 text-blue-700 border border-blue-300 shadow-2xs"
                  : "hover:bg-slate-100 text-slate-600 border border-transparent"
              }`}
              title="Align Left (Ctrl+L)"
            >
              <svg className="h-3 w-3" viewBox="0 0 16 16" fill="currentColor">
                <path d="M1 2.5h14V4H1V2.5zm0 3.5h9V7.5H1V6zm0 3.5h14V11H1V9.5zm0 3.5h9V14.5H1V13z" />
              </svg>
            </button>

            {/* Center */}
            <button
              type="button"
              id="btn-align-center"
              onClick={() => onApplyFormat({ align: "center" })}
              className={`h-6 w-6 rounded flex items-center justify-center transition-colors cursor-pointer ${
                currentFormat.align === "center"
                  ? "bg-blue-50 text-blue-700 border border-blue-300 shadow-2xs"
                  : "hover:bg-slate-100 text-slate-600 border border-transparent"
              }`}
              title="Center (Ctrl+E)"
            >
              <svg className="h-3 w-3" viewBox="0 0 16 16" fill="currentColor">
                <path d="M1 2.5h14V4H1V2.5zm2.5 3.5h9V7.5h-9V6zm-2.5 3.5h14V11H1V9.5zm2.5 3.5h9V14.5h-9V13z" />
              </svg>
            </button>

            {/* Align Right */}
            <button
              type="button"
              id="btn-align-right"
              onClick={() => onApplyFormat({ align: "right" })}
              className={`h-6 w-6 rounded flex items-center justify-center transition-colors cursor-pointer ${
                currentFormat.align === "right"
                  ? "bg-blue-50 text-blue-700 border border-blue-300 shadow-2xs"
                  : "hover:bg-slate-100 text-slate-600 border border-transparent"
              }`}
              title="Align Right (Ctrl+R)"
            >
              <svg className="h-3 w-3" viewBox="0 0 16 16" fill="currentColor">
                <path d="M1 2.5h14V4H1V2.5zm5 3.5h9V7.5H6V6zm-5 3.5h14V11H1V9.5zm5 3.5h9V14.5H6V13z" />
              </svg>
            </button>
          </div>

          {/* Divider */}
          <div className="h-3.5 w-[1px] bg-slate-200 mx-0.5" />

          {/* GROUP 4: MERGE & CLEAR FORMATS */}
          <div className="flex items-center gap-0.5">
            {onToggleMerge && (
              <button
                type="button"
                id="btn-merge-center"
                onClick={onToggleMerge}
                className="h-6 px-1.5 bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200 rounded flex items-center gap-1 transition-colors cursor-pointer text-[9.5px] font-semibold"
                title="Merge & Center selected cells"
              >
                <svg className="h-2.5 w-2.5" viewBox="0 0 16 16" fill="currentColor">
                  <path d="M2 3h12v10H2V3zm1 1v8h10V4H3zm2 3h6v2H5V7z" />
                </svg>
                <span>Merge</span>
              </button>
            )}

            {onClearFormatting && (
              <button
                type="button"
                id="btn-clear-format"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => {
                  applyInlineFormatting("clearFormat");
                  onClearFormatting();
                }}
                className="h-6 px-1.5 hover:bg-slate-100 text-slate-500 hover:text-slate-800 rounded flex items-center transition-colors cursor-pointer text-[9.5px]"
                title="Clear Formatting (Reset text & cell styles)"
              >
                <span>Clear</span>
              </button>
            )}
          </div>
        </div>

        {/* Selection Status Badge */}
        <div className="hidden lg:flex items-center gap-1.5 text-[10px] text-slate-500 font-mono bg-slate-50 px-2 py-1 rounded-md border border-slate-200">
          <span className="h-1.5 w-1.5 rounded-full bg-emerald-500" />
          <span>{selectionSummary || (hasSelection ? "Cell Selected" : "A4 Canvas")}</span>
        </div>
      </div>
    </div>
  );
}
