/**
 * Robust Clipboard & Excel / Google Sheets parser
 * Supports:
 * - HTML table extraction from Excel/Google Sheets rich clipboard
 * - TSV (Tab Separated Values) with multi-line quote escapes and inch-symbol resilience
 * - Intelligent header detection & column alignment
 * - Guarantees commas and semicolons inside cell content are NEVER split into separate cells
 * - Guarantees quotes and inch marks (e.g. 2", 1/2") never swallow subsequent lines or gather items into one cell
 * - Guarantees exact cell-to-cell layout preservation and blank cell filling
 */

import { stripHtml } from "./textFormatter";

export interface ParsedClipboardResult {
  grid: string[][];
  hasHeader: boolean;
  hasSerialColumn: boolean;
  columnMapping?: ("sl" | "desc" | "qty" | "unit" | "price" | "amount" | "ignore")[];
  detectedFormat: "html_table" | "tsv" | "csv" | "plain_lines";
}

/**
 * Helper to clean and flatten cell text into continuous space-separated text,
 * preserving commas, semicolons, quotes, and all punctuation intact.
 */
export function cleanCellText(str: string): string {
  if (!str) return "";
  const plain = str.includes("<") && str.includes(">") ? stripHtml(str) : str;
  let cleaned = plain
    .replace(/[\r\n]+/g, " ")
    .replace(/\u00A0/g, " ")
    .replace(/\s{2,}/g, " ")
    .trim();

  // If cell was wrapped in matching outer quotes from TSV/Excel escaping, unwrap them
  if (cleaned.startsWith('"') && cleaned.endsWith('"') && cleaned.length >= 2) {
    const inner = cleaned.slice(1, -1);
    // If it was wrapped by Excel because it contained quotes or newlines, unescape doubled quotes
    if (inner.includes('""') || !inner.includes('"')) {
      cleaned = inner.replace(/""/g, '"').trim();
    }
  }
  return cleaned;
}

/**
 * Extracts a 2D string array from an HTML table clipboard payload (Excel / Google Sheets).
 * Excel places rich HTML tables with <tr> and <td> tags on the clipboard.
 * Preserves blank cells, colspans, and exact cell positions.
 */
export function parseHTMLTable(html: string): string[][] | null {
  if (!html || !html.includes("<table") || typeof DOMParser === "undefined") {
    return null;
  }

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");
    const table = doc.querySelector("table");
    if (!table) return null;

    const rows = Array.from(table.querySelectorAll("tr"));
    if (rows.length === 0) return null;

    const grid: string[][] = [];

    rows.forEach((tr) => {
      const cells = Array.from(tr.querySelectorAll("th, td"));
      if (cells.length === 0) return;

      const rowValues: string[] = [];
      cells.forEach((cell) => {
        const colSpan = parseInt(cell.getAttribute("colspan") || "1", 10) || 1;
        const clones = cell.cloneNode(true) as HTMLElement;
        const brs = clones.querySelectorAll("br, p, div");
        brs.forEach((br) => br.replaceWith(" "));

        const text = clones.textContent || "";
        const cleaned = cleanCellText(text);
        rowValues.push(cleaned);

        // For spanned columns, pad with empty string so adjacent column indexes stay aligned
        for (let s = 1; s < colSpan; s++) {
          rowValues.push("");
        }
      });

      // Keep rows that have at least one non-empty cell
      const hasContent = rowValues.some((val) => val.length > 0);
      if (hasContent) {
        grid.push(rowValues);
      }
    });

    // Normalize all rows to have the same number of columns (fill with blank "")
    if (grid.length > 0) {
      const maxCols = Math.max(...grid.map((r) => r.length));
      grid.forEach((r) => {
        while (r.length < maxCols) {
          r.push("");
        }
      });
    }

    return grid.length > 0 ? grid : null;
  } catch (e) {
    console.warn("Failed to parse HTML table from clipboard:", e);
    return null;
  }
}

/**
 * State-machine parser for TSV (Tab-Separated Values).
 * Tabs unconditionally separate columns, newlines separate rows.
 * Commas and semicolons are NEVER treated as column delimiters!
 * Handles quotes without letting inch symbols (2", 1/2") swallow subsequent rows.
 */
export function parseTSV(text: string): string[][] {
  if (!text) return [];

  const result: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    // Check if cell starts with an opening quote (Excel wraps multiline cells in quotes)
    if (cell.length === 0 && char === '"') {
      inQuotes = true;
      continue;
    }

    if (inQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          // Escaped quote: "" -> "
          cell += '"';
          i++; // Skip next escaped quote
        } else if (nextChar === "\t" || nextChar === "\r" || nextChar === "\n" || i === text.length - 1) {
          // Closing quote at end of cell
          inQuotes = false;
        } else {
          // Quote within text (e.g. inch mark)
          cell += '"';
        }
      } else if (char === "\t") {
        // Tab is ALWAYS a column boundary in spreadsheet TSV
        row.push(cleanCellText(cell));
        cell = "";
        inQuotes = false;
      } else if (char === "\r" || char === "\n") {
        // Check if there is a closing quote ahead before the end of the text
        const remaining = text.slice(i);
        const hasClosingQuoteAhead =
          remaining.includes('"\t') ||
          remaining.includes('"\r') ||
          remaining.includes('"\n') ||
          remaining.endsWith('"');

        if (hasClosingQuoteAhead) {
          // Multiline cell (Alt+Enter in Excel) - convert to single space
          cell += " ";
          if (char === "\r" && nextChar === "\n") {
            i++; // Skip \n
          }
        } else {
          // Unclosed quote: terminate row safely instead of swallowing subsequent items!
          row.push(cleanCellText(cell));
          result.push(row);
          row = [];
          cell = "";
          inQuotes = false;
          if (char === "\r" && nextChar === "\n") {
            i++;
          }
        }
      } else {
        cell += char;
      }
    } else {
      if (char === "\t") {
        // Tab unconditionally ends the cell
        row.push(cleanCellText(cell));
        cell = "";
      } else if (char === "\r") {
        row.push(cleanCellText(cell));
        result.push(row);
        row = [];
        cell = "";
        if (nextChar === "\n") {
          i++; // Skip \n
        }
      } else if (char === "\n") {
        row.push(cleanCellText(cell));
        result.push(row);
        row = [];
        cell = "";
      } else {
        cell += char;
      }
    }
  }

  if (cell !== "" || row.length > 0) {
    row.push(cleanCellText(cell));
    result.push(row);
  }

  // Normalize grid: ensure all rows have equal columns so blank cells at the end of rows are preserved
  if (result.length > 0) {
    const maxCols = Math.max(...result.map((r) => r.length));
    result.forEach((r) => {
      while (r.length < maxCols) {
        r.push("");
      }
    });
  }

  // Remove completely empty trailing rows
  while (
    result.length > 0 &&
    result[result.length - 1].every((c) => c === "")
  ) {
    result.pop();
  }

  return result;
}

/**
 * Parses CSV text with delimiter (comma or semicolon).
 * Kept for optional file import if required.
 */
export function parseCSV(text: string, delimiter: "," | ";" = ","): string[][] {
  const result: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (inQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          cell += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else if (char === '\r' || char === '\n') {
        cell += ' ';
      } else {
        cell += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === delimiter) {
        row.push(cleanCellText(cell));
        cell = "";
      } else if (char === "\r") {
        if (nextChar === "\n") {
          row.push(cleanCellText(cell));
          result.push(row);
          row = [];
          cell = "";
          i++;
        } else {
          row.push(cleanCellText(cell));
          result.push(row);
          row = [];
          cell = "";
        }
      } else if (char === "\n") {
        row.push(cleanCellText(cell));
        result.push(row);
        row = [];
        cell = "";
      } else {
        cell += char;
      }
    }
  }

  if (cell !== "" || row.length > 0) {
    row.push(cleanCellText(cell));
    result.push(row);
  }

  while (
    result.length > 0 &&
    result[result.length - 1].every((c) => c === "")
  ) {
    result.pop();
  }

  return result;
}

export interface ParseClipboardOptions {
  allowCsv?: boolean;
}

/**
 * Universal clipboard parser that extracts a 2D grid from Excel, Google Sheets, or text.
 * Guarantees that commas and semicolons are NEVER treated as column splitters.
 * Guarantees that every cell in Excel maps cell-to-cell, and any blank places get filled.
 */
export function parseClipboardData(
  input: {
    text: string;
    html?: string;
  },
  _options?: ParseClipboardOptions
): ParsedClipboardResult {
  const { text, html } = input;
  const rawText = (text || "").trim();

  // 1. Prefer HTML Table parsing if available from Excel/Google Sheets rich clipboard.
  // HTML tables have explicit <tr> and <td> tags, ensuring 100% cell-to-cell fidelity
  // and preserving all blank cells without ambiguity.
  if (html && html.includes("<table")) {
    const htmlGrid = parseHTMLTable(html);
    if (htmlGrid && htmlGrid.length > 0 && htmlGrid.some((r) => r.length > 1 || htmlGrid.length > 1)) {
      return analyzeGrid(htmlGrid, "html_table");
    }
  }

  // 2. If clipboard text contains tabs, parse via native spreadsheet TSV.
  // Use text without outer trimming so leading/trailing tab separators are not lost!
  if ((text || "").includes("\t")) {
    const tsvGrid = parseTSV(text);
    if (tsvGrid.length > 0) {
      return analyzeGrid(tsvGrid, "tsv");
    }
  }

  if (!rawText) {
    return {
      grid: [],
      hasHeader: false,
      hasSerialColumn: false,
      detectedFormat: "plain_lines",
    };
  }

  // 3. Plain lines: each line is treated as ONE cell (single-column copy from Excel).
  // Sentences and descriptions frequently contain commas and semicolons;
  // they must NEVER be split into columns or broken into other cells!
  const lines = rawText.split(/\r?\n/).filter((l) => l.trim().length > 0);
  const lineGrid = lines.map((l) => [cleanCellText(l)]);
  return analyzeGrid(lineGrid, "plain_lines");
}

/**
 * Analyzes grid structure, detects headers and whether Column 0 is a Serial number
 */
function analyzeGrid(
  rawGrid: string[][],
  detectedFormat: "html_table" | "tsv" | "csv" | "plain_lines"
): ParsedClipboardResult {
  if (rawGrid.length === 0) {
    return {
      grid: [],
      hasHeader: false,
      hasSerialColumn: false,
      detectedFormat,
    };
  }

  const maxCols = Math.max(...rawGrid.map((r) => r.length));
  let hasHeader = false;

  const headerKeywords = [
    "sl", "s/n", "s.no", "sl.no", "#",
    "description", "particulars", "items", "details", "desc", "material", "specification",
    "qty", "quantity", "qnty",
    "unit", "uom", "pkg", "unit of measure",
    "price", "rate", "unit price", "unit rate",
    "amount", "total", "total amount", "subtotal"
  ];

  const isHeaderCell = (cell: string) => {
    const c = cell.toLowerCase().trim();
    if (!c || c.length > 30) return false;
    return headerKeywords.some((kw) => c === kw || c === kw + "." || c === kw + ":" || c === kw + " #" || c === "#");
  };

  // Only consider header if multiple columns exist, no numeric values in non-SL cells, and distinct columns match header words
  if (rawGrid.length > 1 && maxCols >= 2) {
    const firstRow = rawGrid[0];
    const hasNumericData = firstRow.some((cell, idx) => {
      if (idx === 0) return false;
      const clean = cell.replace(/[,$\s]/g, "").trim();
      return /^\d+(\.\d+)?$/.test(clean) && clean.length > 0;
    });

    if (!hasNumericData) {
      const headerMatches = firstRow.filter((cell) => isHeaderCell(cell));
      if (firstRow.length >= 4) {
        hasHeader = headerMatches.length >= 3;
      } else if (firstRow.length >= 2) {
        hasHeader = headerMatches.length >= 2;
      }
    }
  }

  // Check if column 0 represents a Serial Number (digits/indices) while column 1 represents Description
  const dataRows = hasHeader ? rawGrid.slice(1) : rawGrid;
  let hasSerialColumn = false;

  if (maxCols >= 2 && dataRows.length > 0) {
    const col0NonEmpty = dataRows.filter((r) => r[0] !== undefined && r[0].trim().length > 0);
    if (col0NonEmpty.length > 0) {
      const serialLikeCount = col0NonEmpty.filter((r) => {
        const val = r[0].replace(/[.\-#\s]/g, "").trim();
        return /^\d+$/.test(val) && val.length <= 5;
      }).length;

      const col1HasText = dataRows.some((r) => r.length >= 2 && /[a-zA-Z]/.test(r[1]));

      hasSerialColumn = (serialLikeCount / col0NonEmpty.length >= 0.6) && col1HasText;
    }
  }

  return {
    grid: rawGrid,
    hasHeader,
    hasSerialColumn,
    detectedFormat,
  };
}
