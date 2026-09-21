import ExcelJS from "exceljs";
import { CompanyProfile, QuotationRow } from "../types";
import { numberToWords } from "./numberToWords";

export interface ExcelGeneratorOptions {
  currentCompany: CompanyProfile;
  docType: "quotation" | "challan" | "invoice";
  messers: string;
  address: string;
  vesselName?: string;
  portBerth?: string;
  includeVesselName?: boolean;
  includePortBerth?: boolean;
  invoiceNo?: string;
  challanNo?: string;
  quotationNo?: string;
  requisitionNo?: string;
  poNumber?: string;
  includeInvoiceNo?: boolean;
  includeChallanNo?: boolean;
  includeQuotationNo?: boolean;
  includeRequisitionNo?: boolean;
  includePoNumber?: boolean;
  dateVal: string;
  rows: QuotationRow[];
  includeDiscount?: boolean;
  discountType?: "percentage" | "fixed";
  discountValue?: string | number;
  vatPercent?: string | number;
  transportationFee?: string | number;
  currency?: string;
  currencySymbol?: string;
}

/**
 * Converts standard CSS color (rgb, rgba, hex, named) into 8-character ARGB hex for ExcelJS.
 */
function cssColorToArgb(colorStr: string): string | null {
  if (!colorStr) return null;
  const s = colorStr.trim().toLowerCase();
  if (s === "transparent" || s === "inherit" || s === "initial" || s === "none") return null;

  // #RGB or #RRGGBB
  if (s.startsWith("#")) {
    const hex = s.slice(1);
    if (hex.length === 3) {
      const r = hex[0] + hex[0];
      const g = hex[1] + hex[1];
      const b = hex[2] + hex[2];
      return `FF${r}${g}${b}`.toUpperCase();
    }
    if (hex.length === 6) {
      return `FF${hex}`.toUpperCase();
    }
    if (hex.length === 8) {
      return hex.toUpperCase();
    }
  }

  // rgb(r, g, b) or rgba(r, g, b, a)
  const rgbMatch = s.match(/rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/);
  if (rgbMatch) {
    const r = Math.min(255, parseInt(rgbMatch[1], 10)).toString(16).padStart(2, "0");
    const g = Math.min(255, parseInt(rgbMatch[2], 10)).toString(16).padStart(2, "0");
    const b = Math.min(255, parseInt(rgbMatch[3], 10)).toString(16).padStart(2, "0");
    const a = rgbMatch[4] !== undefined
      ? Math.round(Math.min(1, parseFloat(rgbMatch[4])) * 255).toString(16).padStart(2, "0")
      : "FF";
    return `${a}${r}${g}${b}`.toUpperCase();
  }

  // Named CSS colors map
  const NAMED_COLORS: Record<string, string> = {
    black: "FF000000",
    white: "FFFFFFFF",
    red: "FFFF0000",
    green: "FF008000",
    blue: "FF0000FF",
    yellow: "FFFFFF00",
    cyan: "FF00FFFF",
    magenta: "FFFF00FF",
    gray: "FF808080",
    grey: "FF808080",
    orange: "FFFFA500",
    purple: "FF800080",
    gold: "FFFFD700",
    silver: "FFC0C0C0",
    navy: "FF000080",
    teal: "FF008080",
    maroon: "FF800000",
    olive: "FF808000",
    lime: "FF00FF00",
  };

  if (NAMED_COLORS[s]) {
    return NAMED_COLORS[s];
  }

  return null;
}

/**
 * Parses font size string (e.g., '14pt', '16px', '1.2em') into Excel point number.
 */
function parseFontSizeToPt(sizeStr: string): number | null {
  if (!sizeStr) return null;
  const s = sizeStr.trim().toLowerCase();
  const num = parseFloat(s);
  if (isNaN(num) || num <= 0) return null;

  if (s.endsWith("pt")) return Math.round(num * 10) / 10;
  if (s.endsWith("px")) return Math.round((num * 0.75) * 10) / 10;
  if (s.endsWith("em") || s.endsWith("rem")) return Math.round((num * 8) * 10) / 10;

  return Math.round(num * 10) / 10;
}

interface InlineStyleContext {
  color?: string; // ARGB
  bgColor?: string; // ARGB (highlight)
  size?: number; // pt
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | "double";
  strike?: boolean;
  fontFamily?: string;
}

/**
 * Parses an HTML string (from rich text editor) into an ExcelJS RichText array
 * and extracts any dominant background highlight color for cell filling.
 */
export function parseHtmlToRichText(
  html: string,
  defaultFontName = "Arial",
  defaultSize = 7.5
): { richText: ExcelJS.RichText[]; cellBgColor: string | null; plainText: string } {
  if (!html) {
    return { richText: [], cellBgColor: null, plainText: "" };
  }

  // Fast path if purely plain text (no tags or entities)
  if (!/<[a-z][\s\S]*>/i.test(html) && !/&[a-z0-9#]+;/i.test(html)) {
    const text = html.replace(/\r\n/g, "\n");
    return {
      richText: [
        {
          text,
          font: { name: defaultFontName, size: defaultSize, color: { argb: "FF000000" } },
        },
      ],
      cellBgColor: null,
      plainText: text,
    };
  }

  // In browser environments, use DOMParser for accurate HTML tree traversal
  if (typeof DOMParser !== "undefined") {
    try {
      const parser = new DOMParser();
      // Parse clean HTML directly without preemptive regex that creates fake trailing newlines
      const doc = parser.parseFromString(`<body>${html}</body>`, "text/html");
      const richText: ExcelJS.RichText[] = [];
      let dominantCellBgColor: string | null = null;
      let totalPlainText = "";

      const BLOCK_TAGS = new Set(["p", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6"]);

      function traverse(node: Node, ctx: InlineStyleContext) {
        if (node.nodeType === Node.TEXT_NODE) {
          const rawText = node.textContent || "";
          if (!rawText) return;

          totalPlainText += rawText;

          // Build ExcelJS Font object
          const font: Partial<ExcelJS.Font> = {
            name: ctx.fontFamily || defaultFontName,
            size: ctx.size || defaultSize,
            bold: !!ctx.bold,
            italic: !!ctx.italic,
            underline: ctx.underline === "double" ? "double" : !!ctx.underline,
            strike: !!ctx.strike,
            color: ctx.color ? { argb: ctx.color } : { argb: "FF000000" },
          };

          richText.push({ text: rawText, font });

          if (ctx.bgColor && !dominantCellBgColor) {
            dominantCellBgColor = ctx.bgColor;
          }
          return;
        }

        if (node.nodeType === Node.ELEMENT_NODE) {
          const el = node as HTMLElement;
          const tagName = el.tagName.toLowerCase();

          // If entering a block element and we already have text from a previous block, insert a newline separator
          if (BLOCK_TAGS.has(tagName) && richText.length > 0) {
            const lastItem = richText[richText.length - 1];
            if (lastItem && !lastItem.text.endsWith("\n")) {
              richText.push({
                text: "\n",
                font: { name: defaultFontName, size: defaultSize, color: { argb: "FF000000" } },
              });
              totalPlainText += "\n";
            }
          }

          // Line break tag (<br>)
          if (tagName === "br") {
            // Ignore browser contenteditable trailing placeholder <br> (when it is the last child of a block)
            const isTrailingDummyBr = !el.nextSibling && !!el.parentElement && BLOCK_TAGS.has(el.parentElement.tagName.toLowerCase());
            if (!isTrailingDummyBr) {
              totalPlainText += "\n";
              richText.push({
                text: "\n",
                font: {
                  name: ctx.fontFamily || defaultFontName,
                  size: ctx.size || defaultSize,
                  color: ctx.color ? { argb: ctx.color } : { argb: "FF000000" },
                },
              });
            }
            return;
          }

          // Inherit previous context
          const nextCtx: InlineStyleContext = { ...ctx };

          // HTML Tag-based formatting
          if (tagName === "b" || tagName === "strong") nextCtx.bold = true;
          if (tagName === "i" || tagName === "em") nextCtx.italic = true;
          if (tagName === "u" || tagName === "ins") nextCtx.underline = true;
          if (tagName === "s" || tagName === "strike" || tagName === "del") nextCtx.strike = true;

          // Inline style attributes
          const style = el.style;
          if (style) {
            if (style.color) {
              const argb = cssColorToArgb(style.color);
              if (argb) nextCtx.color = argb;
            }

            // Highlighting / Background color
            const bg = style.backgroundColor || style.background;
            if (bg) {
              const bgArgb = cssColorToArgb(bg);
              if (bgArgb) nextCtx.bgColor = bgArgb;
            }

            // Font size
            if (style.fontSize) {
              const pt = parseFontSizeToPt(style.fontSize);
              if (pt) nextCtx.size = pt;
            }

            // Font weight
            if (style.fontWeight) {
              const fw = style.fontWeight.toLowerCase();
              if (fw === "bold" || fw === "700" || fw === "800" || fw === "900") {
                nextCtx.bold = true;
              } else if (fw === "normal" || fw === "400") {
                nextCtx.bold = false;
              }
            }

            // Font style
            if (style.fontStyle) {
              if (style.fontStyle.toLowerCase() === "italic") nextCtx.italic = true;
              else if (style.fontStyle.toLowerCase() === "normal") nextCtx.italic = false;
            }

            // Text decoration
            if (style.textDecoration || (style as any).textDecorationLine) {
              const deco = (style.textDecoration || (style as any).textDecorationLine).toLowerCase();
              if (deco.includes("underline")) {
                nextCtx.underline = (style as any).textDecorationStyle === "double" ? "double" : true;
              }
              if (deco.includes("line-through")) {
                nextCtx.strike = true;
              }
              if (deco.includes("none")) {
                nextCtx.underline = false;
                nextCtx.strike = false;
              }
            }

            // Font family
            if (style.fontFamily) {
              const cleanFamily = style.fontFamily.split(",")[0].replace(/['"]/g, "").trim();
              if (cleanFamily) nextCtx.fontFamily = cleanFamily;
            }
          }

          // HTML font tag attributes (legacy)
          if (tagName === "font") {
            const fontColor = el.getAttribute("color");
            if (fontColor) {
              const argb = cssColorToArgb(fontColor);
              if (argb) nextCtx.color = argb;
            }
            const fontFace = el.getAttribute("face");
            if (fontFace) nextCtx.fontFamily = fontFace;
          }

          // HTML mark tag (highlight)
          if (tagName === "mark") {
            nextCtx.bgColor = "FFFFFF00"; // default yellow highlight
          }

          Array.from(node.childNodes).forEach((child) => traverse(child, nextCtx));
        }
      }

      Array.from(doc.body.childNodes).forEach((child) => traverse(child, {}));

      // Sanitize richText: filter empty elements and strip leading/trailing blank lines
      let cleanedRichText = richText.filter((rt) => rt.text && rt.text.length > 0);

      // Strip leading whitespace and newlines from the start of richText
      while (cleanedRichText.length > 0 && /^[\r\n\s]/.test(cleanedRichText[0].text)) {
        cleanedRichText[0].text = cleanedRichText[0].text.replace(/^[\r\n\s]+/, "");
        if (!cleanedRichText[0].text) {
          cleanedRichText.shift();
        } else {
          break;
        }
      }

      // Strip trailing whitespace and newlines from the end of richText
      while (cleanedRichText.length > 0 && /[\r\n\s]$/.test(cleanedRichText[cleanedRichText.length - 1].text)) {
        const lastIdx = cleanedRichText.length - 1;
        cleanedRichText[lastIdx].text = cleanedRichText[lastIdx].text.replace(/[\r\n\s]+$/, "");
        if (!cleanedRichText[lastIdx].text) {
          cleanedRichText.pop();
        } else {
          break;
        }
      }

      // Collapse consecutive newlines inside elements so line spacing is tight and uniform
      cleanedRichText.forEach((rt) => {
        rt.text = rt.text.replace(/\r?\n\s*\r?\n/g, "\n").replace(/\n{2,}/g, "\n");
        // Ensure any newline character has a compact font size (max 7.5pt) so Excel does not stretch line spacing
        if (rt.text === "\n" && rt.font && rt.font.size && rt.font.size > 7.5) {
          rt.font.size = 7.5;
        }
      });

      // Also collapse newlines across adjacent elements
      for (let i = 0; i < cleanedRichText.length - 1; i++) {
        if (cleanedRichText[i].text.endsWith("\n") && cleanedRichText[i + 1].text.startsWith("\n")) {
          cleanedRichText[i + 1].text = cleanedRichText[i + 1].text.replace(/^\n+/, "");
        }
      }
      cleanedRichText = cleanedRichText.filter((rt) => rt.text.length > 0);

      // If richText is empty, provide fallback
      if (cleanedRichText.length === 0) {
        const clean = cleanHtmlText(html);
        return {
          richText: [{ text: clean, font: { name: defaultFontName, size: defaultSize, color: { argb: "FF000000" } } }],
          cellBgColor: null,
          plainText: clean,
        };
      }

      const cleanText = cleanedRichText.map((rt) => rt.text).join("");

      return {
        richText: cleanedRichText,
        cellBgColor: dominantCellBgColor,
        plainText: cleanText,
      };
    } catch (e) {
      console.warn("DOMParser rich text parse error, falling back:", e);
    }
  }

  // Regex fallback
  const clean = cleanHtmlText(html);
  return {
    richText: [{ text: clean, font: { name: defaultFontName, size: defaultSize, color: { argb: "FF000000" } } }],
    cellBgColor: null,
    plainText: clean,
  };
}

/**
 * Strips HTML tags, decodes standard HTML entities, and normalizes line breaks
 * so multi-line text wraps cleanly in Excel cells.
 */
export function cleanHtmlText(html: string): string {
  if (!html) return "";
  let text = html
    .replace(/<br\s*[\/]?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<li>/gi, "• ")
    .replace(/<\/li>/gi, "\n")
    .replace(/<[^>]+>/g, "");

  text = text
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");

  // Collapse duplicate line breaks so text lines sit together without blank gaps
  text = text.replace(/\r?\n\s*\r?\n/g, "\n").replace(/\n{2,}/g, "\n");

  return text.trim();
}

/**
 * Accurately estimates rendered lines of text inside an Excel cell considering:
 * 1) Explicit newlines (\n, \r\n) from user pressing enter (collapsed to prevent blank gaps)
 * 2) Natural word-wrapping ONLY when text exceeds the column width
 */
export function estimateTextLines(text: string, colCharCap: number): number {
  if (!text) return 1;
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n{2,}/g, "\n").trim();
  if (!normalized) return 1;

  const paragraphs = normalized.split("\n");
  let totalLines = 0;

  for (const para of paragraphs) {
    const trimmed = para.trim();
    if (!trimmed) continue;

    // If paragraph fits completely within the column capacity, it is strictly 1 line!
    if (trimmed.length <= colCharCap) {
      totalLines += 1;
      continue;
    }

    // Paragraph exceeds column width: simulate natural word wrapping
    const words = trimmed.split(/\s+/);
    let currentLineLen = 0;
    let paraLines = 1;

    for (const word of words) {
      const wordLen = word.length;
      if (wordLen === 0) continue;

      if (currentLineLen === 0) {
        if (wordLen > colCharCap) {
          paraLines += Math.floor((wordLen - 1) / colCharCap);
          currentLineLen = wordLen % colCharCap || colCharCap;
        } else {
          currentLineLen = wordLen;
        }
      } else if (currentLineLen + 1 + wordLen <= colCharCap) {
        currentLineLen += 1 + wordLen;
      } else {
        paraLines += 1;
        if (wordLen > colCharCap) {
          paraLines += Math.floor((wordLen - 1) / colCharCap);
          currentLineLen = wordLen % colCharCap || colCharCap;
        } else {
          currentLineLen = wordLen;
        }
      }
    }
    totalLines += paraLines;
  }

  return Math.max(1, totalLines);
}

/**
 * Accurately determines the line-wrapped structure and row height for an Excel cell containing
 * rich text (or plain text), accounting for:
 * 1. Column character capacity at base font size (Arial 7.5pt):
 *    - Challan description (col 63): ~64 chars
 *    - Quotation description (col 53): ~56 chars
 *    - Remarks (col 18): ~19 chars
 * 2. Mixed font sizes (e.g. 7.5pt body with 12pt/14pt highlight phrase):
 *    - Wider font sizes take proportionally more horizontal space per character
 *    - Line height is calculated based ONLY on the maximum font size on THAT specific line,
 *      NEVER multiplying an entire cell's height by a single large word's font size!
 * 3. Explicit newlines (\n)
 * 4. Snug inter-line spacing and clean vertical cell padding
 */
export function calculateAccurateCellLinesAndHeight(
  richText: ExcelJS.RichText[] | null | undefined,
  plainTextFallback: string,
  colCharCap: number,
  defaultFontSize = 7.5
): { lines: number; height: number } {
  const tokens: { text: string; size: number; isSpace: boolean; isNewline: boolean }[] = [];

  if (richText && richText.length > 0) {
    for (const rt of richText) {
      const txt = rt.text || "";
      const size = rt.font?.size || defaultFontSize;
      const subLines = txt.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
      for (let i = 0; i < subLines.length; i++) {
        if (i > 0) {
          tokens.push({ text: "\n", size, isSpace: false, isNewline: true });
        }
        const line = subLines[i];
        if (!line) continue;
        const matches = line.match(/\S+|\s+/g) || [];
        for (const m of matches) {
          tokens.push({ text: m, size, isSpace: /^\s+$/.test(m), isNewline: false });
        }
      }
    }
  } else {
    const txt = plainTextFallback || "";
    const subLines = txt.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    for (let i = 0; i < subLines.length; i++) {
      if (i > 0) {
        tokens.push({ text: "\n", size: defaultFontSize, isSpace: false, isNewline: true });
      }
      const line = subLines[i];
      if (!line) continue;
      const matches = line.match(/\S+|\s+/g) || [];
      for (const m of matches) {
        tokens.push({ text: m, size: defaultFontSize, isSpace: /^\s+$/.test(m), isNewline: false });
      }
    }
  }

  if (tokens.length === 0) {
    return { lines: 1, height: 12 };
  }

  const lines: { text: string; size: number }[][] = [];
  let currentLine: { text: string; size: number }[] = [];
  let currentUnits = 0;

  for (const tok of tokens) {
    if (tok.isNewline) {
      lines.push(currentLine);
      currentLine = [];
      currentUnits = 0;
      continue;
    }

    const charUnit = tok.size / defaultFontSize;
    const tokUnits = tok.text.length * charUnit;

    if (currentLine.length === 0) {
      if (tok.isSpace) continue; // skip leading space
      if (tokUnits > colCharCap) {
        let remText = tok.text;
        while (remText.length > 0) {
          const fitChars = Math.max(1, Math.floor(colCharCap / charUnit));
          const chunk = remText.slice(0, fitChars);
          remText = remText.slice(fitChars);
          lines.push([{ text: chunk, size: tok.size }]);
        }
        currentLine = [];
        currentUnits = 0;
      } else {
        currentLine.push({ text: tok.text, size: tok.size });
        currentUnits = tokUnits;
      }
    } else if (currentUnits + tokUnits <= colCharCap) {
      currentLine.push({ text: tok.text, size: tok.size });
      currentUnits += tokUnits;
    } else {
      lines.push(currentLine);
      if (tok.isSpace) {
        currentLine = [];
        currentUnits = 0;
      } else if (tokUnits > colCharCap) {
        let remText = tok.text;
        while (remText.length > 0) {
          const fitChars = Math.max(1, Math.floor(colCharCap / charUnit));
          const chunk = remText.slice(0, fitChars);
          remText = remText.slice(fitChars);
          lines.push([{ text: chunk, size: tok.size }]);
        }
        currentLine = [];
        currentUnits = 0;
      } else {
        currentLine = [{ text: tok.text, size: tok.size }];
        currentUnits = tokUnits;
      }
    }
  }

  if (currentLine.length > 0) {
    lines.push(currentLine);
  }

  const validLines = lines.filter((l) => l.length > 0);
  const lineCount = Math.max(1, validLines.length);

  let totalHeight = 2.0; // Top cell padding (pt)
  for (let i = 0; i < validLines.length; i++) {
    const lineTokens = validLines[i];
    let lineMaxSize = defaultFontSize;
    for (const t of lineTokens) {
      if (t.size > lineMaxSize) {
        lineMaxSize = t.size;
      }
    }
    const linePitch = lineMaxSize <= 8.5 ? 9.2 : Math.max(9.2, lineMaxSize * 1.12);
    totalHeight += linePitch;
  }
  totalHeight += 1.5; // Bottom cell padding (pt)

  if (lineCount === 1) {
    let singleMaxSize = defaultFontSize;
    for (const t of validLines[0] || []) {
      if (t.size > singleMaxSize) singleMaxSize = t.size;
    }
    const singleHeight = singleMaxSize <= 8.5 ? 12 : Math.max(12, singleMaxSize * 1.35);
    return { lines: 1, height: singleHeight };
  }

  return { lines: lineCount, height: Math.round(totalHeight * 10) / 10 };
}

/**
 * Calculates compact, line-accurate row height (in pt) for Excel:
 * - 1 line: 12 pt (clean, comfortable, perfectly readable Arial 7.5pt with ~2.25pt padding top/bottom)
 * - Additional lines: +9 pt per extra line (minimal, snug inter-line space without clipping)
 * - Proportional scaling if custom enlarged font size is applied
 */
export function calculateCompactRowHeight(lines: number, maxFontSize = 7.5): number {
  const count = Math.max(1, lines);
  const baseLineRate = maxFontSize > 7.5 ? Math.max(9, maxFontSize * 1.15) : 9;
  const firstLineHeight = maxFontSize > 7.5 ? Math.max(12, maxFontSize * 1.45) : 12;
  return count === 1 ? firstLineHeight : firstLineHeight + (count - 1) * baseLineRate;
}

/**
 * Safely parses a numeric input string or number
 */
function parseNum(val: string | number | undefined): number {
  if (val === undefined || val === null) return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const cleaned = val.toString().replace(/,/g, "").trim();
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
}

/**
 * Fetches an image from URL and converts to ArrayBuffer for embedding in ExcelJS
 */
async function fetchImageBuffer(url: string): Promise<ArrayBuffer | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    return await res.arrayBuffer();
  } catch (err) {
    console.warn("Could not fetch image for Excel embedding:", err);
    return null;
  }
}

const BORDER_COLOR = "FF475569"; // Slate-600 crisp dark border matching web and print styles

const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: "thin", color: { argb: BORDER_COLOR } },
  bottom: { style: "thin", color: { argb: BORDER_COLOR } },
  left: { style: "thin", color: { argb: BORDER_COLOR } },
  right: { style: "thin", color: { argb: BORDER_COLOR } },
};

interface PreparedItem {
  row: QuotationRow;
  originalIndex: number;
  cleanDesc: string;
  richDesc: ExcelJS.RichText[];
  cellBgColor: string | null;
  height: number;
  lines: number;
}

interface PageData {
  pageNumber: number;
  items: PreparedItem[];
  isFirstPage: boolean;
  isLastPage: boolean;
}

/**
 * Generates an Excel document (.xlsx) with:
 * - Row 1: Document format title with NO BORDERS and clean transparent background
 * - Rows 2 to 12: 11 blank rows for pre-printed company letterhead stationery
 * - Utilizes the FULL printable height of A4 paper (~800pt) exclusively to fit maximum items per page
 * - Preserves all user formatting: text colors, font size changes, highlights/background fills, bold, italic, underline
 * - Stamp image placed ABOVE the Authorized Signature text line for Comilla Traders (matching PDF)
 * - Signature section present on EVERY single page according to the format
 * - Dynamic pagination dividing items when the A4 printable area is full across worksheets
 */
export async function generateExcelDocument(options: ExcelGeneratorOptions): Promise<void> {
  const {
    currentCompany,
    docType,
    messers,
    address,
    vesselName = "",
    portBerth = "",
    includeVesselName = false,
    includePortBerth = false,
    invoiceNo = "",
    challanNo = "",
    quotationNo = "",
    requisitionNo = "",
    poNumber = "",
    includeInvoiceNo = true,
    includeChallanNo = true,
    includeQuotationNo = true,
    includeRequisitionNo = true,
    includePoNumber = true,
    dateVal,
    rows,
    includeDiscount = false,
    discountType = "percentage",
    discountValue = "0",
    vatPercent = "0",
    transportationFee = "0",
    currency = "BDT",
    currencySymbol = "Tk",
  } = options;

  const isChallan = docType === "challan";
  const isInvoice = docType === "invoice";
  const colCount = isChallan ? 4 : 6;
  const lastColLetter = isChallan ? "D" : "F";
  const docTitleText = isChallan ? "DELIVERY CHALLAN" : isInvoice ? "INVOICE" : "QUOTATION";
  const baseSheetName = isChallan ? "Challan" : isInvoice ? "Invoice" : "Quotation";

  // Filter active rows or fallback
  const activeRows = rows.filter((r) => {
    const hasDesc = (r.desc || "").trim() !== "";
    const hasQty = (r.qty || "").trim() !== "";
    const hasPrice = (r.price || "").trim() !== "";
    const hasAmount = (r.amount || 0) > 0;
    return hasDesc || hasQty || hasPrice || hasAmount;
  });

  const rowsToExport = activeRows.length > 0 ? activeRows : rows.slice(0, 10);

  // Accurately compute lines and row height using calculateAccurateCellLinesAndHeight
  // so that mixed font sizes (e.g. standard 7.5pt with highlighted 14pt phrase) and natural
  // column wrap capacities are calculated with zero excess vertical whitespace.
  const preparedItems: PreparedItem[] = rowsToExport.map((r, idx) => {
    // Parse HTML to rich text with colors, font sizes, highlighting, bold, italic, underline
    const { richText, cellBgColor, plainText } = parseHtmlToRichText(r.desc || "", "Arial", 7.5);
    const cleanDesc = (plainText || cleanHtmlText(r.desc)).trim();
    
    // In Arial 7.5pt:
    // - Challan column width 63 fits ~64 characters per line
    // - Quotation/Invoice column width 53 fits ~56 characters per line
    const descCharCap = isChallan ? 64 : 56;
    const descResult = calculateAccurateCellLinesAndHeight(richText, cleanDesc, descCharCap, 7.5);

    let remarksResult = { lines: 1, height: 12 };
    if (isChallan && r.unit) {
      const { richText: remRich, plainText: remPlain } = parseHtmlToRichText(r.unit || "", "Arial", 7.5);
      remarksResult = calculateAccurateCellLinesAndHeight(remRich, remPlain || cleanHtmlText(r.unit), 19, 7.5);
    }

    const lines = Math.max(1, descResult.lines, remarksResult.lines);
    const height = Math.max(12, descResult.height, remarksResult.height);

    return {
      row: r,
      originalIndex: idx + 1,
      cleanDesc,
      richDesc: richText,
      cellBgColor,
      height,
      lines,
    };
  });

  // Calculate summary counts for Quotation & Invoice
  const rowsTotal = rowsToExport.reduce((sum, r) => sum + (r.amount || 0), 0);
  const parsedVat = parseNum(vatPercent);
  const parsedTransport = parseNum(transportationFee);
  const parsedDiscVal = parseNum(discountValue);

  let discountAmount = 0;
  if (isInvoice && includeDiscount && parsedDiscVal > 0) {
    if (discountType === "percentage") {
      discountAmount = (rowsTotal * parsedDiscVal) / 100;
    } else {
      discountAmount = parsedDiscVal;
    }
  }
  discountAmount = Math.min(rowsTotal, Math.max(0, discountAmount));
  const netAfterDiscount = Math.max(0, rowsTotal - discountAmount);
  const vatAmount = isInvoice ? (netAfterDiscount * parsedVat) / 100 : 0;
  const grandTotal = isInvoice ? netAfterDiscount + vatAmount + parsedTransport : rowsTotal;

  // Build Metadata Left and Right lines for Page 1 (only include non-empty fields)
  const cleanMessers = cleanHtmlText(messers) || " ";
  const cleanAddress = cleanHtmlText(address) || " ";

  const leftLines: { label: string; value: string; bold?: boolean }[] = [];
  if (cleanMessers.trim()) {
    leftLines.push({ label: "Messers:", value: cleanMessers, bold: true });
  }
  if (vesselName && vesselName.trim()) {
    leftLines.push({ label: "Vessel Name:", value: vesselName.trim(), bold: true });
  }
  if (portBerth && portBerth.trim()) {
    leftLines.push({ label: "Port / Berth:", value: portBerth.trim() });
  }
  if (cleanAddress.trim()) {
    leftLines.push({ label: "Address:", value: cleanAddress });
  }

  const rightLines: { label: string; value: string; bold?: boolean }[] = [];
  if (isInvoice) {
    if (invoiceNo && invoiceNo.trim()) rightLines.push({ label: "Invoice No.:", value: invoiceNo.trim(), bold: true });
    if (challanNo && challanNo.trim()) rightLines.push({ label: "Challan No.:", value: challanNo.trim(), bold: true });
    rightLines.push({ label: "Date:", value: dateVal || new Date().toLocaleDateString("en-GB"), bold: true });
    if (requisitionNo && requisitionNo.trim()) rightLines.push({ label: "Requisition No.:", value: requisitionNo.trim(), bold: true });
    if (poNumber && poNumber.trim()) rightLines.push({ label: "PO Number:", value: poNumber.trim(), bold: true });
  } else if (isChallan) {
    if (challanNo && challanNo.trim()) rightLines.push({ label: "Challan No.:", value: challanNo.trim(), bold: true });
    rightLines.push({ label: "Date:", value: dateVal || new Date().toLocaleDateString("en-GB"), bold: true });
    if (requisitionNo && requisitionNo.trim()) rightLines.push({ label: "Requisition No.:", value: requisitionNo.trim(), bold: true });
    if (poNumber && poNumber.trim()) rightLines.push({ label: "PO Number:", value: poNumber.trim(), bold: true });
  } else {
    // Quotation format
    if (quotationNo && quotationNo.trim()) rightLines.push({ label: "Quotation No.:", value: quotationNo.trim(), bold: true });
    if (requisitionNo && requisitionNo.trim()) rightLines.push({ label: "Requisition No.:", value: requisitionNo.trim(), bold: true });
    rightLines.push({ label: "Date:", value: dateVal || new Date().toLocaleDateString("en-GB"), bold: true });
  }

  const maxMetaRows = Math.max(leftLines.length, rightLines.length, 1);
  const leftColEnd = isChallan ? 2 : 3;
  const rightColStart = leftColEnd + 1;

  // Number of rows in summary block
  let summaryRowsCount = 2; // Subtotal + Grand Total
  if (isInvoice && includeDiscount && discountAmount > 0) summaryRowsCount++;
  if (isInvoice && vatAmount > 0) summaryRowsCount++;
  if (isInvoice && parsedTransport > 0) summaryRowsCount++;

  let calculatedMetaBoxHeight = 0;
  for (let i = 0; i < maxMetaRows; i++) {
    const lText = leftLines[i] ? `${leftLines[i].label} ${leftLines[i].value}` : "";
    const rText = rightLines[i] ? `${rightLines[i].label} ${rightLines[i].value}` : "";
    const lLines = estimateTextLines(lText, isChallan ? 58 : 56);
    const rLines = estimateTextLines(rText, isChallan ? 26 : 32);
    const metaLines = Math.max(1, lLines, rLines);
    calculatedMetaBoxHeight += calculateCompactRowHeight(metaLines);
  }

  // FULL PAGE USAGE IN EXCEL (A4 height with 0.2in margins = ~800pt printable)
  const PAGE_LIMIT = 800;
  const ROW1_TITLE_HEIGHT = 16;
  const BLANK_ROWS_HEIGHT = 10 * 15; // 150pt (10 blank rows: rows 2 to 11 at 15pt normal Excel row height)
  const META_BOX_HEIGHT = calculatedMetaBoxHeight + 3; // ~40-52pt exact metadata box height + gap
  const TABLE_HEADER_HEIGHT = 15;

  const P1_PRE_HEIGHT = ROW1_TITLE_HEIGHT + BLANK_ROWS_HEIGHT + META_BOX_HEIGHT + TABLE_HEADER_HEIGHT;
  const CONT_PRE_HEIGHT = ROW1_TITLE_HEIGHT + BLANK_ROWS_HEIGHT + META_BOX_HEIGHT + TABLE_HEADER_HEIGHT;

  // Signature block is present on EVERY page according to format (16pt gap + 78.5pt block)
  const SIGNATURE_BLOCK_HEIGHT = 94.5;

  // Summary block (Subtotal, VAT, Grand Total) is placed on the final page
  const SUMMARY_BLOCK_HEIGHT = isChallan ? 0 : (summaryRowsCount * 13 + 5);

  const INTERMEDIATE_FOOTER_HEIGHT = SIGNATURE_BLOCK_HEIGHT;
  const FINAL_FOOTER_HEIGHT = SUMMARY_BLOCK_HEIGHT + SIGNATURE_BLOCK_HEIGHT;

  // =========================================================================
  // PAGINATION / FULL PAGE USAGE ALGORITHM
  // =========================================================================
  const pages: PageData[] = [];
  const totalItemsHeight = preparedItems.reduce((s, it) => s + it.height, 0);

  if (P1_PRE_HEIGHT + totalItemsHeight + FINAL_FOOTER_HEIGHT <= PAGE_LIMIT) {
    pages.push({
      pageNumber: 1,
      items: preparedItems,
      isFirstPage: true,
      isLastPage: true,
    });
  } else {
    let remaining = [...preparedItems];
    let pageNum = 1;

    while (remaining.length > 0) {
      const isFirst = pageNum === 1;
      const preHeight = isFirst ? P1_PRE_HEIGHT : CONT_PRE_HEIGHT;
      const maxAvailableFinal = PAGE_LIMIT - preHeight - FINAL_FOOTER_HEIGHT;
      const maxAvailableIntermediate = PAGE_LIMIT - preHeight - INTERMEDIATE_FOOTER_HEIGHT;

      // Check if ALL remaining items plus the final footer fit on this page
      const remainingHeight = remaining.reduce((s, it) => s + it.height, 0);
      if (remainingHeight <= maxAvailableFinal) {
        pages.push({
          pageNumber: pageNum,
          items: remaining,
          isFirstPage: isFirst,
          isLastPage: true,
        });
        remaining = [];
        break;
      }

      // Maximize items placed onto this page (using intermediate footer budget)
      const pageItems: PreparedItem[] = [];
      let currentHeight = 0;

      while (remaining.length > 0) {
        const nextItem = remaining[0];
        if (pageItems.length > 0 && currentHeight + nextItem.height > maxAvailableIntermediate) {
          break;
        }
        currentHeight += nextItem.height;
        pageItems.push(remaining.shift()!);
      }

      // If all items fit on this page but the final footer cannot, shift minimal items to next page
      if (remaining.length === 0 && pageItems.length > 2) {
        const moved = pageItems.splice(pageItems.length - 2, 2);
        remaining = moved;
      }

      pages.push({
        pageNumber: pageNum,
        items: pageItems,
        isFirstPage: isFirst,
        isLastPage: false,
      });

      pageNum++;
    }

    if (remaining.length > 0) {
      pages.push({
        pageNumber: pageNum,
        items: remaining,
        isFirstPage: false,
        isLastPage: true,
      });
    }
  }

  // =========================================================================
  // PRE-FETCH COMPANY STAMP IMAGE (IF CONFIGURED)
  // =========================================================================
  let stampImageId: number | null = null;
  const hasStamp = currentCompany.hasStamp && Boolean(currentCompany.stampUrl);
  let stampBuffer: ArrayBuffer | null = null;

  if (hasStamp && currentCompany.stampUrl) {
    stampBuffer = await fetchImageBuffer(currentCompany.stampUrl);
  }

  // =========================================================================
  // WORKBOOK CREATION & PER-PAGE SHEET GENERATION
  // =========================================================================
  const workbook = new ExcelJS.Workbook();
  workbook.creator = currentCompany.name;
  workbook.lastModifiedBy = currentCompany.name;
  workbook.created = new Date();
  workbook.modified = new Date();

  // Register stamp image with workbook if available
  if (stampBuffer) {
    try {
      stampImageId = workbook.addImage({
        buffer: stampBuffer,
        extension: "png",
      });
    } catch (e) {
      console.warn("Could not register stamp in Excel workbook:", e);
    }
  }

  pages.forEach((page) => {
    const sheetName = pages.length === 1 
      ? baseSheetName 
      : `${baseSheetName} - Page ${page.pageNumber}`.substring(0, 31);

    const ws = workbook.addWorksheet(sheetName, {
      pageSetup: {
        paperSize: 9, // A4
        orientation: "portrait",
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1, // Fit exactly to 1 page tall
        margins: {
          left: 0.2,
          right: 0.2,
          top: 0.2,
          bottom: 0.2,
          header: 0.05,
          footer: 0.05,
        },
      },
      views: [{ showGridLines: true }],
    });

    // Optimized column widths for maximum content visibility & print fitting
    if (isChallan) {
      ws.columns = [
        { key: "sl", width: 5.5 },
        { key: "desc", width: 63 },
        { key: "qty", width: 12 },
        { key: "remarks", width: 18 },
      ];
    } else {
      ws.columns = [
        { key: "sl", width: 5 },
        { key: "desc", width: 53 },
        { key: "qty", width: 8 },
        { key: "unit", width: 8 },
        { key: "rate", width: 14 },
        { key: "total", width: 16.5 },
      ];
    }

    // =========================================================================
    // ROW 1: FORMAT NAME ON THE FIRST ROW (NO BORDERS)
    // =========================================================================
    const pageFormatTitle = pages.length > 1
      ? `${docTitleText}  —  PAGE ${page.pageNumber} OF ${pages.length}`
      : docTitleText;

    const row1 = ws.addRow([pageFormatTitle]);
    row1.height = ROW1_TITLE_HEIGHT;
    ws.mergeCells(`A1:${lastColLetter}1`);
    const row1Cell = ws.getCell("A1");
    row1Cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF0F172A" } };
    row1Cell.alignment = { horizontal: "center", vertical: "middle" };
    row1Cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
    // Remove borders from format name
    row1Cell.border = {};
    row1.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = {};
    });

    // =========================================================================
    // ROWS 2 TO 11: BLANK ROWS (For pre-printed letterhead stationery, normal Excel size 15pt)
    // =========================================================================
    for (let r = 2; r <= 11; r++) {
      const blankRow = ws.addRow([]);
      blankRow.height = 15; // Normal Excel default row height (15pt)
    }

    let currentRow = 12;

    // =========================================================================
    // METADATA BOXES (Unified Bordered Container with Vertical Middle Divider)
    // =========================================================================
    const metaStartRow = currentRow;
    const metaEndRow = currentRow + maxMetaRows - 1;

    for (let i = 0; i < maxMetaRows; i++) {
      const lItem = leftLines[i];
      const rItem = rightLines[i];

      const lText = lItem ? `${lItem.label} ${lItem.value}` : "";
      const rText = rItem ? `${rItem.label} ${rItem.value}` : "";
      const lLines = estimateTextLines(lText, isChallan ? 58 : 56);
      const rLines = estimateTextLines(rText, isChallan ? 26 : 32);
      const metaLines = Math.max(1, lLines, rLines);

      const row = ws.addRow([]);
      row.height = calculateCompactRowHeight(metaLines);

      // Left Column (Messers, Vessel Name, Port / Berth, Address)
      const leftCell = ws.getCell(currentRow, 1);
      if (lItem) {
        leftCell.value = {
          richText: [
            { text: `${lItem.label} `, font: { name: "Arial", size: 7.5, bold: true, color: { argb: "FF0F172A" } } },
            { text: lItem.value || "", font: { name: "Arial", size: 7.5, bold: !!lItem.bold, color: { argb: "FF000000" } } },
          ],
        };
        leftCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true, indent: 1 };
      } else {
        leftCell.value = "";
      }
      ws.mergeCells(currentRow, 1, currentRow, leftColEnd);

      // Right Column (Invoice / Challan / Quotation No, Date, Requisition No, PO Number)
      const rightCell = ws.getCell(currentRow, rightColStart);
      if (rItem) {
        rightCell.value = {
          richText: [
            { text: `${rItem.label} `, font: { name: "Arial", size: 7.5, bold: true, color: { argb: "FF0F172A" } } },
            { text: rItem.value || "", font: { name: "Arial", size: 7.5, bold: !!rItem.bold, color: { argb: "FF000000" } } },
          ],
        };
        rightCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true, indent: 1 };
      } else {
        rightCell.value = "";
      }
      ws.mergeCells(currentRow, rightColStart, currentRow, colCount);

      // Ensure Left and Right compartments across all rows are completely, solidly bordered
      // In ExcelJS, assigning border & fill directly to the master cell applies the full border
      // (top, bottom, left, right) to the entire merged block without slave cells clearing the left border.
      leftCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFF8FAFC" },
      };
      leftCell.border = THIN_BORDER;

      rightCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFF8FAFC" },
      };
      rightCell.border = THIN_BORDER;

      currentRow++;
    }

    // Clean gap before table to ensure metadata box and item table border do not collide
    const gapRow = ws.addRow([]);
    gapRow.height = 10;
    currentRow++;

    // =========================================================================
    // TABLE HEADERS (Compact height)
    // =========================================================================
    const headerValues = isChallan
      ? ["SL", "Description of Marine Items / Spare Parts", "Qty", "Remarks / Unit"]
      : [
          "SL",
          "Description of Marine Items / Spare Parts",
          "Qty",
          "Unit",
          `Unit Price (${currencySymbol || "Tk"})`,
          `Total Price (${currencySymbol || "Tk"})`,
        ];

    const headerRow = ws.addRow(headerValues);
    headerRow.height = Math.max(TABLE_HEADER_HEIGHT, 18);
    for (let c = 1; c <= colCount; c++) {
      const cell = headerRow.getCell(c);
      cell.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF0F172A" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
      cell.border = THIN_BORDER;

      if (c === 1 || c === 3 || (!isChallan && c === 4)) {
        cell.alignment = { horizontal: "center", vertical: "middle" };
      } else if (c === 2 || (isChallan && c === 4)) {
        cell.alignment = { horizontal: "left", vertical: "middle" };
      } else {
        cell.alignment = { horizontal: "right", vertical: "middle" };
      }
    }
    currentRow++;

    // =========================================================================
    // TABLE ITEMS (Preserving Text Color, Font Size, Highlights, Bold, Italic)
    // =========================================================================
    page.items.forEach((item) => {
      const r = item.row;
      const qtyVal = parseNum(r.qty);
      const priceVal = parseNum(r.price);
      const amountVal = r.amount || (qtyVal * priceVal);

      let rowData: any[];
      if (isChallan) {
        rowData = [
          item.originalIndex,
          "", // description cell populated with RichText below
          qtyVal > 0 ? qtyVal : (r.qty || ""),
          r.unit || "",
        ];
      } else {
        rowData = [
          item.originalIndex,
          "", // description cell populated with RichText below
          qtyVal > 0 ? qtyVal : (r.qty || ""),
          r.unit || "",
          priceVal > 0 ? priceVal : "",
          amountVal > 0 ? amountVal : "",
        ];
      }

      const row = ws.addRow(rowData);
      // Explicitly set compact, line-accurate row height for every row so that line 2, line 3, etc.
      // are ALWAYS fully visible without being clipped or hidden behind cell borders in Microsoft Excel,
      // while keeping the layout sleek, compact, and strictly up to professional standards.
      row.height = item.height;

      // Description Cell (Col 2): Apply richText formatting & highlight fill
      const descCell = row.getCell(2);
      if (item.richDesc && item.richDesc.length > 0) {
        descCell.value = { richText: item.richDesc };
      } else {
        descCell.value = item.cleanDesc;
        descCell.font = { name: "Arial", size: 7.5, color: { argb: "FF000000" } };
      }

      // If user applied a background highlight/fill color, reflect it in Excel cell fill
      if (item.cellBgColor) {
        descCell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: item.cellBgColor },
        };
      }

      const vAlign = "middle";

      for (let colNumber = 1; colNumber <= colCount; colNumber++) {
        const cell = row.getCell(colNumber);
        cell.border = THIN_BORDER;

        if (colNumber === 1) {
          cell.font = { name: "Arial", size: 7.5, color: { argb: "FF000000" } };
          cell.alignment = { horizontal: "center", vertical: vAlign };
        } else if (colNumber === 2) {
          cell.alignment = { horizontal: "left", vertical: vAlign, wrapText: true };
        } else if (colNumber === 3) {
          cell.font = { name: "Arial", size: 7.5, color: { argb: "FF000000" } };
          cell.alignment = { horizontal: "center", vertical: vAlign };
          if (typeof cell.value === "number") {
            cell.numFmt = "#,##0.##";
          }
        } else if (colNumber === 4) {
          cell.font = { name: "Arial", size: 7.5, color: { argb: "FF000000" } };
          cell.alignment = { horizontal: isChallan ? "left" : "center", vertical: vAlign, wrapText: true };
        } else if (colNumber === 5) {
          cell.font = { name: "Arial", size: 7.5, color: { argb: "FF000000" } };
          cell.alignment = { horizontal: "right", vertical: vAlign };
          if (typeof cell.value === "number") {
            cell.numFmt = "#,##0.00";
          }
        } else if (colNumber === 6) {
          cell.font = { name: "Arial", size: 7.5, color: { argb: "FF000000" } };
          cell.alignment = { horizontal: "right", vertical: vAlign };
          if (typeof cell.value === "number") {
            cell.numFmt = "#,##0.00";
          }
        }
      }

      currentRow++;
    });

    // If intermediate page, add continuation notice row above the signature section
    if (!page.isLastPage) {
      const contRow = ws.addRow([`Continued on Page ${page.pageNumber + 1}...`]);
      contRow.height = 12;
      ws.mergeCells(`A${currentRow}:${lastColLetter}${currentRow}`);
      const cCell = ws.getCell(`A${currentRow}`);
      cCell.font = { name: "Arial", size: 7.5, italic: true, bold: true, color: { argb: "FF64748B" } };
      cCell.alignment = { horizontal: "right", vertical: "middle" };
      currentRow++;
    }

    // =========================================================================
    // TOTALS & SUMMARY SECTION (Only on Last Page for Quotation & Invoice)
    // =========================================================================
    if (page.isLastPage && !isChallan) {
      const wordsText = numberToWords(grandTotal, currency);
      const summaryStartRow = currentRow;

      // Sub Total
      const subTotalRow = ws.addRow(["", "", "", "", "Sub Total:", rowsTotal]);
      subTotalRow.height = 13;
      currentRow++;

      // Discount (if any)
      if (isInvoice && includeDiscount && discountAmount > 0) {
        const discRow = ws.addRow(["", "", "", "", "Discount:", -discountAmount]);
        discRow.height = 13;
        currentRow++;
      }

      // VAT (if any)
      if (isInvoice && vatAmount > 0) {
        const vatRow = ws.addRow(["", "", "", "", "VAT:", vatAmount]);
        vatRow.height = 13;
        currentRow++;
      }

      // Transportation (if any)
      if (isInvoice && parsedTransport > 0) {
        const transRow = ws.addRow(["", "", "", "", "Transportation:", parsedTransport]);
        transRow.height = 13;
        currentRow++;
      }

      // Grand Total / Net Payable
      const grandTotalLabel = isInvoice ? "Net Payable:" : "Grand Total:";
      const grandTotalRow = ws.addRow(["", "", "", "", grandTotalLabel, grandTotal]);
      grandTotalRow.height = 15;
      const grandTotalRowIndex = currentRow;
      currentRow++;

      const summaryEndRow = currentRow - 1;

      // Merge columns A-D across summary rows for "In Words" section
      ws.mergeCells(`A${summaryStartRow}:D${summaryEndRow}`);
      const wordsCell = ws.getCell(`A${summaryStartRow}`);
      wordsCell.value = `In Words:\n${wordsText}`;
      wordsCell.font = { name: "Arial", size: 7.5, bold: true, italic: true, color: { argb: "FF0F172A" } };
      wordsCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      wordsCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
      wordsCell.border = THIN_BORDER;

      // Style totals cells (Columns E & F)
      for (let r = summaryStartRow; r <= summaryEndRow; r++) {
        const labelCell = ws.getCell(r, 5);
        const valCell = ws.getCell(r, 6);
        const isGrandTotal = r === grandTotalRowIndex;

        labelCell.font = {
          name: "Arial",
          size: isGrandTotal ? 8 : 7.5,
          bold: true,
          color: { argb: "FF000000" },
        };
        labelCell.alignment = { horizontal: "right", vertical: "middle" };

        valCell.font = {
          name: "Arial",
          size: isGrandTotal ? 8.5 : 7.5,
          bold: true,
          color: { argb: "FF000000" },
        };
        valCell.alignment = { horizontal: "right", vertical: "middle" };
        valCell.numFmt = "#,##0.00";

        if (isGrandTotal) {
          labelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
          valCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
          labelCell.border = {
            top: { style: "thin", color: { argb: BORDER_COLOR } },
            bottom: { style: "double", color: { argb: BORDER_COLOR } },
            left: { style: "thin", color: { argb: BORDER_COLOR } },
            right: { style: "thin", color: { argb: BORDER_COLOR } },
          };
          valCell.border = {
            top: { style: "thin", color: { argb: BORDER_COLOR } },
            bottom: { style: "double", color: { argb: BORDER_COLOR } },
            left: { style: "thin", color: { argb: BORDER_COLOR } },
            right: { style: "thin", color: { argb: BORDER_COLOR } },
          };
        } else {
          labelCell.border = THIN_BORDER;
          valCell.border = THIN_BORDER;
        }
      }
    }

    // =========================================================================
    // SIGNATURES & STAMP SECTION (Present on EVERY PAGE according to format)
    // =========================================================================
    // Clean, proper gap above signature area - compact so more items can be inserted naturally
    const sigGap = ws.addRow([]);
    sigGap.height = 16;
    currentRow++;

    // Row 1: "For [Company Name]" above Authorized Signature on the right (quotation & invoice only)
    if (!isChallan) {
      const forCompRow = ws.addRow([]);
      forCompRow.height = 12.5;
      const rightSigColStart = "E";
      ws.mergeCells(`${rightSigColStart}${currentRow}:${lastColLetter}${currentRow}`);
      const forCompCell = ws.getCell(`${rightSigColStart}${currentRow}`);
      forCompCell.value = `For ${currentCompany.name}`;
      forCompCell.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF000000" } };
      forCompCell.alignment = { horizontal: "center", vertical: "middle" };
      currentRow++;

      // Rows 2, 3, 4: Clear Stamp and Pen Signature Space (3 rows of 13pt = 39pt)
      const stampStartRowIndex = currentRow;
      for (let s = 0; s < 3; s++) {
        const spacer = ws.addRow([]);
        spacer.height = 13;
        currentRow++;
      }

      // Embed Stamp Image centered between "For [Company]" and "Authorized Signature" (if enabled)
      if (hasStamp && stampImageId !== null) {
        const stampWidth = 88;
        const stampHeight = 88;
        const nativeCol = 4;
        const nativeColOff = 620000;
        const nativeRow = stampStartRowIndex - 2;
        const nativeRowOff = 47625;

        ws.addImage(stampImageId, {
          tl: {
            nativeCol,
            nativeColOff,
            nativeRow,
            nativeRowOff,
          } as any,
          ext: { width: stampWidth, height: stampHeight },
        });
      }
    } else {
      // Challan format: Only receiving signature space (4 empty spacer rows for signing space)
      for (let s = 0; s < 4; s++) {
        const spacer = ws.addRow([]);
        spacer.height = 13;
        currentRow++;
      }
    }

    // Row 5: Signature Lines & Headings
    const sigLineRow = ws.addRow([]);
    sigLineRow.height = 13;

    // Left Box: Receiver's Signature (Cols A-B)
    ws.mergeCells(`A${currentRow}:B${currentRow}`);
    const recCell = ws.getCell(`A${currentRow}`);
    recCell.value = "Receiver's Signature";
    recCell.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF000000" } };
    recCell.alignment = { horizontal: "center", vertical: "middle" };
    ws.getCell(`A${currentRow}`).border = { top: { style: "medium", color: { argb: BORDER_COLOR } } };
    ws.getCell(`B${currentRow}`).border = { top: { style: "medium", color: { argb: BORDER_COLOR } } };

    // Right Box: Authorized Signature (Cols E-F, only for Quotation and Invoice)
    if (!isChallan) {
      const rightSigColStart = "E";
      ws.mergeCells(`${rightSigColStart}${currentRow}:${lastColLetter}${currentRow}`);
      const authCell = ws.getCell(`${rightSigColStart}${currentRow}`);
      authCell.value = "Authorized Signature";
      authCell.font = { name: "Arial", size: 8, bold: true, color: { argb: "FF000000" } };
      authCell.alignment = { horizontal: "center", vertical: "middle" };

      const startColNum = 5;
      for (let c = startColNum; c <= colCount; c++) {
        ws.getCell(currentRow, c).border = { top: { style: "medium", color: { argb: BORDER_COLOR } } };
      }
    }
    currentRow++;

    // Small gap before legal notice
    const preNoticeGap = ws.addRow([]);
    preNoticeGap.height = 3;
    currentRow++;

    // Legal notice footer
    const noticeRow = ws.addRow(["ITEMS ONCE SOLD ARE NON-RETURNABLE AND NON-EXCHANGEABLE."]);
    noticeRow.height = 11;
    ws.mergeCells(`A${currentRow}:${lastColLetter}${currentRow}`);
    const noticeCell = ws.getCell(`A${currentRow}`);
    noticeCell.font = { name: "Arial", size: 6.5, bold: true, color: { argb: "FF000000" } };
    noticeCell.alignment = { horizontal: "center", vertical: "middle" };
    currentRow++;
  });

  // =========================================================================
  // FILE DOWNLOAD TRIGGER
  // =========================================================================
  const filePrefix = isChallan ? "Challan" : isInvoice ? "Invoice" : "Quotation";
  const identifier = isChallan
    ? (challanNo || "NEW")
    : isInvoice
    ? (invoiceNo || "NEW")
    : (quotationNo || requisitionNo || "NEW");

  const filename = `${filePrefix}_${identifier.replace(/[\/\\?%*:|"<>\s]/g, "_")}.xlsx`;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}
