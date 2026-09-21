import React, { useRef, useEffect, useCallback } from "react";
import { saveCurrentSelection } from "../utils/textFormatter";

interface RichTextCellProps {
  value: string;
  onChange: (value: string) => void;
  onFocus?: () => void;
  onBlur?: () => void;
  onKeyDown?: (e: React.KeyboardEvent<HTMLDivElement>) => void;
  onPaste?: (e: React.ClipboardEvent<HTMLDivElement>) => void;
  onClick?: (e: React.MouseEvent<HTMLDivElement>) => void;
  style?: React.CSSProperties;
  className?: string;
  dataRow?: number;
  dataCol?: number;
  syncId?: string;
  placeholder?: string;
  readOnly?: boolean;
}

export const RichTextCell: React.FC<RichTextCellProps> = ({
  value,
  onChange,
  onFocus,
  onBlur,
  onKeyDown,
  onPaste,
  onClick,
  style,
  className,
  dataRow,
  dataCol,
  syncId,
  placeholder,
  readOnly = false,
}) => {
  const editorRef = useRef<HTMLDivElement>(null);
  const isComposingRef = useRef(false);
  const lastEmittedValueRef = useRef(value || "");

  // Synchronize with paired print preview DOM element immediately
  const syncToPrintElement = useCallback((html: string) => {
    if (!syncId || typeof document === "undefined") return;
    const printEl = document.getElementById(`print-${syncId}`);
    if (printEl) {
      printEl.innerHTML = html || "&nbsp;";
    }
  }, [syncId]);

  // Synchronize external value with innerHTML.
  // If value changed externally (e.g. from Excel paste, undo, redo, or row changes),
  // update innerHTML immediately so stale DOM content does not persist or overwrite.
  useEffect(() => {
    if (editorRef.current) {
      const currentHTML = editorRef.current.innerHTML;
      const normalizedValue = value || "";
      const isTypingInThisElement =
        document.activeElement === editorRef.current &&
        normalizedValue === lastEmittedValueRef.current;

      if (currentHTML !== normalizedValue && !isTypingInThisElement) {
        editorRef.current.innerHTML = normalizedValue;
        lastEmittedValueRef.current = normalizedValue;
      }
      syncToPrintElement(normalizedValue);
    }
  }, [value, syncToPrintElement]);

  const handleInput = useCallback(() => {
    if (editorRef.current && !isComposingRef.current) {
      const html = editorRef.current.innerHTML;
      lastEmittedValueRef.current = html;
      onChange(html);
      syncToPrintElement(html);
      saveCurrentSelection();
    }
  }, [onChange, syncToPrintElement]);

  // Native input event listener to capture programmatic and execCommand DOM mutations immediately
  useEffect(() => {
    const el = editorRef.current;
    if (!el) return;

    const onNativeInput = () => {
      if (!isComposingRef.current) {
        const html = el.innerHTML;
        onChange(html);
        syncToPrintElement(html);
        saveCurrentSelection();
      }
    };

    el.addEventListener("input", onNativeInput);
    return () => {
      el.removeEventListener("input", onNativeInput);
    };
  }, [onChange, syncToPrintElement]);

  // Handle typing to ensure formatting resets for subsequent words in the same cell
  const handleCellKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
    onKeyDown?.(e);
    if (e.defaultPrevented) return;

    // When space is pressed at the boundary of a styled span (highlighted or colored text),
    // break out into an unstyled text node so the next word resets to clean default text
    if (e.key === " " && !e.shiftKey && !e.ctrlKey && !e.metaKey) {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && sel.isCollapsed) {
        const range = sel.getRangeAt(0);
        const node = range.startContainer;
        const styledSpan = (node instanceof HTMLElement ? node : node.parentElement)?.closest("span[style], b, i, u, mark, font");

        if (styledSpan && styledSpan !== editorRef.current && editorRef.current?.contains(styledSpan)) {
          const isAtEnd = node.nodeType === Node.TEXT_NODE && range.startOffset === (node.textContent?.length || 0);
          if (isAtEnd) {
            e.preventDefault();
            // Insert space after the styled span
            const spaceText = document.createTextNode("\u00A0");
            if (styledSpan.nextSibling) {
              styledSpan.parentNode?.insertBefore(spaceText, styledSpan.nextSibling);
            } else {
              styledSpan.parentNode?.appendChild(spaceText);
            }

            const newRange = document.createRange();
            newRange.setStartAfter(spaceText);
            newRange.collapse(true);
            sel.removeAllRanges();
            sel.addRange(newRange);

            handleInput();
            return;
          }
        }
      }
    }
  };

  const handleCellPaste = (e: React.ClipboardEvent<HTMLDivElement>) => {
    if (onPaste) {
      onPaste(e);
      if (e.defaultPrevented) return;
    }
    // Paste plain text to avoid foreign styling contamination
    e.preventDefault();
    const text = e.clipboardData.getData("text/plain");
    document.execCommand("insertText", false, text);
    handleInput();
  };

  return (
    <div
      ref={editorRef}
      contentEditable={!readOnly}
      suppressContentEditableWarning
      data-row={dataRow}
      data-col={dataCol}
      data-sync-id={syncId}
      onInput={handleInput}
      onFocus={(e) => {
        saveCurrentSelection();
        onFocus?.();
      }}
      onBlur={() => {
        onBlur?.();
      }}
      onClick={(e) => {
        saveCurrentSelection();
        onClick?.(e);
      }}
      onMouseUp={() => {
        saveCurrentSelection();
      }}
      onKeyUp={() => {
        saveCurrentSelection();
      }}
      onKeyDown={handleCellKeyDown}
      onPaste={handleCellPaste}
      onCompositionStart={() => {
        isComposingRef.current = true;
      }}
      onCompositionEnd={() => {
        isComposingRef.current = false;
        handleInput();
      }}
      style={{
        outline: "none",
        minHeight: "1.2em",
        lineHeight: 1.3,
        wordBreak: "break-word",
        overflowWrap: "anywhere",
        whiteSpace: "pre-wrap",
        ...style,
      }}
      className={className}
      data-placeholder={placeholder}
    />
  );
};

export default RichTextCell;
