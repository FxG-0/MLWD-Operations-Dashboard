import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  FileSpreadsheet,
  Plus,
  Trash2,
  Copy,
  Download,
  Upload,
  Search,
  CheckCircle2,
  AlertTriangle,
  RefreshCw,
  ArrowUpDown,
  ChevronUp,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Check,
  Save,
  Filter,
  SlidersHorizontal,
  Sparkles,
  Info,
  Layers,
  Calendar,
  X,
  FileText,
  ShieldCheck,
  Eye,
  Lock,
  RotateCcw,
  RotateCw
} from "lucide-react";
import * as XLSX from "xlsx";
import { RowData, Abbreviation, validateRow, parseAnyDate, formatMonthName, UserRole, getHoleSizeFromSection, getSectionFromHoleSize } from "../types";
import { exportToFormattedExcel } from "../lib/excelExport";

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];

import { EmptyState } from "./ui/empty-state";

interface ExcelGridEditorProps {
  data: RowData[];
  abbreviations: Abbreviation[];
  onUpdateData: (newData: RowData[]) => void;
  isLightMode?: boolean;
  userRole?: UserRole;
  canEdit?: boolean;
  onRequestAuthModal?: () => void;
}

type SortField = keyof RowData | null;
type SortOrder = "asc" | "desc";

const SECTION_OPTIONS = ["4.75in", "6.75in", "8in", "9.5in"];
const HOLE_SIZE_OPTIONS = ['17.5"', '16"', '12.25"', '8.5"', '6.125"', '6"', '5 7/8"'];

function ExcelGridEditor({
  data,
  abbreviations,
  onUpdateData,
  isLightMode = false,
  userRole = "editor",
  canEdit = true,
  onRequestAuthModal
}: ExcelGridEditorProps) {
  const [localRows, setLocalRows] = useState<RowData[]>(data);
  const [selectedRowIds, setSelectedRowIds] = useState<Set<string>>(new Set());
  const [activeCell, setActiveCell] = useState<{ rowId: string; field: keyof RowData } | null>(null);
  const [searchTerm, setSearchTerm] = useState<string>(() => {
    try {
      return localStorage.getItem("excel_editor_search_term") || "";
    } catch {
      return "";
    }
  });
  const [sortField, setSortField] = useState<SortField>(() => {
    try {
      return (localStorage.getItem("excel_editor_sort_field") as SortField) || "id";
    } catch {
      return "id";
    }
  });
  const [sortOrder, setSortOrder] = useState<SortOrder>(() => {
    try {
      return (localStorage.getItem("excel_editor_sort_order") as SortOrder) || "asc";
    } catch {
      return "asc";
    }
  });
  const [columnFilters, setColumnFilters] = useState<Record<string, string[]>>(() => {
    try {
      const saved = localStorage.getItem("excel_editor_column_filters");
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });
  const [dateRangeFilters, setDateRangeFilters] = useState<Record<string, { start: string; end: string }>>(() => {
    try {
      const saved = localStorage.getItem("excel_editor_date_range_filters");
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });
  const [dateFilterTab, setDateFilterTab] = useState<Record<string, "range" | "list">>(() => {
    try {
      const saved = localStorage.getItem("excel_editor_date_filter_tab");
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });
  const [calYear, setCalYear] = useState<number>(2026);
  const [calMonth, setCalMonth] = useState<number>(0);
  const [openFilterCol, setOpenFilterCol] = useState<keyof RowData | null>(null);
  const [colSearchTerm, setColSearchTerm] = useState<string>("");
  const [pendingColumnFilter, setPendingColumnFilter] = useState<string[] | null>(null);
  const [pendingDateRange, setPendingDateRange] = useState<{ start: string; end: string }>({ start: "", end: "" });
  const [cellDatePicker, setCellDatePicker] = useState<{ rowId: string; field: "date" } | null>(null);
  const [pickerYear, setPickerYear] = useState<number>(2026);
  const [pickerMonth, setPickerMonth] = useState<number>(new Date().getMonth());
  const [saveToast, setSaveToast] = useState<string | null>(null);

  const handleOpenCellDatePicker = (rowId: string, field: "date", currentDate: string) => {
    if (!canEdit) return;

    let y = 2026;
    let m = new Date().getMonth();

    if (currentDate) {
      const iso = parseDateToISO(currentDate);
      if (iso) {
        const d = new Date(iso);
        if (!isNaN(d.getTime())) {
          y = d.getFullYear();
          m = d.getMonth();
        }
      }
    }

    setPickerYear(y);
    setPickerMonth(m);
    setCellDatePicker({ rowId, field });
  };
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [pasteDataText, setPasteDataText] = useState("");
  const tableContainerRef = useRef<HTMLDivElement>(null);
  const cellDebounceTimerRef = useRef<NodeJS.Timeout | null>(null);
  const isSelfUpdateRef = useRef(false);

  // History stack for Undo / Redo
  const [historyStack, setHistoryStack] = useState<RowData[][]>([]);
  const [futureStack, setFutureStack] = useState<RowData[][]>([]);

  // Sync local state when incoming data changes externally (skip if user is focusing an input field or self-updating)
  useEffect(() => {
    if (isSelfUpdateRef.current) {
      isSelfUpdateRef.current = false;
      return;
    }
    const activeEl = document.activeElement;
    const isEditing = activeEl && (activeEl.tagName === "INPUT" || activeEl.tagName === "TEXTAREA" || activeEl.tagName === "SELECT");
    if (!isEditing) {
      setLocalRows(data);
    }
  }, [data]);

  // Clear active cell selection if read-only
  useEffect(() => {
    if (!canEdit) {
      setActiveCell(null);
    }
  }, [canEdit]);

  // Toast notification auto-hide
  useEffect(() => {
    if (saveToast) {
      const timer = setTimeout(() => setSaveToast(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [saveToast]);

  // Persist Excel Grid filter and search state across tab switches and page reloads
  useEffect(() => {
    try {
      localStorage.setItem("excel_editor_search_term", searchTerm);
    } catch {
      // ignore
    }
  }, [searchTerm]);

  useEffect(() => {
    try {
      localStorage.setItem("excel_editor_sort_field", sortField);
      localStorage.setItem("excel_editor_sort_order", sortOrder);
    } catch {
      // ignore
    }
  }, [sortField, sortOrder]);

  useEffect(() => {
    try {
      localStorage.setItem("excel_editor_column_filters", JSON.stringify(columnFilters));
    } catch {
      // ignore
    }
  }, [columnFilters]);

  useEffect(() => {
    try {
      localStorage.setItem("excel_editor_date_range_filters", JSON.stringify(dateRangeFilters));
    } catch {
      // ignore
    }
  }, [dateRangeFilters]);

  useEffect(() => {
    try {
      localStorage.setItem("excel_editor_date_filter_tab", JSON.stringify(dateFilterTab));
    } catch {
      // ignore
    }
  }, [dateFilterTab]);

  // Master handler to commit updated rows back to parent & server
  const commitRows = (
    updatedRows: RowData[],
    message: string = "Changes saved",
    skipHistory: boolean = false
  ) => {
    if (!canEdit) {
      return;
    }
    // Re-validate each row
    const validated = updatedRows.map((r) => {
      const errs = validateRow(r);
      return {
        ...r,
        validationErrors: errs
      };
    });

    if (!skipHistory) {
      setHistoryStack((prev) => {
        if (
          prev.length > 0 &&
          JSON.stringify(prev[prev.length - 1]) === JSON.stringify(localRows)
        ) {
          return prev;
        }
        return [...prev.slice(-49), localRows];
      });
      setFutureStack([]);
    }

    isSelfUpdateRef.current = true;
    setLocalRows(validated);
    onUpdateData(validated);
    setSaveToast(message);
  };

  const handleUndo = () => {
    if (!canEdit || historyStack.length === 0) return;

    const previousState = historyStack[historyStack.length - 1];
    const newHistory = historyStack.slice(0, historyStack.length - 1);

    setFutureStack((prev) => [localRows, ...prev.slice(0, 49)]);
    setHistoryStack(newHistory);

    setLocalRows(previousState);
    onUpdateData(previousState);
    setSaveToast("Undo: Action reverted");
  };

  const handleRedo = () => {
    if (!canEdit || futureStack.length === 0) return;

    const nextState = futureStack[0];
    const newFuture = futureStack.slice(1);

    setHistoryStack((prev) => [...prev.slice(-49), localRows]);
    setFutureStack(newFuture);

    setLocalRows(nextState);
    onUpdateData(nextState);
    setSaveToast("Redo: Action reapplied");
  };

  // Global Keyboard Shortcuts for Undo (Ctrl+Z / Cmd+Z) and Redo (Ctrl+Y / Cmd+Shift+Z / Ctrl+Shift+Z)
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (!canEdit) return;

      const isMac = typeof navigator !== "undefined" && navigator.platform.toUpperCase().indexOf("MAC") >= 0;
      const cmdOrCtrl = isMac ? e.metaKey : e.ctrlKey;

      if (cmdOrCtrl && e.key.toLowerCase() === "z") {
        if (e.shiftKey) {
          e.preventDefault();
          handleRedo();
        } else {
          e.preventDefault();
          handleUndo();
        }
      } else if (cmdOrCtrl && e.key.toLowerCase() === "y") {
        e.preventDefault();
        handleRedo();
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [canEdit, historyStack, futureStack, localRows]);


  // Default Column Widths state for Excel-like dragging
  const DEFAULT_COL_WIDTHS: Record<string, number> = {
    client: 140,
    rig: 110,
    well: 110,
    section: 110,
    holeSize: 110,
    qty: 80,
    configuration: 450,
    footage: 110,
    date: 110,
    month: 90,
    shopDate: 110,
    comments: 200,
  };

  const [colWidths, setColWidths] = useState<Record<string, number>>(DEFAULT_COL_WIDTHS);
  const [resizingCol, setResizingCol] = useState<string | null>(null);

  const handleResizeMouseDown = (e: React.MouseEvent, colKey: string) => {
    e.stopPropagation();
    e.preventDefault();
    setResizingCol(colKey);

    const startX = e.clientX;
    const startWidth = colWidths[colKey] || 120;

    const handleMouseMove = (moveEvent: MouseEvent) => {
      const deltaX = moveEvent.clientX - startX;
      const newWidth = Math.max(50, startWidth + deltaX);
      setColWidths((prev) => ({
        ...prev,
        [colKey]: newWidth,
      }));
    };

    const handleMouseUp = () => {
      setResizingCol(null);
      window.removeEventListener("mousemove", handleMouseMove);
      window.removeEventListener("mouseup", handleMouseUp);
    };

    window.addEventListener("mousemove", handleMouseMove);
    window.addEventListener("mouseup", handleMouseUp);
  };

  // Date & Month Helper functions for chronological sorting & formatting
  const parseDateToTime = (str: string): number => {
    if (!str || str.toUpperCase() === "N/A") return 0;
    const parts = str.split("-");
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const mIdx = months.findIndex((m) => m.toLowerCase() === parts[1].toLowerCase());
      let yr = parseInt(parts[2], 10);
      if (!isNaN(yr) && yr < 100) yr += 2000;
      if (!isNaN(day) && mIdx !== -1 && !isNaN(yr)) {
        return new Date(yr, mIdx, day).getTime();
      }
    }
    const t = Date.parse(str);
    return isNaN(t) ? 0 : t;
  };

  const getMonthIndex = (mStr: string): number => {
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return months.findIndex((m) => m.toLowerCase() === mStr.toLowerCase());
  };

  const MONTHS_MAP: { [key: string]: string } = useMemo(
    () => ({
      jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06",
      jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12"
    }),
    []
  );

  const parseDateToISO = (dateStr: string): string => {
    if (!dateStr || dateStr.toLowerCase() === "n/a") return "";
    const trimmed = dateStr.trim();

    if (/^\d+(\.\d+)?$/.test(trimmed)) {
      const numVal = Number(trimmed);
      if (numVal > 20000 && numVal < 60000) {
        const utcDays = numVal - 25569;
        const d = new Date(utcDays * 86400 * 1000);
        if (!isNaN(d.getTime())) {
          const year = d.getFullYear();
          const month = String(d.getMonth() + 1).padStart(2, "0");
          const day = String(d.getDate()).padStart(2, "0");
          return `${year}-${month}-${day}`;
        }
      }
    }

    const isoRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (isoRegex.test(trimmed)) return trimmed;

    const parts = trimmed.split(/[-/]/);
    if (parts.length === 3) {
      const monthStr = parts[1].toLowerCase();
      const monthFromMap = MONTHS_MAP[monthStr];
      if (monthFromMap) {
        let day = parts[0].padStart(2, "0");
        let year = parts[2];
        if (year.length === 2) year = "20" + year;
        return `${year}-${monthFromMap}-${day}`;
      }
      const p0 = parseInt(parts[0], 10);
      const p1 = parseInt(parts[1], 10);
      let p2 = parseInt(parts[2], 10);
      if (!isNaN(p0) && !isNaN(p1) && !isNaN(p2)) {
        if (p2 < 100) p2 += 2000;
        if (p1 <= 12 && p0 <= 31) {
          const day = String(p0).padStart(2, "0");
          const month = String(p1).padStart(2, "0");
          return `${p2}-${month}-${day}`;
        }
      }
    }

    const parsed = Date.parse(trimmed);
    if (!isNaN(parsed)) {
      const d = new Date(parsed);
      if (d.getFullYear() > 1900 && d.getFullYear() < 3000) {
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, "0");
        const day = String(d.getDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
      }
    }

    return "";
  };



  // Field change handler for a row
  const handleCellChange = (rowId: string, field: keyof RowData, value: any) => {
    if (!canEdit) {
      return;
    }
    const updated = localRows.map((row) => {
      if (row.id === rowId) {
        let newRow = { ...row, [field]: value };

        // Special handling for qty / footage / date / holeSize
        if (field === "qty") {
          if (value === "" || value === null || value === undefined) {
            (newRow as any).qty = "";
          } else {
            const num = parseInt(value, 10);
            newRow.qty = isNaN(num) ? ("" as any) : Math.max(1, num);
          }
        } else if (field === "footage") {
          if (value === "" || value === null || value === undefined) {
            newRow.footage = null;
          } else {
            const num = parseFloat(value);
            newRow.footage = isNaN(num) ? null : Math.max(0, num);
          }
        } else if (field === "holeSize") {
          newRow.holeSize = value;
        } else if (field === "date") {
          const valStr = typeof value === "string" ? value.trim() : "";
          if (!value || valStr === "" || valStr.toUpperCase() === "N/A") {
            newRow.date = "";
            newRow.month = "";
          } else if (/^\d{4}-\d{2}-\d{2}$/.test(valStr)) {
            // ISO date string selected from calendar picker (e.g., "2026-08-03")
            const [yStr, mStr, dStr] = valStr.split("-");
            const mNum = parseInt(mStr, 10);
            const dNum = parseInt(dStr, 10);
            const yNum = parseInt(yStr, 10) % 100;
            const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            if (mNum >= 1 && mNum <= 12) {
              const monthName = months[mNum - 1];
              const formattedDate = `${dNum}-${monthName}-${String(yNum).padStart(2, "0")}`;
              newRow.date = formattedDate;
              newRow.month = monthName;
            } else {
              newRow.date = value;
            }
          } else {
            newRow.date = value;
            // Only update month if value is a valid full date string with separators
            if (valStr.includes("-") || valStr.includes("/")) {
              const parsed = parseAnyDate(valStr);
              if (parsed.month && parsed.month !== "N/A" && parsed.month !== "") {
                newRow.month = parsed.month;
              }
            }
          }
        } else if (field === "shopDate") {
          const valStr = typeof value === "string" ? value.trim() : "";
          if (!value || valStr === "" || valStr.toUpperCase() === "N/A") {
            newRow.shopDate = "";
          } else {
            newRow.shopDate = value;
          }
        } else if (field === "section") {
          if (newRow.holeSize === undefined || newRow.holeSize === null) {
            newRow.holeSize = getHoleSizeFromSection(value);
          }
        } else if (field === "month") {
          newRow.month = formatMonthName(value);
        }

        const errors = validateRow(newRow);
        newRow.validationErrors = errors;
        return newRow;
      }
      return row;
    });

    // 1. Instant local state update for zero input lag and zero flickering
    setLocalRows(updated);

    // 2. Debounced commit to parent state
    if (cellDebounceTimerRef.current) {
      clearTimeout(cellDebounceTimerRef.current);
    }
    cellDebounceTimerRef.current = setTimeout(() => {
      commitRows(updated, `Updated row #${rowId}`, true);
    }, 250);
  };

  // Add new empty row at the very bottom
  const handleAddRow = () => {
    if (!canEdit) {
      return;
    }

    // Reset search, filters and sort to default ID order so the new empty row stays at the very bottom
    setSearchTerm("");
    setColumnFilters({});
    setDateRangeFilters({});
    setSortField("id");
    setSortOrder("asc");

    const nextIdNumber = localRows.reduce((max, r) => {
      const parsed = parseInt(r.id.replace(/\D/g, ""), 10);
      return !isNaN(parsed) && parsed > max ? parsed : max;
    }, localRows.length);

    const newId = (nextIdNumber + 1).toString();

    const newRow: RowData = {
      id: newId,
      client: "",
      rig: "",
      well: "",
      section: "",
      holeSize: "",
      qty: 1,
      configuration: "",
      footage: null,
      date: "",
      month: "",
      shopDate: "",
      comments: ""
    };

    newRow.validationErrors = validateRow(newRow);
    const updated = [...localRows, newRow];
    commitRows(updated, `Added new row #${newId}`);
    setActiveCell({ rowId: newId, field: "client" });

    // Automatically scroll to the bottom so the new empty row is immediately visible
    setTimeout(() => {
      if (tableContainerRef.current) {
        tableContainerRef.current.scrollTo({
          top: tableContainerRef.current.scrollHeight,
          behavior: "smooth"
        });
      }
      const el = document.getElementById(`grid-row-${newId}`);
      if (el) {
        el.scrollIntoView({ behavior: "smooth", block: "end" });
      }
    }, 60);
  };

  // Insert new empty row after a specific row ID
  const handleInsertRowAfter = (targetRowId: string) => {
    if (!canEdit) {
      return;
    }

    setSortField(null);

    const nextIdNumber = localRows.reduce((max, r) => {
      const parsed = parseInt(r.id.replace(/\D/g, ""), 10);
      return !isNaN(parsed) && parsed > max ? parsed : max;
    }, localRows.length);

    const newId = (nextIdNumber + 1).toString();

    const newRow: RowData = {
      id: newId,
      client: "",
      rig: "",
      well: "",
      section: "",
      holeSize: "",
      qty: 1,
      configuration: "",
      footage: null,
      date: "",
      month: "",
      shopDate: "",
      comments: ""
    };

    newRow.validationErrors = validateRow(newRow);

    const targetIndex = localRows.findIndex((r) => r.id === targetRowId);
    const updated = [...localRows];
    if (targetIndex !== -1) {
      updated.splice(targetIndex + 1, 0, newRow);
    } else {
      updated.push(newRow);
    }

    commitRows(updated, `Inserted row #${newId}`);
    setActiveCell({ rowId: newId, field: "client" });

    setTimeout(() => {
      const el = document.getElementById(`grid-row-${newId}`);
      if (el) {
        el.scrollIntoView({ behavior: "smooth", block: "nearest" });
      }
    }, 60);
  };

  // Insert new empty row at a specific index
  const handleInsertRowAt = (index: number) => {
    if (!canEdit) {
      return;
    }

    setSortField(null);

    const nextIdNumber = localRows.reduce((max, r) => {
      const parsed = parseInt(r.id.replace(/\D/g, ""), 10);
      return !isNaN(parsed) && parsed > max ? parsed : max;
    }, localRows.length);

    const newId = (nextIdNumber + 1).toString();

    const newRow: RowData = {
      id: newId,
      client: "",
      rig: "",
      well: "",
      section: "",
      holeSize: "",
      qty: 1,
      configuration: "",
      footage: null,
      date: "",
      month: "",
      shopDate: "",
      comments: ""
    };

    newRow.validationErrors = validateRow(newRow);

    const updated = [...localRows];
    updated.splice(index, 0, newRow);

    commitRows(updated, `Inserted row #${newId}`);
    setActiveCell({ rowId: newId, field: "client" });

    setTimeout(() => {
      const el = document.getElementById(`grid-row-${newId}`);
      if (el) {
        el.scrollIntoView({ behavior: "smooth", block: "nearest" });
      }
    }, 60);
  };

  // Duplicate active/selected row
  const handleDuplicateRow = (rowId: string) => {
    if (!canEdit) {
      return;
    }
    const target = localRows.find((r) => r.id === rowId);
    if (!target) return;

    const nextIdNumber = localRows.reduce((max, r) => {
      const parsed = parseInt(r.id.replace(/\D/g, ""), 10);
      return !isNaN(parsed) && parsed > max ? parsed : max;
    }, localRows.length);

    const newId = (nextIdNumber + 1).toString();
    const clonedRow: RowData = {
      ...target,
      id: newId,
      comments: target.comments ? `${target.comments} (Copy)` : "Cloned record"
    };
    clonedRow.validationErrors = validateRow(clonedRow);

    const idx = localRows.findIndex((r) => r.id === rowId);
    const updated = [...localRows];
    updated.splice(idx + 1, 0, clonedRow);
    commitRows(updated, `Duplicated row #${rowId} as #${newId}`);
  };

  // Single delete
  const handleDeleteRow = (rowId: string) => {
    if (!canEdit) {
      return;
    }
    const updated = localRows.filter((r) => r.id !== rowId);
    commitRows(updated, `Deleted record #${rowId}`);
    if (activeCell?.rowId === rowId) {
      setActiveCell(null);
    }
    const newSel = new Set(selectedRowIds);
    newSel.delete(rowId);
    setSelectedRowIds(newSel);
  };

  // Bulk duplicate selected rows right below them
  const handleBulkDuplicate = () => {
    if (!canEdit || selectedRowIds.size === 0) {
      return;
    }

    setSortField(null); // Clear active sorting so insertion order is maintained

    let nextIdNumber = localRows.reduce((max, r) => {
      const parsed = parseInt(r.id.replace(/\D/g, ""), 10);
      return !isNaN(parsed) && parsed > max ? parsed : max;
    }, localRows.length);

    // Get indices of selected rows in order of appearance in localRows
    const selectedIndices: number[] = [];
    localRows.forEach((r, idx) => {
      if (selectedRowIds.has(r.id)) {
        selectedIndices.push(idx);
      }
    });

    if (selectedIndices.length === 0) return;

    const newSelectedIds = new Set<string>();
    const updated = [...localRows];

    // Group selected indices into contiguous blocks
    const blocks: number[][] = [];
    let currentBlock: number[] = [];

    for (let i = 0; i < selectedIndices.length; i++) {
      const idx = selectedIndices[i];
      if (currentBlock.length === 0 || idx === currentBlock[currentBlock.length - 1] + 1) {
        currentBlock.push(idx);
      } else {
        blocks.push(currentBlock);
        currentBlock = [idx];
      }
    }
    if (currentBlock.length > 0) {
      blocks.push(currentBlock);
    }

    let firstNewId: string | null = null;

    // Process blocks from bottom to top so index splices don't offset earlier blocks
    for (let b = blocks.length - 1; b >= 0; b--) {
      const block = blocks[b];
      const lastIndexInBlock = block[block.length - 1];

      const clonedRowsForBlock: RowData[] = block.map((idx) => {
        const originalRow = localRows[idx];
        nextIdNumber += 1;
        const newId = nextIdNumber.toString();
        if (!firstNewId) firstNewId = newId;
        newSelectedIds.add(newId);

        const cloned: RowData = {
          ...originalRow,
          id: newId,
          comments: originalRow.comments
            ? originalRow.comments.includes("(Copy)")
              ? originalRow.comments
              : `${originalRow.comments} (Copy)`
            : "Cloned record"
        };
        cloned.validationErrors = validateRow(cloned);
        return cloned;
      });

      // Insert cloned rows right after the last row in the block
      updated.splice(lastIndexInBlock + 1, 0, ...clonedRowsForBlock);
    }

    commitRows(updated, `Duplicated ${selectedRowIds.size} selected row(s)`);
    setSelectedRowIds(newSelectedIds);

    if (firstNewId) {
      const focusId = firstNewId;
      setTimeout(() => {
        const el = document.getElementById(`grid-row-${focusId}`);
        if (el) {
          el.scrollIntoView({ behavior: "smooth", block: "nearest" });
        }
      }, 60);
    }
  };

  // Bulk delete selected
  const handleBulkDelete = () => {
    if (!canEdit) {
      return;
    }
    if (selectedRowIds.size === 0) return;
    const updated = localRows.filter((r) => !selectedRowIds.has(r.id));
    commitRows(updated, `Deleted ${selectedRowIds.size} selected records`);
    setSelectedRowIds(new Set());
    setActiveCell(null);
  };

  // Toggle select row
  const handleToggleSelectRow = (rowId: string) => {
    const newSel = new Set(selectedRowIds);
    if (newSel.has(rowId)) {
      newSel.delete(rowId);
    } else {
      newSel.add(rowId);
    }
    setSelectedRowIds(newSel);
  };

  // Select all visible rows
  const handleToggleSelectAll = () => {
    if (selectedRowIds.size >= processedRows.length && processedRows.length > 0) {
      setSelectedRowIds(new Set());
    } else {
      setSelectedRowIds(new Set(processedRows.map((r) => r.id)));
    }
  };

  // Export updated records to .xlsx
  const handleExportXLSX = () => {
    exportToFormattedExcel(localRows, {
      fileName: `Operational_Logs_Export_${new Date().toISOString().slice(0, 10)}.xlsx`,
      sheetName: "Operational Logs"
    });
    setSaveToast("Exported spreadsheet workbook with formatting & layout!");
  };

  // Paste / Import Tab Delimited Text
  const handleImportPastedText = () => {
    if (!pasteDataText.trim()) return;

    const lines = pasteDataText.trim().split(/\r?\n/);
    const newParsedRows: RowData[] = [];

    lines.forEach((line, idx) => {
      const cols = line.split("\t").map((c) => c.trim());
      if (cols.length < 3) return; // skip junk lines

      // Header row check
      if (cols[0].toLowerCase().includes("client") || cols[0].toLowerCase().includes("id")) return;

      const rowId = (localRows.length + idx + 1).toString();
      
      let client = "Operator";
      let rig = "Rig";
      let well = "Well";
      let section = '8.5"';
      let holeSize = '12.25"';
      let qty = 1;
      let config = "IC - PST - RST";
      let footage: number | null = null;
      let date = "";
      let month = "";
      let shopDate = "";
      let comments = "";

      if (cols.length >= 12) {
        // 12 or more columns pasted (explicit separate holeSize column)
        client = cols[1] || "";
        rig = cols[2] || "";
        well = cols[3] || "";
        section = cols[4] || "";
        holeSize = cols[5] || getHoleSizeFromSection(section);
        const parsedQty = parseInt(cols[6], 10);
        qty = isNaN(parsedQty) ? 1 : Math.max(1, parsedQty);
        config = cols[7] || "";
        const parsedFt = parseFloat(cols[8]);
        footage = isNaN(parsedFt) ? null : parsedFt;
        date = cols[9] || "";
        month = cols[10] || "";
        shopDate = cols[11] || "";
        comments = cols[12] || "";
      } else {
        // Standard 11-column Excel layout (Section column directly populates holeSize and section)
        client = cols[1] || "";
        rig = cols[2] || "";
        well = cols[3] || "";
        section = cols[4] || "";
        holeSize = cols[4] || "";
        const parsedQty = parseInt(cols[5], 10);
        qty = isNaN(parsedQty) ? 1 : Math.max(1, parsedQty);
        config = cols[6] || "";
        const parsedFt = parseFloat(cols[7]);
        footage = isNaN(parsedFt) ? null : parsedFt;
        date = cols[8] || "";
        month = cols[9] || "";
        shopDate = cols[10] || "";
        comments = cols[11] || "";
      }

      const r: RowData = {
        id: cols[0] && !isNaN(parseInt(cols[0], 10)) ? cols[0] : rowId,
        client,
        rig,
        well,
        section,
        holeSize,
        qty,
        configuration: config,
        footage,
        date,
        month,
        shopDate,
        comments
      };
      r.validationErrors = validateRow(r);
      newParsedRows.push(r);
    });

    if (newParsedRows.length > 0) {
      const merged = [...newParsedRows, ...localRows];
      commitRows(merged, `Imported ${newParsedRows.length} rows from pasted text`);
      setPasteDataText("");
      setIsImportModalOpen(false);
    }
  };

  // Sorting
  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortOrder(sortOrder === "asc" ? "desc" : "asc");
    } else {
      setSortField(field);
      setSortOrder("asc");
    }
  };

  // Column Filter Helpers
  const getColumnUniqueValues = (field: keyof RowData) => {
    const countsMap = new Map<string, number>();
    localRows.forEach((row) => {
      let rawVal = row[field];
      if (field === "month") {
        rawVal = formatMonthName(rawVal);
      }
      const val = rawVal !== undefined && rawVal !== null ? String(rawVal) : "";
      countsMap.set(val, (countsMap.get(val) || 0) + 1);
    });

    const uniqueList = Array.from(countsMap.entries()).map(([value, count]) => ({
      value,
      count
    }));

    uniqueList.sort((a, b) => {
      if (field === "date" || field === "shopDate") {
        const timeA = parseDateToTime(a.value);
        const timeB = parseDateToTime(b.value);
        return timeA - timeB;
      }
      if (field === "month") {
        const idxA = getMonthIndex(a.value);
        const idxB = getMonthIndex(b.value);
        if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      }
      const numA = parseFloat(a.value);
      const numB = parseFloat(b.value);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return a.value.localeCompare(b.value);
    });

    return uniqueList;
  };

  const handleTogglePendingValue = (
    value: string,
    allValues: string[]
  ) => {
    setPendingColumnFilter((prev) => {
      const currentSelected = prev ?? allValues;
      const currentSet = new Set(currentSelected);

      if (currentSet.has(value)) {
        currentSet.delete(value);
      } else {
        currentSet.add(value);
      }

      return Array.from(currentSet);
    });
  };

  const handleSelectAllPendingValues = (
    allValues: string[],
    visibleValues?: string[]
  ) => {
    setPendingColumnFilter((prev) => {
      const currentSelected = prev ?? allValues;
      const currentSet = new Set(currentSelected);
      const targetValues = visibleValues && visibleValues.length > 0 ? visibleValues : allValues;

      const isAllTargetSelected = targetValues.every((v) => currentSet.has(v));

      if (isAllTargetSelected) {
        targetValues.forEach((v) => currentSet.delete(v));
      } else {
        targetValues.forEach((v) => currentSet.add(v));
      }

      return Array.from(currentSet);
    });
  };

  const handleApplyColumnFilter = (
    field: keyof RowData,
    allValues: string[]
  ) => {
    const fieldStr = String(field);
    const isDateCol = field === "date" || field === "shopDate";
    const activeTabMode = dateFilterTab[fieldStr] || (isDateCol ? "range" : "list");

    if (activeTabMode === "range") {
      if (pendingDateRange.start || pendingDateRange.end) {
        setDateRangeFilters((prev) => ({
          ...prev,
          [fieldStr]: { ...pendingDateRange }
        }));
      } else {
        setDateRangeFilters((prev) => {
          const copy = { ...prev };
          delete copy[fieldStr];
          return copy;
        });
      }
    } else {
      if (pendingColumnFilter !== null) {
        const isAllSelected =
          allValues.length > 0 &&
          pendingColumnFilter.length === allValues.length &&
          allValues.every((v) => pendingColumnFilter.includes(v));

        if (isAllSelected) {
          setColumnFilters((prev) => {
            const copy = { ...prev };
            delete copy[fieldStr];
            return copy;
          });
        } else {
          setColumnFilters((prev) => ({
            ...prev,
            [fieldStr]: [...pendingColumnFilter]
          }));
        }
      }
    }

    setOpenFilterCol(null);
  };

  const handleOpenFilterPopover = (colKey: keyof RowData) => {
    const uniqueVals = getColumnUniqueValues(colKey);
    const allValues = uniqueVals.map((uv) => uv.value);
    const existingColFilter = columnFilters[colKey];

    if (existingColFilter !== undefined) {
      setPendingColumnFilter([...existingColFilter]);
    } else {
      setPendingColumnFilter([...allValues]);
    }

    if (dateRangeFilters[colKey]) {
      setPendingDateRange({ ...dateRangeFilters[colKey] });
    } else {
      setPendingDateRange({ start: "", end: "" });
    }

    setColSearchTerm("");

    const isDateCol = colKey === "date" || colKey === "shopDate";
    if (isDateCol) {
      const activeDateRange = dateRangeFilters[colKey];
      if (activeDateRange?.start) {
        const d = new Date(activeDateRange.start);
        if (!isNaN(d.getTime())) {
          setCalYear(d.getFullYear());
          setCalMonth(d.getMonth());
        }
      } else {
        const firstRow = localRows.find(
          (r) => r[colKey] && parseDateToISO(String(r[colKey]))
        );
        if (firstRow) {
          const iso = parseDateToISO(String(firstRow[colKey]));
          if (iso) {
            const d = new Date(iso);
            if (!isNaN(d.getTime())) {
              setCalYear(d.getFullYear());
              setCalMonth(d.getMonth());
            }
          }
        }
      }
    }

    setOpenFilterCol(colKey);
  };

  const handleClearColumnFilter = (field: keyof RowData) => {
    const fieldStr = String(field);
    setColumnFilters((prev) => {
      const copy = { ...prev };
      delete copy[fieldStr];
      return copy;
    });
    setDateRangeFilters((prev) => {
      const copy = { ...prev };
      delete copy[fieldStr];
      return copy;
    });
    setPendingColumnFilter(null);
    setPendingDateRange({ start: "", end: "" });
    setOpenFilterCol(null);
  };

  const handleResetAllFilters = () => {
    setSearchTerm("");
    setColumnFilters({});
    setDateRangeFilters({});
    setSortField("id");
    setSortOrder("asc");
  };

  const hasDateFilters = Object.values(dateRangeFilters).some(
    (f) => Boolean(f.start || f.end)
  );

  const isFiltered =
    searchTerm !== "" ||
    Object.keys(columnFilters).length > 0 ||
    hasDateFilters;

  // Filter & Search
  const processedRows = useMemo(() => {
    let result = [...localRows];

    // Global Search term
    if (searchTerm.trim()) {
      const term = searchTerm.trim();
      const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const regex = new RegExp(`\\b${escaped}`, "i");
      result = result.filter(
        (r) =>
          regex.test(r.id) ||
          regex.test(r.client) ||
          regex.test(r.rig) ||
          regex.test(r.well) ||
          regex.test(r.configuration) ||
          regex.test(r.section) ||
          (r.holeSize && regex.test(r.holeSize)) ||
          (r.comments && regex.test(r.comments))
      );
    }

    // Column-specific AutoFilters (multi-select checklist)
    Object.entries(columnFilters).forEach(([colKey, selectedValues]) => {
      if (selectedValues !== undefined) {
        const lowerValSet = new Set(selectedValues.map((v) => v.trim().toLowerCase()));
        result = result.filter((row) => {
          let rawVal = row[colKey as keyof RowData];
          if (colKey === "month") {
            rawVal = formatMonthName(rawVal);
          }
          const valStr = rawVal !== undefined && rawVal !== null ? String(rawVal).trim().toLowerCase() : "";
          return lowerValSet.has(valStr);
        });
      }
    });

    // Date range filters
    Object.entries(dateRangeFilters).forEach(([colKey, range]) => {
      if (range && (range.start || range.end)) {
        result = result.filter((row) => {
          const rawVal = row[colKey as keyof RowData];
          if (!rawVal) return false;
          const rowISO = parseDateToISO(String(rawVal));
          if (!rowISO) return false;
          if (range.start && rowISO < range.start) return false;
          if (range.end && rowISO > range.end) return false;
          return true;
        });
      }
    });

    // Sort
    if (sortField) {
      result.sort((a, b) => {
        let valA = a[sortField];
        let valB = b[sortField];

        if (sortField === "id") {
          const numA = parseInt(String(valA ?? "").replace(/\D/g, ""), 10);
          const numB = parseInt(String(valB ?? "").replace(/\D/g, ""), 10);
          if (!isNaN(numA) && !isNaN(numB)) {
            return sortOrder === "asc" ? numA - numB : numB - numA;
          }
        }

        if (sortField === "date" || sortField === "shopDate") {
          const timeA = parseDateToTime(String(valA ?? ""));
          const timeB = parseDateToTime(String(valB ?? ""));
          return sortOrder === "asc" ? timeA - timeB : timeB - timeA;
        }

        if (sortField === "month") {
          valA = formatMonthName(valA);
          valB = formatMonthName(valB);
          const idxA = getMonthIndex(String(valA));
          const idxB = getMonthIndex(String(valB));
          if (idxA !== -1 && idxB !== -1) {
            return sortOrder === "asc" ? idxA - idxB : idxB - idxA;
          }
        }

        if (typeof valA === "number" && typeof valB === "number") {
          return sortOrder === "asc" ? valA - valB : valB - valA;
        }

        valA = String(valA ?? "").toLowerCase();
        valB = String(valB ?? "").toLowerCase();

        if (valA < valB) return sortOrder === "asc" ? -1 : 1;
        if (valA > valB) return sortOrder === "asc" ? 1 : -1;
        return 0;
      });
    }

    return result;
  }, [
    localRows,
    searchTerm,
    columnFilters,
    dateRangeFilters,
    sortField,
    sortOrder,
  ]);

  // Statistics
  const totalQty = useMemo(() => localRows.reduce((sum, r) => sum + (r.qty || 1), 0), [localRows]);
  const totalFootage = useMemo(
    () => localRows.reduce((sum, r) => sum + (r.footage && !isNaN(r.footage) ? r.footage * (r.qty || 1) : 0), 0),
    [localRows]
  );
  const errorCount = useMemo(
    () => localRows.filter((r) => r.validationErrors && r.validationErrors.length > 0).length,
    [localRows]
  );

  // Active cell info for the formula bar
  const activeRow = useMemo(() => {
    if (!activeCell) return null;
    return localRows.find((r) => r.id === activeCell.rowId) || null;
  }, [activeCell, localRows]);

  const activeCellValue = activeRow && activeCell ? String(activeRow[activeCell.field] ?? "") : "";

  // Column letters mapping
  const columnsList: { key: keyof RowData; letter: string; label: string; width: string }[] = [
    { key: "client", letter: "A", label: "Client", width: "w-32" },
    { key: "rig", letter: "B", label: "Rig", width: "w-28" },
    { key: "well", letter: "C", label: "Well ID", width: "w-28" },
    { key: "holeSize", letter: "D", label: "Hole Size", width: "w-28" },
    { key: "qty", letter: "E", label: "Qty", width: "w-20" },
    { key: "configuration", letter: "F", label: "Configuration & Sub-Tools", width: "w-[450px]" },
    { key: "footage", letter: "G", label: "Footage (ft)", width: "w-28" },
    { key: "date", letter: "H", label: "Log Date", width: "w-28" },
    { key: "month", letter: "I", label: "Month", width: "w-20" },
    { key: "shopDate", letter: "J", label: "Shop Status", width: "w-28" },
    { key: "comments", letter: "K", label: "Comments", width: "w-48" }
  ];

  return (
    <div
      className={`rounded-2xl border transition-colors shadow-2xl flex flex-col ${
        isLightMode
          ? "bg-white border-slate-200 text-slate-800"
          : "bg-[#0E1015] border-[#1E293B] text-[#E0E0E0]"
      }`}
      id="excel-sheet-editor-view"
    >
      {/* 0. Role Permission Status Banner */}
      {canEdit ? (
        <div
          className={`px-4 py-2.5 border-b text-xs font-sans flex flex-wrap items-center justify-between gap-2 rounded-t-2xl ${
            isLightMode
              ? "bg-blue-50 border-blue-200 text-blue-900"
              : "bg-blue-950/25 border-blue-500/20 text-blue-200"
          }`}
        >
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-4 h-4 text-blue-400 shrink-0" />
            <span>
              <strong>Editor Permissions Active:</strong> You have full write access to edit grid cells, add/delete records, paste spreadsheets, and save directly to the database.
            </span>
          </div>
        </div>
      ) : (
        <div
          className={`px-4 py-2.5 border-b text-xs font-sans flex flex-wrap items-center justify-between gap-2 rounded-t-2xl ${
            isLightMode
              ? "bg-amber-50 border-amber-200 text-amber-900"
              : "bg-amber-950/25 border-amber-500/20 text-amber-200"
          }`}
        >
          <div className="flex items-center gap-2">
            <Eye className="w-4 h-4 text-amber-400 shrink-0" />
            <span>
              <strong>Read-Only Viewer Mode:</strong> You can inspect, search, filter columns, and export data. Editing cells or deleting rows requires Editor credentials.
            </span>
          </div>
          <button
            type="button"
            onClick={onRequestAuthModal}
            className="px-2.5 py-1 rounded-lg bg-amber-500 hover:bg-amber-400 text-slate-950 font-extrabold text-[11px] cursor-pointer transition-colors shadow shrink-0 ml-auto flex items-center gap-1"
          >
            <Lock className="w-3 h-3" />
            Sign In as Editor
          </button>
        </div>
      )}

      {/* 1. Header Toolbar */}
      <div
        className={`px-4 py-2.5 border-b flex flex-wrap items-center justify-between gap-3 ${
          isLightMode ? "border-slate-200 bg-slate-100/80" : "border-[#1E293B] bg-[#14171F]"
        }`}
      >
        {/* Sheet Title / Record Info on Left */}
        <div className="flex items-center gap-2.5">
          <div className="p-1.5 rounded-lg bg-blue-100 text-blue-700 dark:bg-blue-950/80 dark:text-blue-300 shrink-0">
            <FileSpreadsheet className="w-4 h-4" />
          </div>
          <div className="flex items-center gap-2">
            <span className="text-xs font-bold text-slate-800 dark:text-slate-100">Sheet Actions</span>
            <span className="text-[11px] font-semibold px-2 py-0.5 rounded-full bg-slate-200 text-slate-700 dark:bg-slate-800 dark:text-slate-300">
              {processedRows.length} {processedRows.length === 1 ? "row" : "rows"}
            </span>
          </div>
        </div>

        {/* Action Buttons on Right */}
        <div className="flex flex-wrap items-center gap-2">
          {canEdit && (
            <div className="flex items-center gap-1 border-r border-slate-700/30 pr-2 mr-0.5">
              <button
                type="button"
                onClick={handleUndo}
                disabled={historyStack.length === 0}
                className={`p-2 text-xs font-bold rounded-lg border transition-all flex items-center gap-1.5 ${
                  historyStack.length > 0
                    ? isLightMode
                      ? "bg-slate-800 border-slate-900 text-white hover:bg-slate-900 cursor-pointer shadow-md"
                      : "bg-[#181B22] border-slate-700 text-slate-100 hover:bg-[#222733] cursor-pointer shadow-md"
                    : "opacity-40 cursor-not-allowed border-transparent text-slate-500"
                }`}
                title="Undo (Ctrl+Z / Cmd+Z)"
              >
                <RotateCcw className="w-3.5 h-3.5 text-blue-400" />
                <span className="hidden sm:inline">Undo</span>
              </button>
              <button
                type="button"
                onClick={handleRedo}
                disabled={futureStack.length === 0}
                className={`p-2 text-xs font-bold rounded-lg border transition-all flex items-center gap-1.5 ${
                  futureStack.length > 0
                    ? isLightMode
                      ? "bg-slate-800 border-slate-900 text-white hover:bg-slate-900 cursor-pointer shadow-md"
                      : "bg-[#181B22] border-slate-700 text-slate-100 hover:bg-[#222733] cursor-pointer shadow-md"
                    : "opacity-40 cursor-not-allowed border-transparent text-slate-500"
                }`}
                title="Redo (Ctrl+Y / Cmd+Shift+Z)"
              >
                <RotateCw className="w-3.5 h-3.5 text-blue-400" />
                <span className="hidden sm:inline">Redo</span>
              </button>
            </div>
          )}

          {canEdit && (
            <button
              type="button"
              onClick={handleAddRow}
              className="flex items-center gap-1.5 px-3 py-2 text-xs font-bold rounded-lg text-white bg-blue-700 hover:bg-blue-800 border border-blue-900 shadow-md transition-all cursor-pointer"
            >
              <Plus className="w-4 h-4 stroke-[2.5]" />
              Add Row
            </button>
          )}

          {canEdit ? (
            <button
              type="button"
              onClick={() => commitRows(localRows, "All records validated & synced")}
              className="flex items-center gap-1.5 px-3 py-2 text-xs font-bold rounded-lg text-white bg-emerald-800 hover:bg-emerald-900 border border-emerald-950 shadow-md transition-all cursor-pointer"
              title="Force immediate save & re-validation"
            >
              <Save className="w-4 h-4" />
              Save & Sync DB
            </button>
          ) : (
            <button
              type="button"
              disabled
              className="flex items-center gap-1.5 px-3 py-2 text-xs font-bold rounded-lg text-slate-400 bg-slate-900 border border-slate-800 opacity-50 cursor-not-allowed"
              title="Read-only mode — Save disabled"
            >
              <Lock className="w-3.5 h-3.5 text-amber-400" />
              Save & Sync DB
            </button>
          )}

          {canEdit && selectedRowIds.size > 0 && (
            <div className="flex items-center gap-1.5 animate-fade-in">
              <button
                type="button"
                onClick={handleBulkDuplicate}
                className="flex items-center gap-1.5 px-3 py-2 text-xs font-bold rounded-lg bg-indigo-800 hover:bg-indigo-900 border border-indigo-950 text-white shadow-md transition-all cursor-pointer"
                title="Duplicate selected row(s) right below"
                id="excel-editor-duplicate-selected-btn"
              >
                <Copy className="w-4 h-4" />
                <span>Duplicate ({selectedRowIds.size})</span>
              </button>
              <button
                type="button"
                onClick={handleBulkDelete}
                className="flex items-center gap-1.5 px-3 py-2 text-xs font-bold rounded-lg bg-rose-800 hover:bg-rose-900 border border-rose-950 text-white shadow-md transition-all cursor-pointer"
                title="Delete selected rows"
                id="excel-editor-delete-selected-btn"
              >
                <Trash2 className="w-4 h-4" />
                <span>Delete ({selectedRowIds.size})</span>
              </button>
            </div>
          )}

          {canEdit && (
            <button
              type="button"
              onClick={() => setIsImportModalOpen(true)}
              className={`flex items-center gap-1.5 px-3 py-2 text-xs font-bold rounded-lg border transition-all cursor-pointer ${
                isLightMode
                  ? "bg-purple-800 border-purple-950 text-white hover:bg-purple-900 shadow-sm"
                  : "bg-purple-900 border-purple-800 text-purple-100 hover:bg-purple-950 shadow-sm"
              }`}
            >
              <Upload className="w-4 h-4 text-purple-300" />
              Paste / Import
            </button>
          )}

          <button
            type="button"
            onClick={handleExportXLSX}
            className={`flex items-center gap-1.5 px-3 py-2 text-xs font-bold rounded-lg border transition-all cursor-pointer ${
              isLightMode
                ? "bg-slate-800 border-slate-900 text-amber-300 hover:bg-slate-900 shadow-sm"
                : "bg-[#181B22] border-slate-700 text-amber-300 hover:bg-[#222733] shadow-sm"
            }`}
          >
            <Download className="w-4 h-4 text-amber-400" />
            Export XLSX
          </button>
        </div>
      </div>

      {/* 2. Excel Formula & Active Cell Value Bar */}
      <div
        className={`px-4 py-2 border-b flex items-center gap-3 font-mono text-xs ${
          isLightMode ? "bg-slate-100/80 border-slate-200" : "bg-[#161920] border-[#1E293B]"
        }`}
      >
        <div
          className={`px-2.5 py-1 rounded font-bold border shrink-0 text-center ${
            activeCell
              ? "bg-blue-600/15 border-blue-500/40 text-blue-400"
              : isLightMode
              ? "bg-slate-200 border-slate-300 text-slate-600"
              : "bg-[#1F232D] border-white/10 text-slate-400"
          }`}
        >
          {activeCell
            ? `${columnsList.find((c) => c.key === activeCell.field)?.letter || "A"}${
                processedRows.findIndex((r) => r.id === activeCell.rowId) + 1
              }`
            : "A1"}
        </div>

        <span className="text-slate-500 font-bold font-sans">fx</span>

        <div className="flex-1 relative">
          <input
            type="text"
            disabled={!canEdit || !activeCell}
            readOnly={!canEdit}
            value={activeCellValue}
            onChange={(e) => {
              if (canEdit && activeCell) {
                handleCellChange(activeCell.rowId, activeCell.field, e.target.value);
              }
            }}
            placeholder={
              canEdit
                ? activeCell
                  ? `Editing ${String(activeCell.field)}...`
                  : "Click any cell below to edit"
                : "Read-Only Mode — Sign in as Editor to edit cells"
            }
            className={`w-full h-8 px-3 rounded text-xs font-mono border focus:outline-none transition-all ${
              isLightMode
                ? "bg-white border-slate-300 text-slate-900 focus:border-blue-500"
                : "bg-[#0E1015] border-white/10 text-[#F5F7FA] focus:border-blue-500"
            } ${!canEdit || !activeCell ? "opacity-60 cursor-not-allowed" : ""}`}
          />
        </div>

        {/* Live Filter Controls */}
        <div className="flex items-center gap-2 shrink-0">
          <div className="relative">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-slate-400" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Search sheet..."
              className={`h-8 pl-8 pr-3 w-40 text-xs rounded border focus:outline-none ${
                isLightMode
                  ? "bg-white border-slate-300 text-slate-800"
                  : "bg-[#0E1015] border-white/10 text-slate-200"
              }`}
            />
            {searchTerm && (
              <button
                type="button"
                onClick={() => setSearchTerm("")}
                className="absolute right-2 top-2 text-slate-400 hover:text-white"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          {isFiltered && (
            <button
              type="button"
              onClick={handleResetAllFilters}
              className="px-2.5 py-1.5 text-xs font-bold rounded bg-rose-800 hover:bg-rose-900 text-white border border-rose-950 shadow-md flex items-center gap-1 transition-all cursor-pointer"
              title="Reset all search & column filters"
            >
              <RefreshCw className="w-3 h-3" />
              Reset Filters
            </button>
          )}
        </div>
      </div>

      {/* Active Filter Chips Bar */}
      {isFiltered && (
        <div
          className={`px-4 py-2 border-b flex flex-wrap items-center gap-2 text-xs ${
            isLightMode ? "bg-blue-50/70 border-blue-200" : "bg-[#121620] border-blue-500/20"
          }`}
        >
          <span className="font-bold text-blue-400 flex items-center gap-1 shrink-0">
            <Filter className="w-3.5 h-3.5" />
            Active Filters ({processedRows.length} of {localRows.length} rows showing):
          </span>

          {searchTerm && (
            <div className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-blue-600/20 border border-blue-500/40 text-blue-300 text-[11px]">
              <span className="font-bold">Search:</span>
              <span className="max-w-[120px] truncate">{searchTerm}</span>
              <button type="button" onClick={() => setSearchTerm("")} className="hover:text-white p-0.5 cursor-pointer">
                <X className="w-3 h-3" />
              </button>
            </div>
          )}

          {Object.entries(columnFilters).map(([colKey, vals]) => {
            const colObj = columnsList.find((c) => c.key === colKey);
            const displayVal = vals.length === 0 ? "(None selected)" : vals.join(", ");
            return (
              <div
                key={colKey}
                className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-blue-600/20 border border-blue-500/40 text-blue-300 font-mono text-[11px]"
              >
                <span className="font-bold">{colObj?.label || colKey}:</span>
                <span className="max-w-[140px] truncate">{displayVal}</span>
                <button
                  type="button"
                  onClick={() => handleClearColumnFilter(colKey as keyof RowData)}
                  className="hover:text-white p-0.5 cursor-pointer"
                  title="Clear filter"
                >
                  <X className="w-3 h-3" />
                </button>
              </div>
            );
          })}

          {Object.entries(dateRangeFilters).map(([colKey, range]) => {
            if (!range || (!range.start && !range.end)) return null;
            const colObj = columnsList.find((c) => c.key === colKey);
            const labelStr =
              range.start && range.end
                ? `${range.start} to ${range.end}`
                : range.start
                ? `From ${range.start}`
                : `Until ${range.end}`;
            return (
              <div
                key={`daterange-${colKey}`}
                className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-blue-600/20 border border-blue-500/40 text-blue-300 font-mono text-[11px]"
              >
                <span className="font-bold">{colObj?.label || colKey}:</span>
                <span className="max-w-[160px] truncate">{labelStr}</span>
                <button
                  type="button"
                  onClick={() => handleClearColumnFilter(colKey as keyof RowData)}
                  className="hover:text-white p-0.5 cursor-pointer"
                >
                  <X className="w-3 h-3" />
                </button>
              </div>
            );
          })}

          <button
            type="button"
            onClick={handleResetAllFilters}
            className="text-[11px] font-bold text-slate-400 hover:text-white underline ml-auto cursor-pointer"
          >
            Clear All
          </button>
        </div>
      )}

      {/* Save Notification Toast */}
      {saveToast && (
        <div className="bg-emerald-600 text-white px-4 py-2 text-xs font-bold font-sans flex items-center justify-between shadow-lg animate-fade-in">
          <div className="flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 stroke-[2.5]" />
            <span>{saveToast}</span>
          </div>
          <span className="text-[10px] font-mono opacity-80">Synced to server DB</span>
        </div>
      )}

      {/* 3. Main Spreadsheet Table Grid */}
      <div ref={tableContainerRef} className={`overflow-x-auto max-h-[600px] overflow-y-auto select-none border-2 ${
        isLightMode ? "border-slate-500 bg-white" : "border-slate-700 bg-[#0E1015]"
      }`}>
        <table className="w-full border-collapse text-left font-sans text-xs table-fixed">
          <colgroup>
            <col style={{ width: "40px" }} />
            {columnsList.map((col) => (
              <col key={col.key} style={{ width: `${colWidths[col.key] || 120}px` }} />
            ))}
          </colgroup>

          <thead>
            {/* Column Letter Bar Row */}
            <tr
              className={`font-mono text-[11px] font-bold border-b-2 ${
                isLightMode ? "bg-slate-300 text-slate-800 border-slate-500" : "bg-[#181B22] text-slate-200 border-slate-700"
              }`}
            >
              <th className={`p-2 text-center border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} w-10`}>
                <input
                  type="checkbox"
                  checked={selectedRowIds.size >= processedRows.length && processedRows.length > 0}
                  onChange={handleToggleSelectAll}
                  className="rounded border-slate-600 cursor-pointer"
                />
              </th>
              {columnsList.map((col) => (
                <th
                  key={col.letter}
                  className={`p-2 text-center border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} relative select-none truncate`}
                  style={{ width: `${colWidths[col.key] || 120}px` }}
                >
                  {col.letter}
                  {/* Column Resize Handle */}
                  <div
                    onMouseDown={(e) => handleResizeMouseDown(e, col.key)}
                    onDoubleClick={(e) => {
                      e.stopPropagation();
                      setColWidths((prev) => ({ ...prev, [col.key]: DEFAULT_COL_WIDTHS[col.key] || 120 }));
                    }}
                    className={`absolute right-0 top-0 bottom-0 w-2.5 cursor-col-resize z-20 group/resize flex items-center justify-center hover:bg-blue-500/50 transition-colors ${
                      resizingCol === col.key ? "bg-blue-500/60" : ""
                    }`}
                    title="Drag edge to resize column (double-click to reset)"
                  >
                    <div
                      className={`w-[1px] h-full ${
                        resizingCol === col.key ? "bg-blue-300" : "bg-slate-600/40 group-hover/resize:bg-blue-400"
                      }`}
                    />
                  </div>
                </th>
              ))}
            </tr>

            {/* Column Title Row with Excel AutoFilter Popovers */}
            <tr
              className={`font-semibold border-b-2 ${
                isLightMode ? "bg-slate-200 text-slate-900 border-slate-500" : "bg-[#13161C] text-[#E0E0E0] border-slate-700"
              }`}
            >
              <th className={`p-2 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} text-center`}>
                <Check className="w-3.5 h-3.5 mx-auto text-slate-500" />
              </th>
              {columnsList.map((col) => {
                const isDateCol = col.key === "date" || col.key === "shopDate";
                const activeDateRange = dateRangeFilters[col.key] || { start: "", end: "" };
                const hasDateRange = Boolean(activeDateRange.start || activeDateRange.end);
                const rawColFilter = columnFilters[col.key];
                const isFiltered = rawColFilter !== undefined || hasDateRange;
                const isSorted = sortField === col.key;
                const isOpen = openFilterCol === col.key;
                const activeTabMode = dateFilterTab[col.key] || (isDateCol ? "range" : "list");
                const uniqueVals = getColumnUniqueValues(col.key);
                const allValues = uniqueVals.map((uv) => uv.value);

                const currentSelected =
                  isOpen && pendingColumnFilter !== null
                    ? pendingColumnFilter
                    : (rawColFilter ?? allValues);
                const activeSelectedSet = new Set(currentSelected);

                const filteredUniqueVals = uniqueVals.filter((uv) => {
                  if (!colSearchTerm.trim()) return true;
                  const escaped = colSearchTerm.trim().replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                  const regex = new RegExp(`\\b${escaped}`, "i");
                  return regex.test(uv.value);
                });
                const visibleValues = filteredUniqueVals.map((uv) => uv.value);

                const isAllVisibleSelected =
                  visibleValues.length > 0 && visibleValues.every((v) => activeSelectedSet.has(v));
                const isSomeVisibleSelected =
                  visibleValues.some((v) => activeSelectedSet.has(v));
                const isIndeterminate = !isAllVisibleSelected && isSomeVisibleSelected;

                return (
                  <th
                    key={col.key}
                    className={`p-2.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} relative select-none ${
                      isOpen ? "z-30" : "z-10"
                    } ${
                      isFiltered
                        ? isLightMode
                          ? "bg-blue-100/70 text-blue-900"
                          : "bg-blue-950/40 text-blue-300"
                        : ""
                    }`}
                    style={{ width: `${colWidths[col.key] || 120}px` }}
                  >
                    <div className="flex items-center justify-between gap-1 pr-1">
                      <span
                        onClick={() => handleSort(col.key)}
                        className="truncate cursor-pointer hover:text-blue-400 flex items-center gap-1 font-bold"
                        title="Click to sort column"
                      >
                        {col.label}
                        {isSorted && (
                          <span className="text-blue-400 font-mono text-[10px]">
                            {sortOrder === "asc" ? "▲" : "▼"}
                          </span>
                        )}
                      </span>

                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          if (isOpen) {
                            setOpenFilterCol(null);
                          } else {
                            handleOpenFilterPopover(col.key);
                          }
                        }}
                        className={`p-1 rounded transition-all cursor-pointer ${
                          isFiltered
                            ? "bg-blue-600 text-white shadow-sm ring-1 ring-blue-400"
                            : isLightMode
                            ? "hover:bg-slate-200 text-slate-500"
                            : "hover:bg-white/10 text-slate-400"
                        }`}
                        title={`Filter column: ${col.label}`}
                      >
                        <Filter className="w-3.5 h-3.5" />
                      </button>
                    </div>

                    {/* Column Resize Handle */}
                    <div
                      onMouseDown={(e) => handleResizeMouseDown(e, col.key)}
                      onDoubleClick={(e) => {
                        e.stopPropagation();
                        setColWidths((prev) => ({ ...prev, [col.key]: DEFAULT_COL_WIDTHS[col.key] || 120 }));
                      }}
                      className={`absolute right-0 top-0 bottom-0 w-2.5 cursor-col-resize z-20 group/resize flex items-center justify-center hover:bg-blue-500/50 transition-colors ${
                        resizingCol === col.key ? "bg-blue-500/60" : ""
                      }`}
                      title="Drag edge to resize column (double-click to reset)"
                    >
                      <div
                        className={`w-[1px] h-full ${
                          resizingCol === col.key ? "bg-blue-300" : "bg-slate-600/40 group-hover/resize:bg-blue-400"
                        }`}
                      />
                    </div>

                    {/* AutoFilter Popover Dropdown */}
                    {isOpen && (
                      <>
                        <div
                          className="fixed inset-0 z-40 cursor-default"
                          onClick={(e) => {
                            e.stopPropagation();
                            setOpenFilterCol(null);
                          }}
                        />
                        <div
                          onClick={(e) => e.stopPropagation()}
                          className={`absolute top-full left-0 mt-1 ${
                            isDateCol && activeTabMode === "range" ? "w-72 sm:w-80" : "w-64"
                          } rounded-xl border shadow-2xl z-50 p-3 font-sans text-xs text-left animate-fade-in ${
                            isLightMode
                              ? "bg-white border-slate-300 text-slate-800 shadow-xl"
                              : "bg-[#161922] border-slate-700 text-slate-200 shadow-2xl"
                          }`}
                        >
                          {/* Header */}
                          <div className="flex items-center justify-between pb-2 border-b border-slate-700/50 mb-2">
                            <span className="font-bold text-blue-400 text-xs flex items-center gap-1.5">
                              {isDateCol ? <Calendar className="w-3.5 h-3.5" /> : <Filter className="w-3.5 h-3.5" />} Filter {col.label}
                            </span>
                            <button
                              type="button"
                              onClick={() => setOpenFilterCol(null)}
                              className="text-slate-400 hover:text-white p-0.5"
                            >
                              <X className="w-4 h-4" />
                            </button>
                          </div>

                          {/* Mode Switcher for Date Columns */}
                          {isDateCol && (
                            <div className={`flex items-center gap-1 p-1 rounded-lg border mb-3 ${
                              isLightMode ? "bg-slate-100 border-slate-200" : "bg-[#0E1015] border-slate-700/40"
                            }`}>
                              <button
                                type="button"
                                onClick={() => setDateFilterTab((prev) => ({ ...prev, [col.key]: "range" }))}
                                className={`flex-1 py-1 px-2 text-[11px] font-bold rounded-md transition-all flex items-center justify-center gap-1 cursor-pointer ${
                                  activeTabMode === "range"
                                    ? "bg-blue-600 text-white shadow-sm"
                                    : "text-slate-400 hover:text-white"
                                }`}
                              >
                                <Calendar className="w-3 h-3" />
                                Calendar & Range
                              </button>
                              <button
                                type="button"
                                onClick={() => setDateFilterTab((prev) => ({ ...prev, [col.key]: "list" }))}
                                className={`flex-1 py-1 px-2 text-[11px] font-bold rounded-md transition-all flex items-center justify-center gap-1 cursor-pointer ${
                                  activeTabMode === "list"
                                    ? "bg-blue-600 text-white shadow-sm"
                                    : "text-slate-400 hover:text-white"
                                }`}
                              >
                                <Filter className="w-3 h-3" />
                                Specific List
                              </button>
                            </div>
                          )}

                          {/* Sort Controls */}
                          <div className="space-y-1 mb-3 text-[11px]">
                            <button
                              type="button"
                              onClick={() => {
                                setSortField(col.key);
                                setSortOrder("asc");
                              }}
                              className="w-full text-left px-2 py-1.5 rounded hover:bg-blue-500/10 flex items-center gap-2 cursor-pointer font-medium"
                            >
                              <ArrowUpDown className="w-3 h-3 text-blue-400" />
                              <span>
                                {isDateCol || col.key === "month"
                                  ? "Sort Oldest to Newest"
                                  : "Sort Ascending (A → Z)"}
                              </span>
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                setSortField(col.key);
                                setSortOrder("desc");
                              }}
                              className="w-full text-left px-2 py-1.5 rounded hover:bg-blue-500/10 flex items-center gap-2 cursor-pointer font-medium"
                            >
                              <ArrowUpDown className="w-3 h-3 text-blue-400 rotate-180" />
                              <span>
                                {isDateCol || col.key === "month"
                                  ? "Sort Newest to Oldest"
                                  : "Sort Descending (Z → A)"}
                              </span>
                            </button>
                          </div>

                          {/* Date Range Calendar View */}
                          {isDateCol && activeTabMode === "range" ? (
                            <div className="border-t border-slate-700/50 pt-2 mb-2">
                              {/* Quick Presets */}
                              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block mb-1">Presets:</span>
                              <div className="flex flex-wrap gap-1 mb-3">
                                {[
                                  { label: "All Time", start: "", end: "" },
                                  { label: "Jan 2026", start: "2026-01-01", end: "2026-01-31", yr: 2026, mo: 0 },
                                  { label: "Feb 2026", start: "2026-02-01", end: "2026-02-28", yr: 2026, mo: 1 },
                                  { label: "Mar 2026", start: "2026-03-01", end: "2026-03-31", yr: 2026, mo: 2 },
                                  { label: "Full Year 2026", start: "2026-01-01", end: "2026-12-31", yr: 2026, mo: 0 },
                                ].map((preset) => (
                                  <button
                                    key={preset.label}
                                    type="button"
                                    onClick={() => {
                                      setPendingDateRange({ start: preset.start, end: preset.end });
                                      if (preset.yr !== undefined) {
                                        setCalYear(preset.yr);
                                        setCalMonth(preset.mo);
                                      }
                                    }}
                                    className={`px-2 py-1 text-[10px] font-semibold rounded-md transition-all cursor-pointer ${
                                      pendingDateRange.start === preset.start && pendingDateRange.end === preset.end
                                        ? "bg-blue-600 text-white font-bold"
                                        : isLightMode
                                        ? "bg-slate-100 text-slate-700 hover:bg-slate-200"
                                        : "bg-white/5 text-slate-300 hover:bg-white/10"
                                    }`}
                                  >
                                    {preset.label}
                                  </button>
                                ))}
                              </div>

                              {/* From & To Inputs */}
                              <div className="grid grid-cols-2 gap-2 mb-3 text-[11px]">
                                <div>
                                  <label className="block text-[10px] text-slate-400 font-bold mb-1">From Date:</label>
                                  <input
                                    type="date"
                                    value={pendingDateRange.start}
                                    onChange={(e) => {
                                      const val = e.target.value;
                                      setPendingDateRange((prev) => ({ ...prev, start: val }));
                                      if (val) {
                                        const d = new Date(val);
                                        if (!isNaN(d.getTime())) {
                                          setCalYear(d.getFullYear());
                                          setCalMonth(d.getMonth());
                                        }
                                      }
                                    }}
                                    onKeyDown={(e) => {
                                      if (e.key === "Enter") {
                                        e.preventDefault();
                                        handleApplyColumnFilter(col.key, allValues);
                                      }
                                    }}
                                    className={`w-full px-2 py-1 text-xs rounded border focus:outline-none cursor-pointer ${
                                      isLightMode
                                        ? "bg-slate-50 border-slate-300 text-slate-800"
                                        : "bg-[#0E1015] border-slate-700 text-white"
                                    }`}
                                  />
                                </div>
                                <div>
                                  <label className="block text-[10px] text-slate-400 font-bold mb-1">To Date:</label>
                                  <input
                                    type="date"
                                    value={pendingDateRange.end}
                                    onChange={(e) => {
                                      const val = e.target.value;
                                      setPendingDateRange((prev) => ({ ...prev, end: val }));
                                    }}
                                    onKeyDown={(e) => {
                                      if (e.key === "Enter") {
                                        e.preventDefault();
                                        handleApplyColumnFilter(col.key, allValues);
                                      }
                                    }}
                                    className={`w-full px-2 py-1 text-xs rounded border focus:outline-none cursor-pointer ${
                                      isLightMode
                                        ? "bg-slate-50 border-slate-300 text-slate-800"
                                        : "bg-[#0E1015] border-slate-700 text-white"
                                    }`}
                                  />
                                </div>
                              </div>

                              {/* Interactive Mini Calendar */}
                              <div className={`p-2.5 rounded-xl border mb-2 ${
                                isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#0E1015] border-slate-800"
                              }`}>
                                {/* Calendar Month/Year Header */}
                                <div className="flex items-center justify-between mb-2">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      if (calMonth === 0) {
                                        setCalMonth(11);
                                        setCalYear((y) => y - 1);
                                      } else {
                                        setCalMonth((m) => m - 1);
                                      }
                                    }}
                                    className="p-1 rounded hover:bg-white/10 text-slate-400 hover:text-white cursor-pointer"
                                    title="Previous Month"
                                  >
                                    <ChevronLeft className="w-4 h-4" />
                                  </button>

                                  <span className="font-bold text-xs font-mono text-blue-400">
                                    {MONTH_NAMES[calMonth]} {calYear}
                                  </span>

                                  <button
                                    type="button"
                                    onClick={() => {
                                      if (calMonth === 11) {
                                        setCalMonth(0);
                                        setCalYear((y) => y + 1);
                                      } else {
                                        setCalMonth((m) => m + 1);
                                      }
                                    }}
                                    className="p-1 rounded hover:bg-white/10 text-slate-400 hover:text-white cursor-pointer"
                                    title="Next Month"
                                  >
                                    <ChevronRight className="w-4 h-4" />
                                  </button>
                                </div>

                                {/* Day of Week Headers */}
                                <div className="grid grid-cols-7 gap-1 text-center text-[10px] font-bold text-slate-500 mb-1">
                                  {["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"].map((day) => (
                                    <div key={day}>{day}</div>
                                  ))}
                                </div>

                                {/* Calendar Days Grid */}
                                <div className="grid grid-cols-7 gap-1 text-center text-[11px]">
                                  {Array.from({ length: new Date(calYear, calMonth, 1).getDay() }).map((_, i) => (
                                    <div key={`empty-${i}`} className="h-6" />
                                  ))}

                                  {Array.from({ length: new Date(calYear, calMonth + 1, 0).getDate() }).map((_, idx) => {
                                    const dayNum = idx + 1;
                                    const dayISO = `${calYear}-${String(calMonth + 1).padStart(2, "0")}-${String(dayNum).padStart(2, "0")}`;
                                    const isStart = pendingDateRange.start === dayISO;
                                    const isEnd = pendingDateRange.end === dayISO;
                                    const isInRange =
                                      pendingDateRange.start &&
                                      pendingDateRange.end &&
                                      dayISO > pendingDateRange.start &&
                                      dayISO < pendingDateRange.end;

                                    return (
                                      <button
                                        key={dayNum}
                                        type="button"
                                        onClick={() => {
                                          if (!pendingDateRange.start || (pendingDateRange.start && pendingDateRange.end)) {
                                            setPendingDateRange({ start: dayISO, end: "" });
                                          } else {
                                            if (dayISO < pendingDateRange.start) {
                                              setPendingDateRange({ start: dayISO, end: "" });
                                            } else {
                                              setPendingDateRange({ start: pendingDateRange.start, end: dayISO });
                                            }
                                          }
                                        }}
                                        className={`h-6 w-full rounded flex items-center justify-center font-mono transition-all text-[11px] cursor-pointer ${
                                          isStart || isEnd
                                            ? "bg-blue-600 text-white font-bold shadow-sm ring-1 ring-blue-300 scale-105 z-10"
                                            : isInRange
                                            ? isLightMode
                                              ? "bg-blue-100 text-blue-900 font-semibold"
                                              : "bg-blue-500/25 text-blue-200 font-semibold"
                                            : isLightMode
                                            ? "hover:bg-slate-200 text-slate-700"
                                            : "hover:bg-white/10 text-slate-300"
                                        }`}
                                      >
                                        {dayNum}
                                      </button>
                                    );
                                  })}
                                </div>
                              </div>
                            </div>
                          ) : (
                            /* Specific Values Checklist View */
                            <div className="border-t border-slate-700/50 pt-2 mb-2">
                              <div className="relative mb-2">
                                <Search className="w-3 h-3 absolute left-2 top-2 text-slate-400" />
                                <input
                                  type="text"
                                  value={colSearchTerm}
                                  onChange={(e) => setColSearchTerm(e.target.value)}
                                  onKeyDown={(e) => {
                                    if (e.key === "Enter") {
                                      e.preventDefault();
                                      handleApplyColumnFilter(col.key, allValues);
                                    }
                                  }}
                                  placeholder="Search values..."
                                  className={`w-full pl-7 pr-2 py-1 text-xs rounded border focus:outline-none ${
                                    isLightMode
                                      ? "bg-slate-100 border-slate-300 text-slate-800"
                                      : "bg-[#0E1015] border-white/10 text-white"
                                  }`}
                                />
                              </div>

                              {/* Select All Checkbox */}
                              <div className="flex items-center justify-between px-1 py-1 text-[11px] font-bold border-b border-slate-700/30 mb-1">
                                <label className="flex items-center gap-2 cursor-pointer select-none">
                                  <input
                                    type="checkbox"
                                    checked={isAllVisibleSelected}
                                    ref={(el) => {
                                      if (el) el.indeterminate = isIndeterminate;
                                    }}
                                    onChange={() =>
                                      handleSelectAllPendingValues(
                                        allValues,
                                        colSearchTerm.trim() ? visibleValues : undefined
                                      )
                                    }
                                    className="rounded border-slate-600 cursor-pointer accent-blue-600"
                                  />
                                  <span>
                                    {colSearchTerm.trim() ? "(Select All Search Results)" : "(Select All)"}
                                  </span>
                                </label>
                                <span className="text-[10px] text-slate-400 font-mono">
                                  {colSearchTerm.trim()
                                    ? `${visibleValues.length}/${uniqueVals.length}`
                                    : `${uniqueVals.length}`}{" "}
                                  items
                                </span>
                              </div>

                              {/* Checklist of Unique Values */}
                              <div className="max-h-36 overflow-y-auto space-y-0.5 pr-1 font-mono text-[11px]">
                                {filteredUniqueVals.map((item) => {
                                  const isChecked = activeSelectedSet.has(item.value);
                                  return (
                                    <label
                                      key={item.value}
                                      className="flex items-center justify-between px-1.5 py-1 rounded hover:bg-white/5 cursor-pointer text-xs select-none"
                                    >
                                      <div className="flex items-center gap-2 truncate pr-2">
                                        <input
                                          type="checkbox"
                                          checked={isChecked}
                                          onChange={() =>
                                            handleTogglePendingValue(item.value, allValues)
                                          }
                                          className="rounded border-slate-600 cursor-pointer shrink-0 accent-blue-600"
                                        />
                                        <span className="truncate">{item.value || "(Blank)"}</span>
                                      </div>
                                      <span className="text-[10px] text-slate-500 font-sans">({item.count})</span>
                                    </label>
                                  );
                                })}

                                {filteredUniqueVals.length === 0 && (
                                  <div className="p-2 text-center text-slate-500 text-[11px]">
                                    No matching values
                                  </div>
                                )}
                              </div>
                            </div>
                          )}

                        {/* Popover Actions */}
                        <div className="flex items-center justify-between pt-2.5 border-t border-slate-700/50 mt-2 gap-2">
                          <button
                            type="button"
                            onClick={() => handleClearColumnFilter(col.key)}
                            className="px-2 py-1 text-[11px] font-bold text-white bg-rose-800 hover:bg-rose-900 border border-rose-950 rounded transition-all cursor-pointer shadow-sm"
                          >
                            Clear Filter
                          </button>
                          <div className="flex items-center gap-1.5">
                            <button
                              type="button"
                              onClick={() => setOpenFilterCol(null)}
                              className="px-2.5 py-1 text-[11px] font-bold rounded transition-all cursor-pointer bg-slate-800 hover:bg-slate-900 text-white border border-slate-900 shadow-sm"
                            >
                              Cancel
                            </button>
                            <button
                              type="button"
                              onClick={() => handleApplyColumnFilter(col.key, allValues)}
                              className="px-3 py-1 bg-blue-700 hover:bg-blue-800 text-white border border-blue-900 rounded text-xs font-bold cursor-pointer shadow-md flex items-center gap-1 transition-all"
                            >
                              <Check className="w-3.5 h-3.5" />
                              Apply
                            </button>
                          </div>
                        </div>
                      </div>
                    </>
                  )}
                  </th>
                );
              })}
            </tr>

          </thead>

          <tbody className={`divide-y-2 ${isLightMode ? "divide-slate-500" : "divide-slate-700"}`}>
            {processedRows.map((row, idx) => {
              const isSelected = selectedRowIds.has(row.id);

              return (
                <tr
                  key={row.id}
                  id={`grid-row-${row.id}`}
                  className={`group relative transition-colors font-sans ${
                    isSelected
                      ? isLightMode
                        ? "bg-blue-50/80"
                        : "bg-blue-950/30"
                      : isLightMode
                      ? "hover:bg-slate-50"
                      : "hover:bg-white/[0.02]"
                  }`}
                >
                  {/* Row Checkbox & Hover Insert Plus Icons */}
                  <td className={`relative p-2 text-center border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"}`}>
                    <input
                      type="checkbox"
                      checked={isSelected}
                      onChange={() => handleToggleSelectRow(row.id)}
                      className="rounded border-slate-600 cursor-pointer"
                    />

                    {/* Top plus icon for row 0 (insert above row 0) */}
                    {idx === 0 && canEdit && (
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleInsertRowAt(0);
                        }}
                        className="absolute -top-2.5 -left-1 z-30 w-5 h-5 rounded-full bg-blue-600 hover:bg-blue-500 text-white opacity-0 group-hover:opacity-100 transition-all duration-150 flex items-center justify-center shadow-lg border border-white/30 cursor-pointer hover:scale-125"
                        title="Insert row above"
                      >
                        <Plus className="w-3.5 h-3.5 stroke-[3]" />
                      </button>
                    )}

                    {/* Bottom plus icon (insert between this row and next row) */}
                    {canEdit && (
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleInsertRowAfter(row.id);
                        }}
                        className="absolute -bottom-2.5 -left-1 z-30 w-5 h-5 rounded-full bg-blue-600 hover:bg-blue-500 text-white opacity-0 group-hover:opacity-100 transition-all duration-150 flex items-center justify-center shadow-lg border border-white/30 cursor-pointer hover:scale-125"
                        title="Insert row here"
                      >
                        <Plus className="w-3.5 h-3.5 stroke-[3]" />
                      </button>
                    )}
                  </td>

                  {/* Client / Operator Field */}
                  <td
                    onClick={() => {
                      if (canEdit) setActiveCell({ rowId: row.id, field: "client" });
                    }}
                    className={`p-1.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} ${
                      canEdit && activeCell?.rowId === row.id && activeCell?.field === "client"
                        ? "ring-2 ring-blue-500 rounded bg-blue-500/10"
                        : ""
                    }`}
                  >
                    <input
                      type="text"
                      readOnly={!canEdit}
                      disabled={!canEdit}
                      tabIndex={canEdit ? 0 : -1}
                      value={row.client}
                      onChange={(e) => canEdit && handleCellChange(row.id, "client", e.target.value)}
                      className={`w-full px-1.5 py-1 rounded bg-transparent border-0 font-semibold ${
                        canEdit ? "focus:ring-1 focus:ring-blue-500 cursor-text" : "cursor-default pointer-events-none focus:outline-none focus:ring-0"
                      } ${isLightMode ? "text-slate-900" : "text-white"}`}
                    />
                  </td>

                  {/* Rig Field */}
                  <td
                    onClick={() => {
                      if (canEdit) setActiveCell({ rowId: row.id, field: "rig" });
                    }}
                    className={`p-1.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} ${
                      canEdit && activeCell?.rowId === row.id && activeCell?.field === "rig"
                        ? "ring-2 ring-blue-500 rounded bg-blue-500/10"
                        : ""
                    }`}
                  >
                    <input
                      type="text"
                      readOnly={!canEdit}
                      disabled={!canEdit}
                      tabIndex={canEdit ? 0 : -1}
                      value={row.rig}
                      onChange={(e) => canEdit && handleCellChange(row.id, "rig", e.target.value)}
                      className={`w-full px-1.5 py-1 rounded bg-transparent border-0 font-medium ${
                        canEdit ? "focus:ring-1 focus:ring-blue-500 cursor-text" : "cursor-default pointer-events-none focus:outline-none focus:ring-0"
                      } ${isLightMode ? "text-slate-800" : "text-slate-200"}`}
                    />
                  </td>

                  {/* Well Field */}
                  <td
                    onClick={() => {
                      if (canEdit) setActiveCell({ rowId: row.id, field: "well" });
                    }}
                    className={`p-1.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} ${
                      canEdit && activeCell?.rowId === row.id && activeCell?.field === "well"
                        ? "ring-2 ring-blue-500 rounded bg-blue-500/10"
                        : ""
                    }`}
                  >
                    <input
                      type="text"
                      readOnly={!canEdit}
                      disabled={!canEdit}
                      tabIndex={canEdit ? 0 : -1}
                      value={row.well}
                      onChange={(e) => canEdit && handleCellChange(row.id, "well", e.target.value)}
                      className={`w-full px-1.5 py-1 rounded bg-transparent border-0 ${
                        canEdit ? "focus:ring-1 focus:ring-blue-500 cursor-text" : "cursor-default pointer-events-none focus:outline-none focus:ring-0"
                      } ${isLightMode ? "text-slate-800" : "text-slate-200"}`}
                    />
                  </td>

                  {/* Hole Size Field */}
                  <td
                    onClick={() => {
                      if (canEdit) setActiveCell({ rowId: row.id, field: "holeSize" });
                    }}
                    className={`p-1.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} ${
                      canEdit && activeCell?.rowId === row.id && activeCell?.field === "holeSize"
                        ? "ring-2 ring-blue-500 rounded bg-blue-500/10"
                        : ""
                    }`}
                  >
                    <input
                      type="text"
                      readOnly={!canEdit}
                      disabled={!canEdit}
                      tabIndex={canEdit ? 0 : -1}
                      value={row.holeSize ?? ""}
                      onChange={(e) => canEdit && handleCellChange(row.id, "holeSize", e.target.value)}
                      placeholder='e.g. 12.25"'
                      className={`w-full px-1.5 py-1 rounded bg-transparent border-0 font-mono text-xs ${
                        canEdit ? "focus:ring-1 focus:ring-blue-500 cursor-text" : "cursor-default pointer-events-none focus:outline-none focus:ring-0"
                      } ${isLightMode ? "text-slate-800 font-medium" : "text-amber-300 font-medium"}`}
                    />
                  </td>

                  {/* Qty Field */}
                  <td
                    onClick={() => {
                      if (canEdit) setActiveCell({ rowId: row.id, field: "qty" });
                    }}
                    className={`p-1.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} text-center font-mono ${
                      canEdit && activeCell?.rowId === row.id && activeCell?.field === "qty"
                        ? "ring-2 ring-blue-500 rounded bg-blue-500/10"
                        : ""
                    }`}
                  >
                    <input
                      type="number"
                      min={1}
                      readOnly={!canEdit}
                      disabled={!canEdit}
                      tabIndex={canEdit ? 0 : -1}
                      value={row.qty ?? ""}
                      onChange={(e) => canEdit && handleCellChange(row.id, "qty", e.target.value)}
                      className={`w-full px-1 py-1 text-center rounded bg-transparent border-0 font-bold ${
                        canEdit ? "focus:ring-1 focus:ring-blue-500 cursor-text" : "cursor-default pointer-events-none focus:outline-none focus:ring-0"
                      } ${isLightMode ? "text-slate-900" : "text-white"}`}
                    />
                  </td>

                  {/* Configuration & Sub-Tools */}
                  <td
                    onClick={() => {
                      if (canEdit) setActiveCell({ rowId: row.id, field: "configuration" });
                    }}
                    className={`p-1.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} font-mono text-xs ${
                      canEdit && activeCell?.rowId === row.id && activeCell?.field === "configuration"
                        ? "ring-2 ring-blue-500 rounded bg-blue-500/10"
                        : ""
                    }`}
                  >
                    <input
                      type="text"
                      readOnly={!canEdit}
                      disabled={!canEdit}
                      tabIndex={canEdit ? 0 : -1}
                      value={row.configuration}
                      onChange={(e) => canEdit && handleCellChange(row.id, "configuration", e.target.value)}
                      className={`w-full px-1.5 py-1 rounded bg-transparent border-0 text-blue-400 font-semibold ${
                        canEdit ? "focus:ring-1 focus:ring-blue-500 cursor-text" : "cursor-default pointer-events-none focus:outline-none focus:ring-0"
                      } ${isLightMode ? "text-blue-700" : "text-blue-400"}`}
                    />
                  </td>

                  {/* Footage Field */}
                  <td
                    onClick={() => {
                      if (canEdit) setActiveCell({ rowId: row.id, field: "footage" });
                    }}
                    className={`p-1.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} font-mono ${
                      canEdit && activeCell?.rowId === row.id && activeCell?.field === "footage"
                        ? "ring-2 ring-blue-500 rounded bg-blue-500/10"
                        : ""
                    }`}
                  >
                    <input
                      type="number"
                      readOnly={!canEdit}
                      disabled={!canEdit}
                      tabIndex={canEdit ? 0 : -1}
                      value={row.footage ?? ""}
                      placeholder="0"
                      onChange={(e) => canEdit && handleCellChange(row.id, "footage", e.target.value)}
                      className={`w-full px-1.5 py-1 rounded bg-transparent border-0 text-emerald-400 font-bold ${
                        canEdit ? "focus:ring-1 focus:ring-blue-500 cursor-text" : "cursor-default pointer-events-none focus:outline-none focus:ring-0"
                      } ${isLightMode ? "text-emerald-700" : "text-emerald-400"}`}
                    />
                  </td>

                  {/* Date Field */}
                  <td
                    onClick={() => {
                      if (canEdit) {
                        setActiveCell({ rowId: row.id, field: "date" });
                        handleOpenCellDatePicker(row.id, "date", row.date);
                      }
                    }}
                    className={`p-1.5 border-r-2 relative ${isLightMode ? "border-slate-500" : "border-slate-700"} font-mono cursor-pointer ${
                      canEdit && activeCell?.rowId === row.id && activeCell?.field === "date"
                        ? "ring-2 ring-blue-500 rounded bg-blue-500/10"
                        : ""
                    }`}
                  >
                    <div className="flex items-center gap-1 w-full">
                      <input
                        type="text"
                        readOnly={true}
                        disabled={!canEdit}
                        tabIndex={canEdit ? 0 : -1}
                        value={row.date}
                        placeholder="Select Date"
                        onClick={(e) => {
                          e.stopPropagation();
                          if (canEdit) {
                            setActiveCell({ rowId: row.id, field: "date" });
                            handleOpenCellDatePicker(row.id, "date", row.date);
                          }
                        }}
                        className={`w-full px-1 py-1 rounded bg-transparent border-0 text-xs font-mono select-none ${
                          canEdit ? "cursor-pointer hover:bg-blue-500/10 focus:ring-1 focus:ring-blue-500" : "cursor-default pointer-events-none focus:outline-none focus:ring-0"
                        } ${isLightMode ? "text-slate-800" : "text-slate-200"}`}
                      />

                      {canEdit && (
                        <div className="flex items-center gap-0.5 shrink-0">
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              handleOpenCellDatePicker(row.id, "date", row.date);
                            }}
                            title="Open calendar pop-up picker"
                            className={`p-1 rounded transition-all cursor-pointer ${
                              isLightMode
                                ? "text-blue-600 hover:bg-blue-100"
                                : "text-blue-400 hover:bg-blue-500/20"
                            }`}
                          >
                            <Calendar className="w-3.5 h-3.5" />
                          </button>

                          {row.date && (
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                handleCellChange(row.id, "date", "");
                              }}
                              title="Clear log date (leave blank)"
                              className="p-1 rounded text-slate-400 hover:text-red-500 hover:bg-red-500/10 cursor-pointer transition-all"
                            >
                              <X className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </div>
                      )}
                    </div>

                    {/* Interactive Calendar Pop-up Popover */}
                    {cellDatePicker?.rowId === row.id && cellDatePicker?.field === "date" && (
                      <>
                        <div
                          className="fixed inset-0 z-40 cursor-default"
                          onClick={(e) => {
                            e.stopPropagation();
                            setCellDatePicker(null);
                          }}
                        />
                        <div
                          className={`absolute right-0 top-full mt-1 z-50 p-3 rounded-xl border shadow-2xl w-64 select-none ${
                            isLightMode
                              ? "bg-white border-slate-300 text-slate-800 shadow-xl"
                              : "bg-[#16191E] border-white/15 text-white shadow-2xl"
                          }`}
                          onClick={(e) => e.stopPropagation()}
                        >
                          <div className="flex items-center justify-between mb-2 pb-1.5 border-b border-slate-700/30">
                            <span className="text-xs font-bold text-blue-500 flex items-center gap-1.5">
                              <Calendar className="w-3.5 h-3.5" /> Select Log Date
                            </span>
                            <button
                              type="button"
                              onClick={() => setCellDatePicker(null)}
                              className="text-slate-400 hover:text-white p-0.5 rounded cursor-pointer"
                            >
                              <X className="w-3.5 h-3.5" />
                            </button>
                          </div>

                          {/* Month/Year Navigation */}
                          <div className="flex items-center justify-between mb-2 px-1">
                            <button
                              type="button"
                              onClick={() => {
                                if (pickerMonth === 0) {
                                  setPickerMonth(11);
                                  setPickerYear((y) => y - 1);
                                } else {
                                  setPickerMonth((m) => m - 1);
                                }
                              }}
                              className="p-1 rounded hover:bg-slate-500/20 text-slate-400 hover:text-white cursor-pointer"
                              title="Previous Month"
                            >
                              <ChevronLeft className="w-4 h-4" />
                            </button>

                            <span className="font-bold text-xs font-mono text-blue-400">
                              {MONTH_NAMES[pickerMonth]} {pickerYear}
                            </span>

                            <button
                              type="button"
                              onClick={() => {
                                if (pickerMonth === 11) {
                                  setPickerMonth(0);
                                  setPickerYear((y) => y + 1);
                                } else {
                                  setPickerMonth((m) => m + 1);
                                }
                              }}
                              className="p-1 rounded hover:bg-slate-500/20 text-slate-400 hover:text-white cursor-pointer"
                              title="Next Month"
                            >
                              <ChevronRight className="w-4 h-4" />
                            </button>
                          </div>

                          {/* Day of Week Headers */}
                          <div className="grid grid-cols-7 gap-1 text-center text-[10px] font-bold text-slate-400 mb-1">
                            {["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"].map((day) => (
                              <div key={day}>{day}</div>
                            ))}
                          </div>

                          {/* Calendar Days Grid */}
                          <div className="grid grid-cols-7 gap-1 text-center text-[11px] mb-3">
                            {Array.from({ length: new Date(pickerYear, pickerMonth, 1).getDay() }).map((_, i) => (
                              <div key={`empty-${i}`} className="h-6" />
                            ))}

                            {Array.from({ length: new Date(pickerYear, pickerMonth + 1, 0).getDate() }).map((_, idx) => {
                              const dayNum = idx + 1;
                              const iso = `${pickerYear}-${String(pickerMonth + 1).padStart(2, "0")}-${String(dayNum).padStart(2, "0")}`;

                              return (
                                <button
                                  key={dayNum}
                                  type="button"
                                  onClick={() => {
                                    handleCellChange(row.id, "date", iso);
                                    setCellDatePicker(null);
                                  }}
                                  className={`h-6 w-full rounded flex items-center justify-center font-mono text-[11px] transition-all cursor-pointer ${
                                    isLightMode
                                      ? "hover:bg-blue-100 hover:text-blue-700 text-slate-700 font-medium"
                                      : "hover:bg-blue-600/30 hover:text-blue-200 text-slate-200 font-medium"
                                  }`}
                                >
                                  {dayNum}
                                </button>
                              );
                            })}
                          </div>

                          {/* Quick Action Footer */}
                          <div className="pt-2 border-t border-slate-700/40 flex flex-col gap-2">
                            <div className="flex items-center justify-between gap-1.5">
                              <button
                                type="button"
                                onClick={() => {
                                  const today = new Date();
                                  const iso = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
                                  handleCellChange(row.id, "date", iso);
                                  setCellDatePicker(null);
                                }}
                                className="px-2 py-1 text-[11px] font-semibold rounded bg-blue-600/20 text-blue-400 hover:bg-blue-600/30 cursor-pointer transition-all flex-1 text-center"
                              >
                                Today
                              </button>

                              <button
                                type="button"
                                onClick={() => {
                                  handleCellChange(row.id, "date", "");
                                  setCellDatePicker(null);
                                }}
                                className="px-2 py-1 text-[11px] font-semibold rounded bg-red-500/15 text-red-400 hover:bg-red-500/25 cursor-pointer transition-all flex-1 text-center"
                              >
                                Clear (Leave Blank)
                              </button>
                            </div>
                          </div>
                        </div>
                      </>
                    )}
                  </td>

                  {/* Month Field */}
                  <td
                    onClick={() => {
                      if (canEdit) setActiveCell({ rowId: row.id, field: "month" });
                    }}
                    className={`p-1.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} font-mono ${
                      canEdit && activeCell?.rowId === row.id && activeCell?.field === "month"
                        ? "ring-2 ring-blue-500 rounded bg-blue-500/10"
                        : ""
                    }`}
                  >
                    <input
                      type="text"
                      readOnly={!canEdit}
                      disabled={!canEdit}
                      tabIndex={canEdit ? 0 : -1}
                      value={formatMonthName(row.month)}
                      onChange={(e) => canEdit && handleCellChange(row.id, "month", e.target.value)}
                      className={`w-full px-1.5 py-1 rounded bg-transparent border-0 ${
                        canEdit ? "focus:ring-1 focus:ring-blue-500 cursor-text" : "cursor-default pointer-events-none focus:outline-none focus:ring-0"
                      } ${isLightMode ? "text-slate-800" : "text-slate-300"}`}
                    />
                  </td>

                  {/* Shop Date Status */}
                  <td
                    onClick={() => {
                      if (canEdit) setActiveCell({ rowId: row.id, field: "shopDate" });
                    }}
                    className={`p-1.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} ${
                      canEdit && activeCell?.rowId === row.id && activeCell?.field === "shopDate"
                        ? "ring-2 ring-blue-500 rounded bg-blue-500/10"
                        : ""
                    }`}
                  >
                    <input
                      type="text"
                      readOnly={!canEdit}
                      disabled={!canEdit}
                      tabIndex={canEdit ? 0 : -1}
                      value={row.shopDate || ""}
                      placeholder="Status/Date"
                      onChange={(e) => canEdit && handleCellChange(row.id, "shopDate", e.target.value)}
                      className={`w-full px-1.5 py-1 rounded bg-transparent border-0 text-xs ${
                        canEdit ? "focus:ring-1 focus:ring-blue-500 cursor-text" : "cursor-default pointer-events-none focus:outline-none focus:ring-0"
                      } ${isLightMode ? "text-slate-800" : "text-slate-300"}`}
                    />
                  </td>

                  {/* Comments Field */}
                  <td
                    onClick={() => {
                      if (canEdit) setActiveCell({ rowId: row.id, field: "comments" });
                    }}
                    className={`p-1.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"} ${
                      canEdit && activeCell?.rowId === row.id && activeCell?.field === "comments"
                        ? "ring-2 ring-blue-500 rounded bg-blue-500/10"
                        : ""
                    }`}
                  >
                    <input
                      type="text"
                      readOnly={!canEdit}
                      disabled={!canEdit}
                      tabIndex={canEdit ? 0 : -1}
                      value={row.comments}
                      onChange={(e) => canEdit && handleCellChange(row.id, "comments", e.target.value)}
                      className={`w-full px-1.5 py-1 rounded bg-transparent border-0 ${
                        canEdit ? "focus:ring-1 focus:ring-blue-500 cursor-text" : "cursor-default pointer-events-none focus:outline-none focus:ring-0"
                      } ${isLightMode ? "text-slate-600" : "text-slate-400"}`}
                    />
                  </td>
                </tr>
              );
            })}

            {processedRows.length === 0 && (
              <tr>
                <td colSpan={columnsList.length + 1} className="p-4">
                  <EmptyState
                    title="No records match your filter criteria"
                    description="Try resetting your search query or column filters to view all spreadsheet rows."
                    actionLabel="Reset Filters"
                    onAction={handleResetAllFilters}
                  />
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>



      {/* Floating Batch Action Dock when rows are selected */}
      {canEdit && selectedRowIds.size > 0 && (
        <div className="sticky bottom-4 z-40 mx-auto -mt-12 w-fit flex items-center gap-2.5 px-4 py-2.5 bg-slate-900/95 text-white border border-slate-700/80 rounded-2xl shadow-2xl backdrop-blur-md animate-fade-in font-sans">
          <span className="text-xs font-bold text-slate-300 pr-2 border-r border-slate-700/80">
            {selectedRowIds.size} row{selectedRowIds.size > 1 ? "s" : ""} selected
          </span>
          <button
            type="button"
            onClick={handleBulkDuplicate}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-xl bg-indigo-800 hover:bg-indigo-900 border border-indigo-950 text-white shadow-md transition-all cursor-pointer hover:scale-105"
            title="Duplicate selected row(s) right below"
            id="floating-duplicate-selected-btn"
          >
            <Copy className="w-3.5 h-3.5" />
            <span>Duplicate Right Below</span>
          </button>
          <button
            type="button"
            onClick={handleBulkDelete}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-xl bg-rose-800 hover:bg-rose-900 border border-rose-950 text-white shadow-md transition-all cursor-pointer hover:scale-105"
            title="Delete selected row(s)"
            id="floating-delete-selected-btn"
          >
            <Trash2 className="w-3.5 h-3.5" />
            <span>Delete</span>
          </button>
          <button
            type="button"
            onClick={() => setSelectedRowIds(new Set())}
            className="p-1 text-slate-300 hover:text-white rounded-lg hover:bg-slate-800 transition-all cursor-pointer ml-1"
            title="Deselect all rows"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* Paste / Import Modal */}
      {isImportModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm animate-fade-in">
          <div
            className={`w-full max-w-2xl rounded-2xl border shadow-2xl p-6 ${
              isLightMode ? "bg-white border-slate-300 text-slate-900" : "bg-[#14171F] border-white/10 text-white"
            }`}
          >
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2">
                <Upload className="w-5 h-5 text-purple-400" />
                <h3 className="text-base font-bold font-sans">Paste Tab-Separated Spreadsheet Rows</h3>
              </div>
              <button
                type="button"
                onClick={() => setIsImportModalOpen(false)}
                className="p-1 rounded-lg hover:bg-white/10 text-slate-400 hover:text-white"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <p className="text-xs text-slate-400 mb-3 font-sans">
              Copy cells directly from Excel or Google Sheets and paste them below. Expected columns: ID, Client, Rig, Well, Section, Qty, Configuration, Footage, Date, Month, Shop Status, Comments.
            </p>

            <textarea
              rows={8}
              value={pasteDataText}
              onChange={(e) => setPasteDataText(e.target.value)}
              placeholder={`1\tChevron\tRig 101\tWell 04\t6.75in\t1\tIC - PST - RST\t1200\t22-Jun-26\tJun\tDelivered\tNew run`}
              className={`w-full p-3 font-mono text-xs rounded-xl border focus:outline-none mb-4 ${
                isLightMode
                  ? "bg-slate-50 border-slate-300 text-slate-800 focus:border-purple-500"
                  : "bg-[#0E1015] border-white/10 text-slate-200 focus:border-purple-500"
              }`}
            />

            <div className="flex justify-end gap-3">
              <button
                type="button"
                onClick={() => setIsImportModalOpen(false)}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-slate-800 hover:bg-slate-900 text-white border border-slate-900 shadow-sm cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleImportPastedText}
                disabled={!pasteDataText.trim()}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-purple-800 hover:bg-purple-900 border border-purple-950 text-white shadow-md disabled:opacity-50 cursor-pointer"
              >
                Import Rows
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default React.memo(ExcelGridEditor);
