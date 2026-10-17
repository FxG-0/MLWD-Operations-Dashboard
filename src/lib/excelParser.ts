import * as XLSX from "xlsx";
import { RowData, parseAnyDate, validateRow, formatMonthName, getHoleSizeFromSection, getSectionFromHoleSize } from "../types";

export interface ValidationReport {
  isValid: boolean;
  criticalMissing: string[];
  optionalMissing: string[];
  unrecognized: string[];
}

export const SCHEMA_FIELD_LABELS: { [key: string]: string } = {
  client: "Client Name (e.g. Client, Client Name)",
  rig: "Rig Name (e.g. Rig, Rig Name)",
  well: "Well Name (e.g. Well, Well Name)",
  section: "Tool Size / Section (e.g. Section, Size)",
  holeSize: "Hole Size (e.g. Hole Size, Hole Dia)",
  qty: "Quantity (e.g. Qty, Quantity)",
  configuration: "Tool Configuration (e.g. Configuration, Config, Tools)",
  footage: "Footage (e.g. Footage, Feet, Depth)",
  date: "Log Date (e.g. Date, Log Date)",
  month: "Month",
  shopDate: "Shop Status (e.g. Shop Status, Status)",
  comments: "Comments (e.g. Comments, Remarks)"
};

export const determineDateFormat = (rawDates: any[], rawMonths?: any[]): "DD/MM" | "MM/DD" => {
  const monthsLower = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  
  let ddMmConfidence = 0;
  let mmDdConfidence = 0;

  if (rawMonths && rawMonths.length === rawDates.length) {
    for (let i = 0; i < rawDates.length; i++) {
      const dVal = rawDates[i];
      const mVal = rawMonths[i];
      if (dVal === null || dVal === undefined || mVal === null || mVal === undefined) continue;
      if (dVal instanceof Date) continue;

      const dStr = String(dVal).trim();
      const mStr = String(mVal).trim().toLowerCase();
      if (!dStr || dStr.toLowerCase() === "n/a" || !mStr || mStr === "n/a") continue;

      const monthIdx = monthsLower.findIndex(m => mStr.startsWith(m));
      if (monthIdx === -1) continue;
      const expectedMonthNum = monthIdx + 1;

      const parts = dStr.split(/[\/\-\.]/);
      if (parts.length === 3) {
        if (parts[0].length === 4) continue;
        const p1 = Number(parts[0]);
        const p2 = Number(parts[1]);
        if (!isNaN(p1) && !isNaN(p2)) {
          if (p1 === expectedMonthNum && p2 !== expectedMonthNum) {
            mmDdConfidence++;
          } else if (p2 === expectedMonthNum && p1 !== expectedMonthNum) {
            ddMmConfidence++;
          }
        }
      }
    }
  }

  if (ddMmConfidence > mmDdConfidence) return "DD/MM";
  if (mmDdConfidence > ddMmConfidence) return "MM/DD";

  let ddMmCount = 0;
  let mmDdCount = 0;
  
  const p1Values: number[] = [];
  const p2Values: number[] = [];

  for (const val of rawDates) {
    if (val === null || val === undefined) continue;
    if (val instanceof Date) continue;
    const str = String(val).trim();
    if (!str || str.toLowerCase() === "n/a") continue;
    if (/^\d+(\.\d+)?$/.test(str)) continue;

    const parts = str.split(/[\/\-\.]/);
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        continue;
      }
      const p1 = Number(parts[0]);
      const p2 = Number(parts[1]);
      if (!isNaN(p1) && !isNaN(p2)) {
        p1Values.push(p1);
        p2Values.push(p2);
        if (p1 > 12 && p2 <= 12) {
          ddMmCount++;
        } else if (p2 > 12 && p1 <= 12) {
          mmDdCount++;
        }
      }
    }
  }

  if (ddMmCount > 0 && mmDdCount === 0) return "DD/MM";
  if (mmDdCount > 0 && ddMmCount === 0) return "MM/DD";
  if (ddMmCount > 0 && mmDdCount > 0) {
    return ddMmCount >= mmDdCount ? "DD/MM" : "MM/DD";
  }

  if (p2Values.length > 0) {
    const p2AllJuneJuly = p2Values.every(v => v === 6 || v === 7);
    const p1AllJuneJuly = p1Values.every(v => v === 6 || v === 7);

    if (p2AllJuneJuly && !p1AllJuneJuly) {
      return "DD/MM";
    }
    if (p1AllJuneJuly && !p2AllJuneJuly) {
      return "MM/DD";
    }
  }

  return "DD/MM";
};

export const analyzeWorksheet = (
  worksheet: XLSX.WorkSheet,
  sheetName: string
): {
  name: string;
  rows: RowData[];
  report: ValidationReport;
  rawRowsCount: number;
  matchScore: number;
} => {
  const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
  if (rows.length === 0) {
    return {
      name: sheetName,
      rows: [],
      report: { isValid: false, criticalMissing: ["client", "rig", "well", "configuration"], optionalMissing: [], unrecognized: [] },
      rawRowsCount: 0,
      matchScore: 0
    };
  }

  const HEADER_KEYWORDS = [
    "client", "operator", "company", "rig", "well", "slot", "section", "size", "holesize", "hole",
    "qty", "quantity", "config", "configuration", "tool", "tools", "footage", 
    "feet", "depth", "date", "month", "status", "shop", "comment", "remarks"
  ];

  let headerRowIndex = 0;
  let maxMatches = 0;

  // Auto-detect header row index by searching for keywords within first 25 rows
  for (let r = 0; r < Math.min(rows.length, 25); r++) {
    const rowCells = rows[r];
    if (!Array.isArray(rowCells)) continue;
    
    let matches = 0;
    for (const cell of rowCells) {
      if (cell !== null && cell !== undefined && String(cell).trim() !== "") {
        const cleanCell = String(cell).toLowerCase().trim().replace(/[^a-z0-9]/g, "");
        if (HEADER_KEYWORDS.includes(cleanCell) || HEADER_KEYWORDS.some(k => cleanCell.includes(k))) {
          matches++;
        }
      }
    }
    
    if (matches > maxMatches) {
      maxMatches = matches;
      headerRowIndex = r;
    }
  }

  const headerRow = rows[headerRowIndex] || [];
  const colMapping: { [colIndex: number]: string } = {};

  // Build a dynamic and extremely permissive mapping of columns to expected schema fields
  headerRow.forEach((cell, colIndex) => {
    if (cell === null || cell === undefined) return;
    const clean = String(cell).toLowerCase().trim();
    const rawKey = clean.replace(/[\s_\-\/\\\(\)#]/g, "");

    // 1. Month
    if (rawKey === "month" || clean.includes("month")) {
      colMapping[colIndex] = "month";
    }
    // 2. Shop Status / Shop Date
    else if (
      clean.includes("shop") || 
      clean.includes("status") || 
      clean.includes("delivery") || 
      clean.includes("delivered") || 
      clean.includes("maint") || 
      clean.includes("ship") || 
      clean.includes("receive")
    ) {
      colMapping[colIndex] = "shopDate";
    }
    // 3. Date
    else if (
      clean.includes("date") || 
      clean.includes("day") || 
      clean.includes("logged") || 
      clean.includes("run-date") || 
      clean.includes("rundate")
    ) {
      colMapping[colIndex] = "date";
    }
    // 4. Client / Operator
    else if (
      clean.includes("client") || 
      clean.includes("operator") || 
      clean.includes("company") || 
      clean.includes("customer") || 
      clean.includes("firm")
    ) {
      colMapping[colIndex] = "client";
    }
    // 5. Rig
    else if (clean.includes("rig") || clean.includes("derrick")) {
      colMapping[colIndex] = "rig";
    }
    // 6. Hole Size (explicit matching before general section/well)
    else if (
      clean.includes("hole size") || 
      clean.includes("holesize") || 
      clean.includes("hole dia") || 
      clean.includes("hole diameter") || 
      clean.includes("bore hole") ||
      rawKey === "holesize" ||
      rawKey === "hole"
    ) {
      colMapping[colIndex] = "holeSize";
    }
    // 7. Section / Tool Size
    else if (
      clean.includes("section") || 
      clean.includes("size") || 
      clean.includes("diameter") || 
      rawKey === "section" || 
      rawKey === "size"
    ) {
      colMapping[colIndex] = "section";
    }
    // 8. Well
    else if (
      clean.includes("well") || 
      clean.includes("slot") || 
      clean.includes("bore") || 
      clean.includes("hole-id")
    ) {
      colMapping[colIndex] = "well";
    }
    // 9. Qty
    else if (
      clean.includes("qty") || 
      clean.includes("quantity") || 
      clean.includes("count") || 
      clean.includes("number of") || 
      clean.includes("no. of") || 
      clean.includes("amount")
    ) {
      colMapping[colIndex] = "qty";
    }
    // 10. Configuration
    else if (
      clean.includes("config") || 
      clean.includes("tool") || 
      clean.includes("string") || 
      clean.includes("bha") || 
      clean.includes("assembly")
    ) {
      colMapping[colIndex] = "configuration";
    }
    // 11. Footage
    else if (
      clean.includes("footage") || 
      clean.includes("feet") || 
      clean.includes("depth") || 
      clean.includes("drilled") || 
      clean.includes("ft") || 
      clean.includes("meters") || 
      clean.includes("metres") || 
      clean.includes("runft")
    ) {
      colMapping[colIndex] = "footage";
    }
    // 12. Comments
    else if (
      clean.includes("comment") || 
      clean.includes("remark") || 
      clean.includes("note") || 
      clean.includes("desc") || 
      clean.includes("info")
    ) {
      colMapping[colIndex] = "comments";
    }
  });

  // Find date column index and month column index if mapped
  let dateColIdx = -1;
  let monthColIdx = -1;
  headerRow.forEach((cell, colIndex) => {
    if (colMapping[colIndex] === "date") {
      dateColIdx = colIndex;
    }
    if (colMapping[colIndex] === "month") {
      monthColIdx = colIndex;
    }
  });

  // Collect raw date and month values from row cells to detect the date format
  const rawDates: any[] = [];
  const rawMonths: any[] = [];
  if (dateColIdx !== -1) {
    for (let r = headerRowIndex + 1; r < rows.length; r++) {
      const rowCells = rows[r];
      if (rowCells && rowCells[dateColIdx] !== undefined && rowCells[dateColIdx] !== null) {
        rawDates.push(rowCells[dateColIdx]);
        rawMonths.push(monthColIdx !== -1 ? rowCells[monthColIdx] : null);
      }
    }
  }

  const detectedFormat = determineDateFormat(rawDates, rawMonths);
  const parsedRows: RowData[] = [];

  // Parse actual record rows starting from the detected header row
  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const rowCells = rows[r];
    if (!rowCells || rowCells.length === 0) continue;
    
    const isRowEmpty = rowCells.every(
      (cell) => cell === null || cell === undefined || String(cell).trim() === ""
    );
    if (isRowEmpty) continue;

    const rawItem: any = {};
    headerRow.forEach((_, colIdx) => {
      const field = colMapping[colIdx];
      if (field) {
        rawItem[field] = rowCells[colIdx];
      }
    });

    // Standardize and map all row items with sturdy defaults
    let rawClient = rawItem.client !== undefined && rawItem.client !== null ? String(rawItem.client).trim() : "";
    if (!rawClient) {
      rawClient = "N/A";
    }

    let rawRig = rawItem.rig !== undefined && rawItem.rig !== null ? String(rawItem.rig).trim() : "";
    if (!rawRig || rawRig.toLowerCase() === "n/a") {
      rawRig = "N/A";
    }

    let rawWell = rawItem.well !== undefined && rawItem.well !== null ? String(rawItem.well).trim() : "";
    if (!rawWell || rawWell.toLowerCase() === "n/a") {
      rawWell = "N/A";
    }

    let rawSectionVal = rawItem.section !== undefined && rawItem.section !== null ? String(rawItem.section).trim() : "";
    let rawHoleSizeVal = rawItem.holeSize !== undefined && rawItem.holeSize !== null ? String(rawItem.holeSize).trim() : "";

    let rawSection = "";
    let rawHoleSize = "";

    if (rawSectionVal && rawSectionVal.toLowerCase() !== "n/a") {
      rawSection = getSectionFromHoleSize(rawSectionVal);
      rawHoleSize = (rawHoleSizeVal && rawHoleSizeVal.toLowerCase() !== "n/a") ? rawHoleSizeVal : rawSectionVal;
    } else if (rawHoleSizeVal && rawHoleSizeVal.toLowerCase() !== "n/a") {
      rawSection = getSectionFromHoleSize(rawHoleSizeVal);
      rawHoleSize = rawHoleSizeVal;
    } else {
      rawSection = "8in";
      rawHoleSize = '12.25"';
    }

    let rawQty = 1;
    if (rawItem.qty !== undefined && rawItem.qty !== null) {
      const parsedQty = parseInt(String(rawItem.qty).replace(/,/g, ""), 10);
      if (!isNaN(parsedQty)) {
        rawQty = parsedQty;
      }
    }

    let rawConfig = rawItem.configuration !== undefined && rawItem.configuration !== null ? String(rawItem.configuration).trim() : "";
    if (!rawConfig || rawConfig.toLowerCase() === "n/a") {
      rawConfig = "No Config Specified";
    }

    let rawFootage: number | null = null;
    if (rawItem.footage !== undefined && rawItem.footage !== null && String(rawItem.footage).trim() !== "" && String(rawItem.footage).toLowerCase() !== "n/a") {
      const cleanedFootageStr = String(rawItem.footage).replace(/,/g, "").trim();
      const parsedFootage = parseFloat(cleanedFootageStr);
      if (!isNaN(parsedFootage)) {
        rawFootage = parsedFootage;
      }
    }

    const rawDateVal = rawItem.date;
    let derivedMonth = rawItem.month !== undefined && rawItem.month !== null ? String(rawItem.month).trim() : "";
    derivedMonth = formatMonthName(derivedMonth);

    const parsedDateObj = parseAnyDate(rawDateVal, detectedFormat, derivedMonth);
    const formattedDate = parsedDateObj.date;

    if (!derivedMonth || derivedMonth.toLowerCase() === "n/a") {
      derivedMonth = parsedDateObj.month;
    } else {
      derivedMonth = formatMonthName(derivedMonth);
    }

    const rawShopDate = rawItem.shopDate !== undefined && rawItem.shopDate !== null ? String(rawItem.shopDate).trim() : "";
    const rawComments = rawItem.comments !== undefined && rawItem.comments !== null ? String(rawItem.comments).trim() : "";

    parsedRows.push({
      id: String(r + 1),
      client: rawClient,
      rig: rawRig,
      well: rawWell,
      section: rawSection,
      holeSize: rawHoleSize,
      qty: rawQty,
      configuration: rawConfig,
      footage: rawFootage,
      date: formattedDate,
      month: derivedMonth,
      shopDate: rawShopDate,
      comments: rawComments
    });
  }

  // Apply clean filtering to ensure we only exclude rows that are completely empty or have no useful details
  const filteredRows = parsedRows.filter(row => {
    const hasSomeData = 
      row.client !== "N/A" || 
      row.rig !== "N/A" || 
      row.well !== "N/A" || 
      (row.configuration && row.configuration !== "No Config Specified") || 
      row.comments || 
      (row.date && row.date !== "N/A");
    return hasSomeData;
  });

  // Schema validation reporting
  const criticalKeys = ["client", "rig", "well", "configuration"];
  const optionalKeys = ["section", "qty", "footage", "date", "month", "shopDate", "comments"];

  const matchedKeys = new Set(Object.values(colMapping));
  const criticalMissing = criticalKeys.filter((k) => !matchedKeys.has(k));
  const optionalMissing = optionalKeys.filter((k) => !matchedKeys.has(k));
  const unrecognized: string[] = [];

  headerRow.forEach((cell, colIndex) => {
    if (cell !== null && cell !== undefined && String(cell).trim() !== "") {
      if (!colMapping[colIndex]) {
        unrecognized.push(String(cell));
      }
    }
  });

  const isValid = criticalMissing.length === 0;

  const report: ValidationReport = {
    isValid,
    criticalMissing,
    optionalMissing,
    unrecognized
  };

  const matchScore = (criticalKeys.length - criticalMissing.length) * 10 + (optionalKeys.length - optionalMissing.length);

  return {
    name: sheetName,
    rows: filteredRows,
    report,
    rawRowsCount: rows.length,
    matchScore
  };
};

/**
 * Safely parses an ArrayBuffer or byte array into an XLSX.WorkBook.
 * Validates that the buffer is not HTML or JSON before calling XLSX.read
 * to prevent SheetJS from throwing "Invalid HTML: could not find <table>".
 */
export const safeReadExcelWorkbook = (data: ArrayBuffer | ArrayLike<number>): XLSX.WorkBook => {
  if (!data) {
    throw new Error("Empty file data.");
  }

  const byteLength = (data as ArrayBuffer).byteLength ?? (data as any).length ?? 0;
  if (byteLength < 4) {
    throw new Error("File buffer too small to be a valid spreadsheet.");
  }

  const u8 = new Uint8Array(data as any, 0, Math.min(byteLength, 200));
  const headStr = new TextDecoder().decode(u8).trim().toLowerCase();

  if (
    headStr.startsWith("<") ||
    headStr.includes("<!doctype") ||
    headStr.includes("<html") ||
    headStr.includes("<head") ||
    headStr.startsWith("{") ||
    headStr.startsWith("[")
  ) {
    throw new Error("File content is HTML or JSON, not a valid Excel spreadsheet.");
  }

  return XLSX.read(data, { type: "array" });
};
