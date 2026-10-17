import * as XLSX from "xlsx";
import { RowData, formatMonthName } from "../types";

export interface ExportExcelOptions {
  fileName?: string;
  sheetName?: string;
}

/**
 * Exports application RowData to an Excel (.xlsx) file with formatting,
 * exact original column layout, auto-filters, numeric typing, and auto column widths.
 */
export function exportToFormattedExcel(rows: RowData[], options: ExportExcelOptions = {}) {
  const fileName = options.fileName || `Operational_Logs_Export_${new Date().toISOString().slice(0, 10)}.xlsx`;
  const sheetName = options.sheetName || "Operational Logs";

  // Exact standard column layout matching uploaded Excel operational logs
  const headers = [
    "ID",
    "Client / Operator",
    "Rig",
    "Well ID",
    "Section / Size",
    "Hole Size",
    "Qty",
    "Configuration",
    "Footage (ft)",
    "Date",
    "Month",
    "Shop Date",
    "Comments"
  ];

  // Construct 2D array: Row 0 is headers, following rows are values
  const data2D: any[][] = [headers];

  rows.forEach((r, idx) => {
    const qtyNum = typeof r.qty === "number" ? r.qty : parseInt(String(r.qty || 1), 10);
    const rawFt = r.footage as any;
    const footageNum = rawFt !== null && rawFt !== undefined && rawFt !== "" ? Number(rawFt) : null;

    data2D.push([
      r.id || String(idx + 1),
      r.client || "",
      r.rig || "",
      r.well || "",
      r.section || "",
      r.holeSize || "",
      isNaN(qtyNum) ? 1 : qtyNum,
      r.configuration || "",
      footageNum !== null && !isNaN(footageNum) ? footageNum : "",
      r.date || "",
      formatMonthName(r.month) || "",
      r.shopDate || "",
      r.comments || ""
    ]);
  });

  // Convert to worksheet
  const worksheet = XLSX.utils.aoa_to_sheet(data2D);

  // Apply number formats to Qty and Footage cells
  const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:M1");

  for (let R = range.s.r + 1; R <= range.e.r; ++R) {
    // Qty is column 6 (index 6)
    const qtyCellAddr = XLSX.utils.encode_cell({ r: R, c: 6 });
    if (worksheet[qtyCellAddr] && typeof worksheet[qtyCellAddr].v === "number") {
      worksheet[qtyCellAddr].z = "#,##0";
    }

    // Footage is column 8 (index 8)
    const footageCellAddr = XLSX.utils.encode_cell({ r: R, c: 8 });
    if (worksheet[footageCellAddr] && typeof worksheet[footageCellAddr].v === "number") {
      worksheet[footageCellAddr].z = "#,##0";
    }
  }

  // Calculate auto column widths with padding
  const colWidths = headers.map((header, colIdx) => {
    let maxLen = header.length;
    for (let rowIdx = 1; rowIdx < data2D.length; rowIdx++) {
      const val = data2D[rowIdx][colIdx];
      if (val !== null && val !== undefined) {
        const len = String(val).length;
        if (len > maxLen) maxLen = len;
      }
    }

    const minWidths: Record<string, number> = {
      "ID": 8,
      "Client / Operator": 18,
      "Rig": 14,
      "Well ID": 16,
      "Section / Size": 16,
      "Hole Size": 14,
      "Qty": 10,
      "Configuration": 24,
      "Footage (ft)": 15,
      "Date": 14,
      "Month": 14,
      "Shop Date": 14,
      "Comments": 32
    };

    const minW = minWidths[header] || 12;
    return { wch: Math.max(maxLen + 4, minW) };
  });

  worksheet["!cols"] = colWidths;

  // Set Excel auto-filters across header row
  worksheet["!autofilter"] = {
    ref: XLSX.utils.encode_range({
      s: { c: 0, r: 0 },
      e: { c: headers.length - 1, r: Math.max(data2D.length - 1, 0) }
    })
  };

  // Freeze top header row for smooth scrolling
  worksheet["!views"] = [{ state: "frozen", ySplit: 1, activeCell: "A2" }];

  // Construct workbook and export
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

  XLSX.writeFile(workbook, fileName);
}
