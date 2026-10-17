import React, { useState, useMemo } from "react";
import { ChevronUp, ChevronDown, Search, ArrowRight, BookOpen, Layers, FileSpreadsheet, Download, AlertTriangle } from "lucide-react";
import * as XLSX from "xlsx";
import { RowData, Abbreviation, getHoleSizeFromSection } from "../types";
import { exportToFormattedExcel } from "../lib/excelExport";
import BorderGlow from "./BorderGlow";
import { EmptyState } from "./ui/empty-state";

interface SpreadsheetTableProps {
  data: RowData[];
  abbreviations: Abbreviation[];
  highlightedRowIds: string[];
  onAbbreviationClick: (abb: string) => void;
  isLightMode?: boolean;
}

type SortField = "id" | "client" | "rig" | "well" | "section" | "holeSize" | "qty" | "footage" | "date" | "shopDate";
type SortOrder = "asc" | "desc";

function SpreadsheetTable({
  data,
  abbreviations,
  highlightedRowIds,
  onAbbreviationClick,
  isLightMode
}: SpreadsheetTableProps) {
  const [searchTerm, setSearchTerm] = useState("");
  const [sortField, setSortField] = useState<SortField>("id");
  const [sortOrder, setSortOrder] = useState<SortOrder>("asc");
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;

  // 1. Sort Handling
  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortOrder(sortOrder === "asc" ? "desc" : "asc");
    } else {
      setSortField(field);
      setSortOrder("asc");
    }
  };

  // 2. Highlighting abbreviations in Configuration field
  const renderConfig = (config: string) => {
    if (!config) return <span className="text-slate-500">N/A</span>;

    // Split using a regex that captures word boundaries so we don't lose parentheses or dashes
    const parts = config.split(/([A-Za-z0-9]+)/);

    return (
      <div className="flex flex-wrap items-center gap-1 font-mono text-xs">
        {parts.map((part, idx) => {
          const upperToken = part.toUpperCase();
          let match = abbreviations.find(
            (a) => a.abbreviation.toUpperCase() === upperToken || a.fullName.toUpperCase().replace(/[\s\-_]/g, "") === upperToken.replace(/[\s\-_]/g, "")
          );
          const isES = upperToken.startsWith("ES") || upperToken.startsWith("EXRX") || upperToken.startsWith("EARTHSTAR");
          const isEM = upperToken.startsWith("EMBATT") || upperToken.startsWith("EM_BATT");

          if (match || isES || isEM) {
            const targetAbb = match ? match.abbreviation : isEM ? "EM BATT" : "EarthStar";
            const titleText = match 
              ? `${match.fullName}: ${match.description}` 
              : isEM 
              ? "EM Battery: Electromagnetic downhole lithium battery pack supplying high-capacity power." 
              : "EarthStar: Electromagnetic telemetry sub system.";
            return (
              <button
                key={idx}
                type="button"
                onClick={() => onAbbreviationClick(targetAbb)}
                className={`px-1.5 py-0.5 text-[11px] font-bold rounded-md border transition-all cursor-pointer inline-flex items-center ${
                  isLightMode
                    ? "bg-blue-100 text-blue-700 border-blue-300 hover:bg-blue-200 shadow-2xs"
                    : "bg-blue-900/40 text-blue-300 border-blue-700/50 hover:bg-blue-800/60 hover:text-blue-200"
                }`}
                title={titleText}
              >
                {part}
              </button>
            );
          }
          return (
            <span key={idx} className="text-slate-400 font-mono text-xs">
              {part}
            </span>
          );
        })}
      </div>
    );
  };

  // Render comments with tool badge highlighting for tools mentioned in comments
  const renderComments = (comments: string) => {
    if (!comments) return <span className="text-slate-600">-</span>;

    const parts = comments.split(/(IC[\s\-_]*\(?BM\)?|[A-Za-z0-9]+)/i);

    return (
      <span>
        {parts.map((part, idx) => {
          const upperToken = part.toUpperCase();
          const isICBM = /IC[\s\-_]*\(?BM\)?/i.test(upperToken);
          let match = abbreviations.find(
            (a) => a.abbreviation.toUpperCase() === upperToken || a.fullName.toUpperCase().replace(/[\s\-_]/g, "") === upperToken.replace(/[\s\-_]/g, "")
          );
          const isES = upperToken.startsWith("ES") || upperToken.startsWith("EXRX") || upperToken.startsWith("EARTHSTAR");
          const isEM = upperToken.startsWith("EMBATT") || upperToken.startsWith("EM_BATT");
          const isCaliper = upperToken === "CALIPER" || upperToken === "BSTCALIPER";
          const isSS = upperToken === "STRATASTAR" || upperToken === "SS" || upperToken === "SS1" || upperToken === "SS2";

          if (match || isICBM || isES || isEM || isCaliper || isSS) {
            const targetAbb = isICBM
              ? "IC BM"
              : match 
              ? match.abbreviation 
              : isEM 
              ? "EM BATT" 
              : isES 
              ? "EarthStar" 
              : isCaliper 
              ? "BST Caliper" 
              : isSS 
              ? "StrataStar" 
              : upperToken;
            const titleText = isICBM
              ? "IC BM: iCruise Bus Master (IC BM) interface and telemetry tool module."
              : match 
              ? `${match.fullName}: ${match.description}` 
              : isEM 
              ? "EM Battery: Electromagnetic downhole lithium battery pack supplying high-capacity power." 
              : isES 
              ? "EarthStar: Electromagnetic telemetry sub system."
              : isCaliper
              ? "BST Caliper: Caliper sensor logging borehole dimensions and volume calculations."
              : "StrataStar: Deep azimuthal resistivity sensor.";
            return (
              <button
                key={idx}
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  onAbbreviationClick(targetAbb);
                }}
                className={`mx-0.5 px-1.5 py-0.5 text-[10px] font-bold font-mono rounded-md border transition-all cursor-pointer inline-flex items-center ${
                  isLightMode
                    ? "bg-blue-100 text-blue-700 border-blue-300 hover:bg-blue-200 shadow-2xs"
                    : "bg-blue-900/40 text-blue-300 border-blue-700/50 hover:bg-blue-800/60 hover:text-blue-200"
                }`}
                title={titleText}
              >
                {part}
              </button>
            );
          }
          return <span key={idx}>{part}</span>;
        })}
      </span>
    );
  };

  // 3. Search and Sort Processing
  const processedData = useMemo(() => {
    // A. Filter by search
    let filtered = data;

    if (searchTerm.trim()) {
      const term = searchTerm.trim();
      const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const regex = new RegExp(`\\b${escaped}`, "i");
      filtered = filtered.filter(
        (row) =>
          regex.test(row.client) ||
          regex.test(row.rig) ||
          regex.test(row.well) ||
          regex.test(row.configuration) ||
          regex.test(row.section) ||
          (row.holeSize && regex.test(row.holeSize)) ||
          (row.comments && regex.test(row.comments))
      );
    }

    // B. Sort
    const sorted = [...filtered];
    sorted.sort((a, b) => {
      let valA: any = a[sortField];
      let valB: any = b[sortField];

      // Handle nulls in footage
      if (sortField === "footage") {
        if (valA === null) valA = sortOrder === "asc" ? Infinity : -Infinity;
        if (valB === null) valB = sortOrder === "asc" ? Infinity : -Infinity;
      }

      // String ID sorting numerical conversion if possible
      if (sortField === "id") {
        const numA = parseInt(valA) || 0;
        const numB = parseInt(valB) || 0;
        return sortOrder === "asc" ? numA - numB : numB - numA;
      }

      if (typeof valA === "string") {
        return sortOrder === "asc" ? valA.localeCompare(valB) : valB.localeCompare(valA);
      } else {
        return sortOrder === "asc" ? valA - valB : valB - valA;
      }
    });

    return sorted;
  }, [data, searchTerm, sortField, sortOrder]);

  // 4. Pagination
  const totalPages = Math.ceil(processedData.length / itemsPerPage) || 1;
  const paginatedData = useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return processedData.slice(startIndex, startIndex + itemsPerPage);
  }, [processedData, currentPage]);

  const SortIcon = ({ field }: { field: SortField }) => {
    if (sortField !== field) return null;
    return sortOrder === "asc" ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />;
  };

  // Export active filtered log data to Excel (.xlsx) file
  const handleExport = () => {
    if (processedData.length === 0) return;
    exportToFormattedExcel(processedData, {
      fileName: `Operational_Logs_Export_${new Date().toISOString().split("T")[0]}.xlsx`,
      sheetName: "Operational Log"
    });
  };

  return (
    <BorderGlow
      borderRadius={16}
      backgroundColor={isLightMode ? "#FFFFFF" : "#0F1420"}
      glowColor={isLightMode ? "220 80 50" : "210 100 70"}
      glowIntensity={isLightMode ? 0.6 : 1.2}
      edgeSensitivity={25}
      className="w-full"
    >
      <div className="overflow-hidden w-full" id="spreadsheet-table-component">
      {/* Search Header */}
      <div className={`p-4 border-b flex flex-col md:flex-row items-center justify-between gap-3 ${
        isLightMode ? "bg-slate-50/80 border-slate-200 text-slate-800" : "bg-[#121316] border-white/5 text-[#F5F7FA]"
      }`}>
        <div className="flex items-center gap-2.5">
          <Layers className="w-4 h-4 text-[#3B82F6]" />
          <h3 className={`text-sm font-semibold font-sans ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"}`}>Operational Log Sheet</h3>
          <span className={`text-[11px] font-mono px-2 py-0.5 rounded-md border ${
            isLightMode ? "bg-white text-slate-700 border-slate-200 shadow-2xs" : "bg-[#1A1D21] text-[#9AA0A6] border-white/5"
          }`}>
            {processedData.length} records found
          </span>
        </div>

        <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2 w-full md:w-auto">
          <div className="relative w-full sm:w-64">
            <Search className={`absolute left-3 top-2.5 w-3.5 h-3.5 ${isLightMode ? "text-slate-400" : "text-[#9AA0A6]"}`} />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value);
                setCurrentPage(1); // Reset page on search
              }}
              placeholder="Search within columns..."
              className={`w-full h-9 pl-9 pr-3 text-[13px] rounded-lg focus:outline-none transition-all font-sans ${
                isLightMode
                  ? "bg-white border border-slate-300 text-slate-900 focus:border-blue-500 shadow-2xs"
                  : "bg-[#1A1D21] border border-white/10 text-[#E0E0E0] focus:border-[#3B82F6]"
              }`}
            />
          </div>

          <button
            type="button"
            onClick={handleExport}
            disabled={processedData.length === 0}
            className={`inline-flex items-center justify-center gap-1.5 h-9 px-4 text-xs font-semibold rounded-lg disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer font-sans ${
              isLightMode
                ? "bg-white text-blue-600 border border-slate-200 hover:bg-slate-50 shadow-2xs"
                : "bg-[#1A1D21] text-[#3B82F6] border border-white/10 hover:bg-[#3B82F6]/10 hover:text-[#3B82F6]"
            }`}
            title="Export currently filtered list to formatted Excel (.xlsx)"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>Export Excel</span>
            <Download className="w-3 h-3 text-[#3B82F6]" />
          </button>
        </div>
      </div>

      {/* Table Container */}
      <div className={`overflow-x-auto rounded-xl border-2 ${isLightMode ? "border-slate-500 bg-white" : "border-slate-700 bg-[#0F1420]"}`}>
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className={`border-b-2 text-xs font-semibold uppercase tracking-wider font-sans ${
              isLightMode
                ? "bg-slate-200 border-slate-500 text-slate-900"
                : "bg-[#0F1420] border-slate-700 text-[#9AA0A6]"
            }`}>
              <th className={`py-2.5 px-2.5 pl-4 cursor-pointer border-r-2 ${isLightMode ? "border-slate-500 hover:bg-slate-300/60" : "border-slate-700 hover:bg-white/[0.03]"}`} onClick={() => handleSort("id")}>
                <div className="flex items-center gap-1 font-mono">Row ID <SortIcon field="id" /></div>
              </th>
              <th className={`py-2.5 px-2.5 cursor-pointer border-r-2 ${isLightMode ? "border-slate-500 hover:bg-slate-300/60" : "border-slate-700 hover:bg-white/[0.03]"}`} onClick={() => handleSort("client")}>
                <div className="flex items-center gap-1">Client <SortIcon field="client" /></div>
              </th>
              <th className={`py-2.5 px-2.5 cursor-pointer border-r-2 ${isLightMode ? "border-slate-500 hover:bg-slate-300/60" : "border-slate-700 hover:bg-white/[0.03]"}`} onClick={() => handleSort("rig")}>
                <div className="flex items-center gap-1">Rig <SortIcon field="rig" /></div>
              </th>
              <th className={`py-2.5 px-2.5 cursor-pointer border-r-2 ${isLightMode ? "border-slate-500 hover:bg-slate-300/60" : "border-slate-700 hover:bg-white/[0.03]"}`} onClick={() => handleSort("well")}>
                <div className="flex items-center gap-1">Well <SortIcon field="well" /></div>
              </th>
              <th className={`py-2.5 px-2.5 cursor-pointer border-r-2 ${isLightMode ? "border-slate-500 hover:bg-slate-300/60" : "border-slate-700 hover:bg-white/[0.03]"}`} onClick={() => handleSort("section")}>
                <div className="flex items-center gap-1">Tool Size <SortIcon field="section" /></div>
              </th>
              <th className={`py-2.5 px-2.5 text-center cursor-pointer border-r-2 ${isLightMode ? "border-slate-500 hover:bg-slate-300/60" : "border-slate-700 hover:bg-white/[0.03]"}`} onClick={() => handleSort("qty")}>
                <div className="flex items-center justify-center gap-1">Qty <SortIcon field="qty" /></div>
              </th>
              <th className={`py-2.5 px-2.5 normal-case font-semibold font-sans text-xs tracking-tight border-r-2 ${isLightMode ? "border-slate-500 text-slate-900" : "border-slate-700 text-[#F5F7FA]"}`}>Configuration & Sub-Tools (Interactive Tags)</th>
              <th className={`py-2.5 px-2.5 text-right cursor-pointer border-r-2 ${isLightMode ? "border-slate-500 hover:bg-slate-300/60" : "border-slate-700 hover:bg-white/[0.03]"}`} onClick={() => handleSort("footage")}>
                <div className="flex items-center justify-end gap-1">Footage (ft) <SortIcon field="footage" /></div>
              </th>
              <th className={`py-2.5 px-2.5 cursor-pointer border-r-2 ${isLightMode ? "border-slate-500 hover:bg-slate-300/60" : "border-slate-700 hover:bg-white/[0.03]"}`} onClick={() => handleSort("date")}>
                <div className="flex items-center gap-1">Date <SortIcon field="date" /></div>
              </th>
              <th className={`py-2.5 px-2.5 cursor-pointer border-r-2 ${isLightMode ? "border-slate-500 hover:bg-slate-300/60" : "border-slate-700 hover:bg-white/[0.03]"}`} onClick={() => handleSort("shopDate")}>
                <div className="flex items-center gap-1">Shop Status <SortIcon field="shopDate" /></div>
              </th>
              <th className="py-2.5 px-2.5 pr-4">Comments</th>
            </tr>
          </thead>
          <tbody className={`divide-y-2 text-[13px] font-normal ${
            isLightMode ? "divide-slate-500 text-slate-800" : "divide-slate-700 text-[#E0E0E0]"
          }`}>
            {paginatedData.length === 0 ? (
              <tr>
                <td colSpan={11} className="p-4">
                  <EmptyState
                    title="No log records match search or filter"
                    description="Try searching with a different term or clearing active filters."
                    actionLabel="Clear Search"
                    onAction={() => setSearchTerm("")}
                  />
                </td>
              </tr>
            ) : (
              paginatedData.map((row) => {
                const isRowHighlighted = highlightedRowIds.includes(row.id);
                return (
                  <tr
                    key={row.id}
                    className={`transition-colors ${
                      isLightMode
                        ? isRowHighlighted
                          ? "bg-amber-500/10 border-l-4 border-l-amber-500 font-semibold"
                          : "hover:bg-slate-100/80 even:bg-white odd:bg-slate-50/50"
                        : isRowHighlighted
                          ? "bg-amber-500/5 border-l-4 border-l-amber-500 font-semibold"
                          : "hover:bg-white/[0.03] bg-transparent"
                    }`}
                  >
                    <td className={`py-2 px-2.5 pl-4 font-mono font-semibold border-r-2 ${isLightMode ? "border-slate-500 text-slate-500" : "border-slate-700 text-[#9AA0A6]"}`}>
                      <span>{row.id}</span>
                    </td>
                    <td className={`py-2 px-2.5 font-medium font-sans border-r-2 ${isLightMode ? "border-slate-500 text-slate-900" : "border-slate-700 text-[#E0E0E0]"}`}>{row.client}</td>
                    <td className={`py-2 px-2.5 font-mono text-[11px] border-r-2 ${isLightMode ? "border-slate-500 text-slate-600" : "border-slate-700 text-[#9AA0A6]"}`}>{row.rig}</td>
                    <td className={`py-2 px-2.5 font-mono text-[11px] border-r-2 ${isLightMode ? "border-slate-500 text-slate-600" : "border-slate-700 text-[#9AA0A6]"}`}>{row.well}</td>
                    <td className={`py-2 px-2.5 font-mono border-r-2 ${isLightMode ? "border-slate-500 text-slate-800 font-medium" : "border-slate-700 text-[#E0E0E0]"}`}>{row.section}</td>
                    <td className={`py-2 px-2.5 text-center font-mono font-bold border-r-2 ${isLightMode ? "border-slate-500 text-slate-900" : "border-slate-700 text-[#E0E0E0]"}`}>{row.qty}</td>
                    <td className={`py-2 px-2.5 max-w-[320px] border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"}`}>{renderConfig(row.configuration)}</td>
                    <td className={`py-2 px-2.5 text-right font-mono font-bold border-r-2 ${isLightMode ? "border-slate-500 text-slate-900" : "border-slate-700 text-[#F5F7FA]"}`}>
                      {row.footage !== null ? row.footage.toLocaleString() : <span className={isLightMode ? "text-slate-400" : "text-[#9AA0A6]"}>-</span>}
                    </td>
                    <td className={`py-2 px-2.5 font-mono text-[11px] whitespace-nowrap border-r-2 ${isLightMode ? "border-slate-500 text-slate-600" : "border-slate-700 text-[#9AA0A6]"}`}>{row.date}</td>
                    <td className={`py-2 px-2.5 border-r-2 ${isLightMode ? "border-slate-500" : "border-slate-700"}`}>
                      {row.shopDate === "Delivered" ? (
                        <span className={`font-semibold text-xs ${isLightMode ? "text-emerald-700 font-bold" : "text-[#10B981]"}`}>Delivered</span>
                      ) : row.shopDate ? (
                        <span className={`font-semibold text-xs ${isLightMode ? "text-rose-700 font-bold" : "text-[#EF4444]"}`}>{row.shopDate}</span>
                      ) : (
                        <span className={`text-xs ${isLightMode ? "text-slate-400" : "text-[#9AA0A6]"}`}>Pending</span>
                      )}
                    </td>
                    <td className={`py-2 px-2.5 pr-4 max-w-[200px] truncate font-sans ${isLightMode ? "text-slate-600" : "text-[#9AA0A6]"}`} title={row.comments}>
                      {renderComments(row.comments)}
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination Footer */}
      <div className={`p-4 border-t flex items-center justify-between text-xs ${
        isLightMode ? "bg-slate-50/80 border-slate-200 text-slate-600" : "bg-[#121316] border-white/5 text-[#9AA0A6]"
      }`}>
        <span className="font-sans">
          Showing <span className={`font-bold ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"}`}>{paginatedData.length}</span> of{" "}
          <span className={`font-bold ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"}`}>{processedData.length}</span> entries
          {highlightedRowIds.length > 0 && (
            <span className="ml-2 text-amber-600 bg-amber-50 px-2.5 py-0.5 rounded-md border border-amber-200 font-medium">
              {highlightedRowIds.length} highlighted
            </span>
          )}
        </span>

        <div className="flex items-center gap-1.5 font-sans">
          <button
            onClick={() => setCurrentPage(Math.max(1, currentPage - 1))}
            disabled={currentPage === 1}
            className={`px-3 h-8 flex items-center justify-center rounded-lg border text-xs font-semibold disabled:opacity-30 cursor-pointer transition-all ${
              isLightMode
                ? "border-slate-200 bg-white text-slate-700 hover:bg-slate-100 shadow-2xs"
                : "border-white/10 bg-[#1A1D21] text-[#E0E0E0] hover:text-[#F5F7FA]"
            }`}
          >
            Prev
          </button>
          <span className={`font-mono ${isLightMode ? "text-slate-700 font-semibold" : "text-[#9AA0A6]"}`}>
            {currentPage} / {totalPages}
          </span>
          <button
            onClick={() => setCurrentPage(Math.min(totalPages, currentPage + 1))}
            disabled={currentPage === totalPages}
            className={`px-3 h-8 flex items-center justify-center rounded-lg border text-xs font-semibold disabled:opacity-30 cursor-pointer transition-all ${
              isLightMode
                ? "border-slate-200 bg-white text-slate-700 hover:bg-slate-100 shadow-2xs"
                : "border-white/10 bg-[#1A1D21] text-[#E0E0E0] hover:text-[#F5F7FA]"
            }`}
          >
            Next
          </button>
        </div>
      </div>
    </div>
  </BorderGlow>
  );
}

export default React.memo(SpreadsheetTable);
