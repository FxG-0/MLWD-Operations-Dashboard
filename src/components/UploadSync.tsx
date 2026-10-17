import React, { useState, useRef, useEffect } from "react";
import { motion } from "motion/react";
import * as XLSX from "xlsx";
import { Upload, FileSpreadsheet, RefreshCw, AlertCircle, CheckCircle2, AlertTriangle, Check, Eye, Info, LogOut, Trash2, Cloud } from "lucide-react";
import { RowData, parseAnyDate, validateRow } from "../types";
import BorderGlow from "./BorderGlow";
import { ValidationReport, SCHEMA_FIELD_LABELS, analyzeWorksheet, safeReadExcelWorkbook } from "../lib/excelParser";
import { msSignIn, msSignOut, searchOneDriveExcelFiles, fetchOneDriveFileContent, getAccessToken, fetchExcelData, convertDispatchRowsToRowData } from "../lib/msAuth";

interface UploadSyncProps {
  onDataLoaded: (newData: RowData[]) => void;
  onRestoreDefault: () => void;
  onClearAllData: () => void;
  dataLength: number;
  isLightMode?: boolean;
}

export default function UploadSync({
  onDataLoaded,
  onRestoreDefault,
  onClearAllData,
  dataLength,
  isLightMode,
}: UploadSyncProps) {
  const [activeSubTab, setActiveSubTab] = useState<"file" | "onedrive">("file");

  // Microsoft Live Sync States
  const [msUser, setMsUser] = useState<any>(() => {
    const saved = localStorage.getItem("ms_user");
    return saved ? JSON.parse(saved) : null;
  });
  const [msToken, setMsToken] = useState<string | null>(() => localStorage.getItem("ms_token"));
  const [msFileId, setMsFileId] = useState<string>(() => localStorage.getItem("ms_file_id") || "");
  const [msFileName, setMsFileName] = useState<string>(() => localStorage.getItem("ms_file_name") || "");
  const [msLastSynced, setMsLastSynced] = useState<string>(() => localStorage.getItem("ms_last_synced") || "");
  const [msAutoSync, setMsAutoSync] = useState<boolean>(() => localStorage.getItem("ms_auto_sync") === "true");
  const [isSyncingMs, setIsSyncingMs] = useState(false);
  const [msSyncError, setMsSyncError] = useState("");
  const [oneDriveFiles, setOneDriveFiles] = useState<any[]>([]);
  const [isLoadingOneDrive, setIsLoadingOneDrive] = useState(false);

  // Auto-scan OneDrive files when entering the Microsoft sub tab
  useEffect(() => {
    if (activeSubTab === "onedrive" && msToken) {
      setIsLoadingOneDrive(true);
      searchOneDriveExcelFiles(msToken)
        .then((files) => {
          setOneDriveFiles(files);
          setIsLoadingOneDrive(false);
        })
        .catch((err) => {
          console.error("Failed to fetch OneDrive files:", err);
          setIsLoadingOneDrive(false);
        });
    }
  }, [activeSubTab, msToken]);

  const parseAndLoadOneDriveFile = async (accessToken: string, fileId: string, fileName: string) => {
    setIsSyncingMs(true);
    setMsSyncError("");
    try {
      let activeToken = accessToken;
      try {
        const freshToken = await getAccessToken();
        if (freshToken) {
          activeToken = freshToken;
          setMsToken(freshToken);
          localStorage.setItem("ms_token", freshToken);
        }
      } catch (tokenErr) {
        console.warn("Could not silently refresh MS token, proceeding with current token:", tokenErr);
      }

      const dispatchRows = await fetchExcelData();
      if (!dispatchRows || dispatchRows.length === 0) {
        throw new Error("The Excel table returned no rows or could not be loaded.");
      }

      const formattedRows = convertDispatchRowsToRowData(dispatchRows);

      // Load data
      onDataLoaded(formattedRows);
      
      // Update local states
      const timestamp = new Date().toLocaleTimeString();
      setMsLastSynced(timestamp);
      localStorage.setItem("ms_last_synced", timestamp);
      
      setStatus({
        type: "success",
        message: `Successfully synced ${formattedRows.length} rows from Microsoft Graph table!`
      });
      setTimeout(() => setStatus({ type: "idle", message: "" }), 4000);
    } catch (err: any) {
      console.error("OneDrive sync error:", err);
      setMsSyncError(err.message || "Failed to sync OneDrive file.");
    } finally {
      setIsSyncingMs(false);
    }
  };

  // Microsoft Auto-Sync 60s interval
  useEffect(() => {
    if (!msAutoSync || !msFileId || !msToken) return;

    // Sync immediately on mount / settings change
    parseAndLoadOneDriveFile(msToken, msFileId, msFileName);

    const interval = setInterval(() => {
      parseAndLoadOneDriveFile(msToken, msFileId, msFileName);
    }, 60000);

    return () => clearInterval(interval);
  }, [msAutoSync, msFileId, msToken]);

  const handleMsSignIn = async () => {
    try {
      const res = await msSignIn();
      if (res) {
        setMsUser(res.account);
        setMsToken(res.accessToken);
        localStorage.setItem("ms_user", JSON.stringify(res.account));
        localStorage.setItem("ms_token", res.accessToken);
        // Automatically fetch OneDrive Excel files after sign in
        setIsLoadingOneDrive(true);
        searchOneDriveExcelFiles(res.accessToken)
          .then((files) => {
            setOneDriveFiles(files);
            setIsLoadingOneDrive(false);
          })
          .catch((err) => {
            console.error("Failed to fetch OneDrive files:", err);
            setIsLoadingOneDrive(false);
          });
      }
    } catch (e: any) {
      console.error("Microsoft sign in failed", e);
      setMsSyncError(e.message || "Failed to sign in with Microsoft.");
    }
  };

  const handleMsSignOut = async () => {
    try {
      await msSignOut();
    } catch (e) {
      console.error("Sign out error", e);
    } finally {
      setMsUser(null);
      setMsToken(null);
      setMsFileId("");
      setMsFileName("");
      setMsLastSynced("");
      setOneDriveFiles([]);
      localStorage.removeItem("ms_user");
      localStorage.removeItem("ms_token");
      localStorage.removeItem("ms_file_id");
      localStorage.removeItem("ms_file_name");
      localStorage.removeItem("ms_last_synced");
    }
  };

  const [isDragging, setIsDragging] = useState(false);
  const [status, setStatus] = useState<{ type: "idle" | "success" | "error"; message: string }>({
    type: "idle",
    message: ""
  });
  const [pendingSync, setPendingSync] = useState<{
    fileName: string;
    sheets: {
      name: string;
      rows: RowData[];
      report: ValidationReport;
      rawRowsCount: number;
      matchScore: number;
    }[];
    selectedSheetIndex: number;
  } | null>(null);
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFile = (file: File) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result as ArrayBuffer;
        if (!data) throw new Error("Could not read file data");

        const workbook = safeReadExcelWorkbook(data);
        
        const sheets = workbook.SheetNames.map(name => {
          return analyzeWorksheet(workbook.Sheets[name], name);
        }).filter(s => s.rawRowsCount > 0);

        if (sheets.length === 0) {
          throw new Error("The Excel file appears to be empty or has no usable rows.");
        }

        // Sort sheets so that the sheet with highest match score comes first
        sheets.sort((a, b) => b.matchScore - a.matchScore);

        const bestSheet = sheets[0];

        // Auto-load the best sheet's rows immediately so the user gets instant visual feedback on the main dashboard
        onDataLoaded(bestSheet.rows);

        if (sheets.length === 1 && bestSheet.report.isValid) {
          setStatus({
            type: "success",
            message: `Successfully loaded ${bestSheet.rows.length} rows from sheet "${bestSheet.name}"!`
          });
          setPendingSync(null);
          setTimeout(() => setStatus({ type: "idle", message: "" }), 4000);
        } else {
          // If multiple sheets exist or there are missing fields, ask for verification but pre-load the best sheet
          setPendingSync({
            fileName: file.name,
            sheets,
            selectedSheetIndex: 0
          });
          setStatus({
            type: "success",
            message: `Loaded ${bestSheet.rows.length} rows from sheet "${bestSheet.name}". Please review column mapping below if needed.`
          });
        }
      } catch (err: any) {
        setStatus({
          type: "error",
          message: err.message || "Failed to parse Excel file. Ensure it is a valid spreadsheet."
        });
      }
    };

    reader.onerror = () => {
      setStatus({ type: "error", message: "File reading error." });
    };

    reader.readAsArrayBuffer(file);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file);
  };

  const handleDownloadTemplate = (e: React.MouseEvent) => {
    e.stopPropagation(); // Avoid triggering file chooser dialog

    const templateData = [
      {
        "Client": "Petronas",
        "Rig": "Rig-101",
        "Well": "P-45",
        "Section": "12.25\"",
        "Qty": 1,
        "Configuration": "IC-LCS-LST-iStar-BST-ILS",
        "Footage": 2450,
        "Date": "02-Jul-26",
        "Month": "Jul",
        "Status": "Delivered",
        "Comments": "Standard drilling run setup"
      },
      {
        "Client": "ADCO",
        "Rig": "AD-114",
        "Well": "ADW-9",
        "Section": "8.5\"",
        "Qty": 1,
        "Configuration": "ILS-STS-UCS-AFR-GT-P4M",
        "Footage": 1050,
        "Date": "29-Jun-26",
        "Month": "Jun",
        "Status": "Pending",
        "Comments": "High pressure high temperature reservoir zone"
      }
    ];

    const worksheet = XLSX.utils.json_to_sheet(templateData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Log Template");

    worksheet["!cols"] = [
      { wch: 15 }, // Client
      { wch: 12 }, // Rig
      { wch: 12 }, // Well
      { wch: 12 }, // Section
      { wch: 8 },  // Qty
      { wch: 28 }, // Configuration
      { wch: 12 }, // Footage
      { wch: 12 }, // Date
      { wch: 8 },  // Month
      { wch: 12 }, // Status
      { wch: 35 }  // Comments
    ];

    XLSX.writeFile(workbook, "rig_tool_log_template.xlsx");
  };

  return (
    <BorderGlow
      borderRadius={16}
      backgroundColor={isLightMode ? "#FFFFFF" : "#0F1420"}
      glowColor={isLightMode ? "220 80 50" : "210 100 70"}
      glowIntensity={isLightMode ? 0.6 : 1.2}
      edgeSensitivity={25}
      className="w-full h-full flex flex-col"
    >
      <div className="p-5 flex flex-col justify-between h-full" id="upload-sync-container">
        <div className="w-full">
          <div className="flex items-start justify-between gap-3 mb-4">
            <div>
              <h2 className={`text-sm font-semibold font-sans flex items-center gap-2 ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"}`}>
                <FileSpreadsheet className="w-4 h-4 text-[#3B82F6]" />
                Excel Upload & Live Sync
              </h2>
              <p className={`text-xs mt-1 leading-relaxed font-sans ${isLightMode ? "text-slate-600" : "text-[#9AA0A6]"}`}>
                Manage data feed using a local Excel file upload, or connect live with OneDrive.
              </p>
            </div>
            <span className={`text-[11px] font-mono font-semibold px-2 py-0.5 rounded border shrink-0 ${
              isLightMode ? "bg-slate-100 text-slate-700 border-slate-300" : "bg-[#1A1D21] text-[#9AA0A6] border-white/5"
            }`}>
              {dataLength} rows
            </span>
          </div>

          {/* Sub tabs list */}
          <div className={`flex p-1 rounded-xl mb-4 gap-1 border ${
            isLightMode ? "bg-slate-100 border-slate-200" : "border-slate-800/80 bg-slate-950/40"
          }`}>
            <button
              onClick={() => setActiveSubTab("file")}
              className={`flex-1 py-1.5 text-[11px] sm:text-xs font-bold rounded-lg transition-all cursor-pointer flex items-center justify-center gap-1.5 ${
                activeSubTab === "file"
                  ? "bg-blue-600 text-white shadow-sm"
                  : isLightMode ? "text-slate-600 hover:text-slate-900 hover:bg-slate-200/60" : "text-slate-400 hover:text-slate-200 hover:bg-slate-800/40"
              }`}
            >
              <Upload className="w-3.5 h-3.5" />
              Local File Upload
            </button>
            <button
              onClick={() => setActiveSubTab("onedrive")}
              className={`flex-1 py-1.5 text-[11px] sm:text-xs font-bold rounded-lg transition-all cursor-pointer flex items-center justify-center gap-1.5 ${
                activeSubTab === "onedrive"
                  ? "bg-cyan-600 text-white shadow-sm"
                  : isLightMode ? "text-slate-600 hover:text-slate-900 hover:bg-slate-200/60" : "text-slate-400 hover:text-slate-200 hover:bg-slate-800/40"
              }`}
            >
              <Cloud className="w-3.5 h-3.5" />
              OneDrive Live Sync
            </button>
          </div>

          {/* Active Tab: Local File Upload */}
          {activeSubTab === "file" && (
            <div className="space-y-4">
              <p className="text-[11px] text-slate-400 leading-relaxed font-sans">
                Drag-and-drop your customized operations report to trigger permissive mapping. Get our{" "}
                <button
                  type="button"
                  onClick={handleDownloadTemplate}
                  className="text-blue-400 hover:text-blue-300 underline font-semibold focus:outline-none cursor-pointer inline"
                >
                  Excel Template (.xlsx)
                </button>
              </p>

        {pendingSync ? (() => {
          const currentSheet = pendingSync.sheets[pendingSync.selectedSheetIndex];
          return (
            <motion.div
              initial={{ opacity: 0, scale: 0.98 }}
              animate={{ opacity: 1, scale: 1 }}
              className="mt-4 bg-[#151B2A] border border-blue-900/40 rounded-xl p-4 space-y-3"
              id="schema-validation-review"
            >
              <div className="flex items-start gap-2.5">
                <AlertTriangle className="w-4 h-4 text-amber-400 mt-0.5 shrink-0" />
                <div>
                  <h3 className="text-xs font-bold text-white">Verification Required</h3>
                  <p className="text-[11px] text-slate-400 mt-0.5">
                    <span className="font-mono text-blue-400 font-semibold">{pendingSync.fileName}</span> has been analyzed.
                  </p>
                </div>
              </div>

              {pendingSync.sheets.length > 1 && (
                <div className="space-y-1.5 p-2.5 bg-slate-900/40 border border-slate-800 rounded-lg">
                  <label className="text-[10px] font-bold text-slate-400 block uppercase tracking-wider font-mono">
                    Select Worksheet Tab:
                  </label>
                  <select
                    value={pendingSync.selectedSheetIndex}
                    onChange={(e) => {
                      const idx = Number(e.target.value);
                      setPendingSync(prev => prev ? { ...prev, selectedSheetIndex: idx } : null);
                    }}
                    className="w-full text-xs bg-[#10141D] text-white border border-slate-800 rounded px-2.5 py-1.5 focus:outline-none focus:border-blue-500 cursor-pointer"
                  >
                    {pendingSync.sheets.map((s, idx) => (
                      <option key={idx} value={idx}>
                        {s.name} ({s.rows.length} rows, score {s.matchScore})
                      </option>
                    ))}
                  </select>
                </div>
              )}

              <div className="space-y-2 max-h-[160px] overflow-y-auto pr-1">
                {currentSheet.report.criticalMissing.length > 0 && (
                  <div className="p-2.5 bg-rose-950/20 rounded-lg border border-rose-900/30 text-[11px]">
                    <h4 className="font-bold text-rose-400 mb-1">Missing Columns ({currentSheet.report.criticalMissing.length})</h4>
                    <ul className="list-disc pl-3 text-rose-300/90 space-y-0.5 font-mono text-[10px]">
                      {currentSheet.report.criticalMissing.map((f) => (
                        <li key={f}>{SCHEMA_FIELD_LABELS[f] || f}</li>
                      ))}
                    </ul>
                  </div>
                )}

                {currentSheet.report.optionalMissing.length > 0 && (
                  <div className="p-2.5 bg-amber-950/20 rounded-lg border border-amber-900/30 text-[11px]">
                    <h4 className="font-bold text-amber-400 mb-1">Optional Fallbacks ({currentSheet.report.optionalMissing.length})</h4>
                    <ul className="list-disc pl-3 text-amber-300/90 space-y-0.5 font-mono text-[10px]">
                      {currentSheet.report.optionalMissing.map((f) => (
                        <li key={f}>{SCHEMA_FIELD_LABELS[f] || f}</li>
                      ))}
                    </ul>
                  </div>
                )}

                {currentSheet.report.unrecognized.length > 0 ? (
                  <div className="p-2.5 bg-[#0F1420] rounded-lg border border-slate-800 text-[11px]">
                    <h4 className="font-bold text-slate-300 mb-1">Ignored Columns ({currentSheet.report.unrecognized.length})</h4>
                    <div className="flex flex-wrap gap-1 mt-1">
                      {currentSheet.report.unrecognized.map((col) => (
                        <span key={col} className="text-[9px] font-mono bg-[#151B2A] text-slate-400 px-1.5 py-0.5 rounded border border-slate-800">
                          {col}
                        </span>
                      ))}
                    </div>
                  </div>
                ) : (
                  <div className="p-2 bg-emerald-950/10 rounded-lg border border-emerald-900/20 text-[10px] text-emerald-400 flex items-center gap-1.5 font-sans">
                    <Check className="w-3 h-3" /> Clean Column Mappings Match
                  </div>
                )}

                {/* Minimal preview block */}
                <div className="p-2.5 bg-[#0F1420] rounded-lg border border-slate-800 text-[10px]">
                  <h4 className="font-bold text-slate-300 mb-1 flex items-center gap-1">
                    <Eye className="w-3 h-3 text-blue-400" /> Data Preview (First 2 Rows)
                  </h4>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-slate-400 border-collapse">
                      <tbody>
                        {currentSheet.rows.slice(0, 2).map((row, idx) => (
                          <tr key={idx} className="border-b border-slate-800/40 font-mono">
                            <td className="py-0.5 pr-2 text-white font-sans truncate max-w-[80px]">{row.client}</td>
                            <td className="py-0.5 pr-2 truncate max-w-[80px]">{row.rig}</td>
                            <td className="py-0.5 truncate max-w-[120px]">{row.configuration}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Row Schema Diagnostics Panel */}
                {(() => {
                  const malformedCount = currentSheet.rows.filter(r => validateRow(r).length > 0).length;
                  return (
                    <div className="p-2.5 bg-[#121622] rounded-lg border border-slate-800 text-[11px] space-y-2">
                      <h4 className="font-bold text-slate-300 flex items-center gap-1.5">
                        <AlertTriangle className={`w-3.5 h-3.5 ${malformedCount > 0 ? "text-amber-400" : "text-emerald-500"}`} />
                        <span>Row-Level Diagnostics</span>
                      </h4>
                      {malformedCount > 0 ? (
                        <p className="text-[10px] text-amber-300 leading-relaxed">
                          Detected <span className="font-mono font-bold bg-[#1A1D28] px-1.5 py-0.5 rounded text-amber-400">{malformedCount}</span> out of <span className="font-mono">{currentSheet.rows.length}</span> rows with schema flags (e.g. formatting suggestions). All rows will be imported to the dashboard.
                        </p>
                      ) : (
                        <p className="text-[10px] text-emerald-400">
                          All {currentSheet.rows.length} rows fully comply with LWD schema rules! No errors found.
                        </p>
                      )}
                    </div>
                  );
                })()}
              </div>

              <div className="flex gap-2 justify-end pt-2 border-t border-slate-800">
                <button
                  onClick={() => {
                    setPendingSync(null);
                    setStatus({ type: "idle", message: "" });
                  }}
                  className="px-2.5 py-1.5 text-[10px] font-bold text-slate-400 hover:text-white rounded-lg transition-colors cursor-pointer"
                >
                  Discard
                </button>
                <button
                  onClick={() => {
                    const finalRows = currentSheet.rows;

                    onDataLoaded(finalRows);
                    setStatus({
                      type: "success",
                      message: `Loaded all ${finalRows.length} rows from sheet "${currentSheet.name}"!`
                    });
                    setPendingSync(null);
                    setTimeout(() => setStatus({ type: "idle", message: "" }), 5000);
                  }}
                  className="flex items-center gap-1 px-3 py-1.5 text-[10px] font-bold text-white bg-blue-600 hover:bg-blue-500 rounded-lg transition-colors cursor-pointer shadow-sm"
                >
                  <Check className="w-3 h-3" />
                  Sync Log
                </button>
              </div>
            </motion.div>
          );
        })() : (
          <motion.div
            whileHover={{ scale: 1.005 }}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            className={`border-2 border-dashed rounded-xl p-5 text-center transition-all cursor-pointer ${
              isDragging
                ? "border-blue-500 bg-blue-500/10"
                : isLightMode
                  ? "border-slate-300 hover:border-blue-500 bg-slate-50 hover:bg-blue-50/50"
                  : "border-slate-800/60 hover:border-blue-500/50 bg-[#151B2A]/20"
            }`}
            onClick={() => fileInputRef.current?.click()}
          >
            <div className="flex flex-col items-center justify-center gap-2">
              <div className={`w-9 h-9 rounded-xl flex items-center justify-center ${
                isLightMode ? "bg-blue-50 border border-blue-200 text-blue-600" : "bg-blue-950/50 border border-blue-900/30 text-blue-400"
              }`}>
                <Upload className="w-4.5 h-4.5 animate-pulse" />
              </div>
              <div>
                <p className={`text-xs font-semibold ${isLightMode ? "text-slate-800" : "text-slate-300"}`}>
                  Drop Excel file, or <span className="text-blue-600 hover:underline">browse</span>
                </p>
                <p className={`text-[10px] mt-1 font-sans ${isLightMode ? "text-slate-500" : "text-slate-500"}`}>
                  Supports .xlsx, .xls, and .csv files
                </p>
              </div>
            </div>
          </motion.div>
        )}

        {/* Upload Status Alert */}
        {status.type !== "idle" && (
          <motion.div
            initial={{ opacity: 0, y: 4 }}
            animate={{ opacity: 1, y: 0 }}
            className={`mt-3.5 flex items-start gap-2 p-3 rounded-xl text-xs transition-all leading-normal ${
              status.type === "success"
                ? "bg-emerald-950/20 border border-emerald-900/30 text-emerald-300"
                : "bg-rose-950/20 border border-rose-900/30 text-rose-300"
            }`}
            id="upload-status-indicator"
          >
            {status.type === "success" ? (
              <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
            ) : (
              <AlertCircle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
            )}
            <span className="break-words font-sans">
              {status.message}
            </span>
          </motion.div>
        )}
      </div>
    )}

    {/* Active Tab: OneDrive Live Sync */}
    {activeSubTab === "onedrive" && (
      <div className="space-y-4">
        {!msUser ? (
          <div className="space-y-3.5 text-center p-4 bg-slate-900/20 border border-slate-800/60 rounded-xl">
            <div className="w-9 h-9 rounded-full bg-cyan-950/50 border border-cyan-900/30 flex items-center justify-center text-cyan-400 mx-auto">
              <Cloud className="w-4.5 h-4.5" />
            </div>
            <div>
              <h4 className="text-xs font-bold text-white">Live-Connect Your OneDrive / SharePoint</h4>
              <p className="text-[11px] text-slate-400 mt-1 max-w-sm mx-auto font-sans leading-relaxed">
                Connect your OneDrive to live-sync your Excel workbook. Any changes to the Excel file will immediately reflect on this dashboard without manual reupload!
              </p>
            </div>
            <button
              onClick={handleMsSignIn}
              className="mx-auto flex items-center justify-center gap-2.5 px-4 py-2.5 bg-white hover:bg-slate-50 text-slate-900 font-bold text-xs rounded-xl transition-all shadow-md cursor-pointer border border-slate-200"
            >
              <svg className="w-4 h-4 shrink-0" viewBox="0 0 23 23">
                <rect x="1" y="1" width="10" height="10" fill="#f25022" />
                <rect x="12" y="1" width="10" height="10" fill="#7fba00" />
                <rect x="1" y="12" width="10" height="10" fill="#00a4ef" />
                <rect x="12" y="12" width="10" height="10" fill="#ffb900" />
              </svg>
              Sign in with Microsoft
            </button>
          </div>
        ) : (
          <div className="space-y-4">
            {/* Connected Microsoft Profile Summary */}
            <div className="flex items-center justify-between p-2.5 bg-[#151B2A]/40 border border-slate-800/80 rounded-xl text-xs">
              <div className="flex items-center gap-2 truncate">
                <div className="w-5 h-5 rounded-full bg-cyan-500 text-white flex items-center justify-center font-bold text-[10px] select-none uppercase">
                  {msUser.username ? msUser.username.charAt(0) : "M"}
                </div>
                <span className="text-slate-300 font-semibold truncate max-w-[200px]" title={msUser.username || msUser.name}>
                  {msUser.username || msUser.name}
                </span>
              </div>
              <button
                onClick={handleMsSignOut}
                className="text-[10px] font-bold text-slate-400 hover:text-white flex items-center gap-1 bg-slate-900 border border-slate-800 px-2.5 py-1 rounded cursor-pointer transition-colors whitespace-nowrap shrink-0"
              >
                <LogOut className="w-3 h-3" />
                Sign Out
              </button>
            </div>

            {/* Connected OneDrive Sheet Information */}
            {msFileId && (
              <div className="p-3 bg-cyan-950/15 border border-cyan-900/30 rounded-xl space-y-3 text-xs">
                <div className="flex items-center justify-between gap-2">
                  <div className="truncate">
                    <span className="text-[9px] text-slate-500 block uppercase font-bold tracking-wider">Active Connected OneDrive File</span>
                    <span className="font-semibold text-cyan-200 truncate block max-w-[280px]">{msFileName || "Connected File"}</span>
                    <span className="font-mono text-[9px] text-slate-400 block truncate max-w-[280px]">ID: {msFileId}</span>
                  </div>
                  <button
                    onClick={() => {
                      setMsFileId("");
                      setMsFileName("");
                      setMsLastSynced("");
                      localStorage.removeItem("ms_file_id");
                      localStorage.removeItem("ms_file_name");
                      localStorage.removeItem("ms_last_synced");
                    }}
                    className="text-[9px] font-bold text-rose-400 hover:text-rose-300 bg-rose-950/10 px-2 py-1 rounded border border-rose-900/30 cursor-pointer whitespace-nowrap"
                  >
                    Disconnect
                  </button>
                </div>

                <div className="flex items-center justify-between pt-2 border-t border-cyan-900/20">
                  <div className="flex items-center gap-1.5">
                    <span className="text-[10px] text-slate-400 font-semibold">Auto-Sync Loop</span>
                    <button
                      onClick={() => {
                        const nextAuto = !msAutoSync;
                        setMsAutoSync(nextAuto);
                        localStorage.setItem("ms_auto_sync", String(nextAuto));
                      }}
                      className="text-cyan-400 hover:text-cyan-300 focus:outline-none"
                    >
                      {msAutoSync ? (
                        <span className="inline-flex items-center text-[9px] bg-cyan-900/30 px-1.5 py-0.5 rounded border border-cyan-800/40 font-bold uppercase tracking-wider text-cyan-300">Enabled • 60s</span>
                      ) : (
                        <span className="inline-flex items-center text-[9px] bg-slate-900 px-1.5 py-0.5 rounded border border-slate-800 text-slate-400 font-bold uppercase tracking-wider">Disabled</span>
                      )}
                    </button>
                  </div>
                  <button
                    onClick={() => parseAndLoadOneDriveFile(msToken || "", msFileId, msFileName)}
                    disabled={isSyncingMs}
                    className="flex items-center gap-1.5 px-2.5 py-1 bg-cyan-900/30 hover:bg-cyan-800/30 text-cyan-300 rounded border border-cyan-700/30 text-[10px] font-bold transition-all disabled:opacity-50"
                  >
                    <RefreshCw className={`w-3 h-3 ${isSyncingMs ? "animate-spin" : ""}`} />
                    {isSyncingMs ? "Syncing..." : "Sync Now"}
                  </button>
                </div>
              </div>
            )}

            {/* OneDrive Files List */}
            <div className="space-y-1.5">
              <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Scan OneDrive Files</label>
              {isLoadingOneDrive ? (
                <div className="flex items-center justify-center py-5 gap-2 text-xs text-slate-400 bg-slate-900/20 rounded-xl border border-slate-900">
                  <RefreshCw className="w-4 h-4 animate-spin text-cyan-500" />
                  Scanning OneDrive...
                </div>
              ) : oneDriveFiles.length === 0 ? (
                <div className="text-center py-5 text-xs text-slate-600 bg-slate-900/20 border border-slate-900 rounded-xl">
                  No Excel (.xlsx) files found in your OneDrive root.
                </div>
              ) : (
                <div className="max-h-[140px] overflow-y-auto space-y-1 pr-1 border border-slate-900 bg-slate-950/20 p-1.5 rounded-xl" id="onedrive-files-list">
                  {oneDriveFiles.map((file) => (
                    <div
                      key={file.id}
                      onClick={async () => {
                        setMsFileId(file.id);
                        setMsFileName(file.name);
                        localStorage.setItem("ms_file_id", file.id);
                        localStorage.setItem("ms_file_name", file.name);
                        await parseAndLoadOneDriveFile(msToken || "", file.id, file.name);
                      }}
                      className={`flex items-center justify-between p-2 rounded-lg border transition-all cursor-pointer text-xs ${
                        msFileId === file.id
                          ? "bg-cyan-950/20 border-cyan-500/40 text-cyan-200"
                          : "bg-slate-900/30 border-transparent hover:bg-[#151B2A]/30 text-slate-300"
                      }`}
                    >
                      <div className="flex items-center gap-2 truncate">
                        <FileSpreadsheet className="w-3.5 h-3.5 text-cyan-500 shrink-0" />
                        <div className="truncate leading-tight">
                          <p className="font-semibold truncate text-[11px]">{file.name}</p>
                          <p className="text-[8px] text-slate-500 font-mono">
                            Modified: {file.lastModifiedDateTime ? new Date(file.lastModifiedDateTime).toLocaleDateString() : "Unknown"}
                          </p>
                        </div>
                      </div>
                      <button className="text-[10px] font-bold text-cyan-400 hover:underline px-1.5 py-0.5 shrink-0 font-sans">
                        {msFileId === file.id ? "Connected" : "Select"}
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {/* Sync Alerts */}
        {(status.type !== "idle" || msSyncError) && (
          <motion.div
            initial={{ opacity: 0, y: 4 }}
            animate={{ opacity: 1, y: 0 }}
            className={`mt-4 flex items-start gap-2 p-3 rounded-xl text-xs transition-all leading-normal ${
              status.type === "success" || (!msSyncError && status.type === "idle")
                ? "bg-emerald-950/20 border border-emerald-900/30 text-emerald-300"
                : "bg-rose-950/20 border border-rose-900/30 text-rose-300"
            }`}
            id="onedrive-status-indicator"
          >
            {status.type === "success" ? (
              <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
            ) : (
              <AlertCircle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
            )}
            <span className="break-words font-sans">
              {status.type === "success" ? status.message : (status.message || msSyncError)}
            </span>
          </motion.div>
        )}
      </div>
    )}
    </div>

    {/* Global Footer Controls */}
    <div className="flex items-center gap-2 mt-5 pt-4 border-t border-white/5">
      <button
        onClick={() => {
          setPendingSync(null);
          onRestoreDefault();
        }}
        className="flex-1 flex items-center justify-center gap-1.5 h-9 px-3 text-xs font-semibold text-[#9AA0A6] hover:text-[#F5F7FA] bg-[#1A1D21] border border-white/5 rounded-md transition-all cursor-pointer"
        id="restore-defaults-btn"
        title="Restore default screenshot sample data"
      >
        <RefreshCw className="w-3 h-3" />
        Default Log
      </button>

      <button
        onClick={() => {
          setPendingSync(null);
          onClearAllData();
        }}
        className="flex-1 flex items-center justify-center gap-1.5 h-9 px-3 text-xs font-semibold text-rose-400 hover:text-rose-300 hover:bg-rose-950/20 bg-[#1A1D21] border border-white/5 rounded-md transition-all cursor-pointer"
        id="clear-all-data-btn"
        title="Get rid of all data uploaded so far to start fresh"
      >
        <Trash2 className="w-3 h-3 text-rose-400" />
        Clear All Data
      </button>
      
      {activeSubTab === "file" ? (
        <button
          onClick={() => fileInputRef.current?.click()}
          className="flex-1 flex items-center justify-center gap-1.5 h-9 px-3 text-xs font-semibold text-white bg-[#3B82F6] hover:bg-[#3B82F6]/90 rounded-md transition-all cursor-pointer"
          id="trigger-file-upload"
        >
          <Upload className="w-3 h-3" />
          Upload Log
        </button>
      ) : (
        <button
          onClick={async () => {
            if (msFileId) {
              await parseAndLoadOneDriveFile(msToken || "", msFileId, msFileName);
            } else {
              setActiveSubTab("file");
            }
          }}
          disabled={isSyncingMs || !msFileId}
          className="flex-1 flex items-center justify-center gap-1.5 h-9 px-3 text-xs font-semibold text-white bg-cyan-600 hover:bg-cyan-500 rounded-md transition-all cursor-pointer disabled:opacity-50"
        >
          <RefreshCw className={`w-3 h-3 ${isSyncingMs ? "animate-spin" : ""}`} />
          {isSyncingMs ? "Syncing..." : "Sync OneDrive"}
        </button>
      )}

      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileInputChange}
        accept=".xlsx,.xls,.csv"
        className="hidden"
      />
    </div>
  </div>
</BorderGlow>
);
}

