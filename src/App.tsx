import React, { useState, useEffect, useLayoutEffect, useCallback, useMemo } from "react";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";
import { entraHandleReturn, entraSignInRedirect, entraSignOut, getAuthHeader, authFetch, getEntraProfile } from "./lib/entraAuth";
import { analyzeWorksheet, safeReadExcelWorkbook } from "./lib/excelParser";
import { DEFAULT_ABBREVIATIONS, RowData, Abbreviation, parseAnyDate, validateRow, formatMonthName, UserAccount, INITIAL_DEFAULT_USERS, UserRole, UserStatus, getHoleSizeFromSection, getSectionFromHoleSize } from "./types";
import UploadSync from "./components/UploadSync";
import StatsDashboard from "./components/StatsDashboard";
import SpreadsheetTable from "./components/SpreadsheetTable";
import ExcelGridEditor from "./components/ExcelGridEditor";
import RigWellGanttChart from "./components/RigWellGanttChart";
import AuthModal from "./components/AuthModal";
import AdminPortal from "./components/AdminPortal";
import BorderGlow from "./components/BorderGlow";
import { Layers, HelpCircle, RefreshCw, RotateCcw, SlidersHorizontal, HardHat, FileSpreadsheet, BookOpen, LayoutDashboard, Settings, Sun, Moon, Trash2, ChevronLeft, ChevronRight, Mail, Lock, ShieldCheck, LogOut, ArrowRight, ChevronDown, Check, Cloud, Table, Edit3, Eye, Shield, Users, CheckCircle2, ArrowLeftRight, Globe, ExternalLink, Wifi, MapPin, Activity, Link2, CalendarRange, Clock } from "lucide-react";

const sanitizeRowDate = (dateVal: any, expectedMonth?: string): { date: string; month: string } => {
  return parseAnyDate(dateVal, "DD/MM", expectedMonth);
};

const normalizeSectionSize = (section: string): string => {
  return getSectionFromHoleSize(section);
};

const correctAndSanitizeRows = (rows: RowData[]): RowData[] => {
  return rows.map(row => {
    let current = { ...row };

    // Auto-heal if qty contains configuration text (due to column shift bug)
    if (typeof current.qty === "string" || isNaN(Number(current.qty))) {
      const rawQtyStr = String(current.qty || "");
      if (rawQtyStr.length > 2 || /[a-zA-Z]/.test(rawQtyStr)) {
        const shiftedConfig = rawQtyStr;
        const shiftedFootageStr = String(current.configuration || "").replace(/,/g, "").trim();
        const shiftedFootage = parseFloat(shiftedFootageStr);
        const shiftedDate = String(current.footage || "");
        const shiftedMonth = String(current.date || "");
        const shiftedShop = String(current.month || "");
        
        current.configuration = shiftedConfig;
        current.footage = !isNaN(shiftedFootage) ? shiftedFootage : null;
        current.date = shiftedDate;
        current.month = shiftedMonth;
        current.shopDate = shiftedShop;
        current.qty = parseInt(String(current.holeSize || 1), 10) || 1;
      }
    }

    const derivedSection = (current.section !== undefined && current.section !== null && String(current.section).trim() !== "")
      ? getSectionFromHoleSize(String(current.section))
      : (current.holeSize ? getSectionFromHoleSize(current.holeSize) : "8in");

    const derivedHoleSize = (current.holeSize !== undefined && current.holeSize !== null && String(current.holeSize).trim() !== "")
      ? String(current.holeSize)
      : (derivedSection ? getHoleSizeFromSection(derivedSection) : "12.25\"");

    const rawMonthClean = formatMonthName(current.month);
    const sanitized = sanitizeRowDate(current.date, rawMonthClean);

    const correctedRow: RowData = {
      ...current,
      section: derivedSection,
      holeSize: derivedHoleSize,
      qty: typeof current.qty === "number" && !isNaN(current.qty) ? current.qty : (parseInt(String(current.qty || 1), 10) || 1),
      date: sanitized.date,
      month: rawMonthClean && rawMonthClean !== "N/A" ? rawMonthClean : sanitized.month
    };
    const errors = validateRow(correctedRow);
    if (errors.length > 0) {
      correctedRow.validationErrors = errors;
    } else {
      delete correctedRow.validationErrors;
    }
    return correctedRow;
  });
};

export default function App() {
  // 1. Core State
  const [data, setData] = useState<RowData[]>(() => {
    const saved = localStorage.getItem("internship_rig_data");
    let rawData: RowData[] = [];
    if (saved && saved.trim()) {
      try {
        rawData = JSON.parse(saved);
      } catch (e) {
        console.error("Error parsing local storage logs:", e);
      }
    }
    return correctAndSanitizeRows(rawData);
  });

  const [abbreviations, setAbbreviations] = useState<Abbreviation[]>(() => {
    const saved = localStorage.getItem("tool_abbreviations");
    let result = DEFAULT_ABBREVIATIONS;
    if (saved) {
      try {
        const parsed: Abbreviation[] = JSON.parse(saved);
        const mergedMap = new Map<string, Abbreviation>();
        parsed.forEach(item => {
          if (item && item.abbreviation) {
            const abbUpper = item.abbreviation.trim().toUpperCase();
            const fnUpper = (item.fullName || "").trim().toUpperCase();
            const isES = abbUpper.startsWith("ES") || abbUpper.startsWith("EARTHSTAR") || fnUpper.startsWith("ES") || fnUpper.startsWith("EARTHSTAR");
            const isEM = abbUpper.startsWith("EMBATT") || abbUpper.startsWith("EM BATT") || fnUpper === "EM BATTERY" || fnUpper === "EM BATT";
            const isICBM = abbUpper === "IC BM" || abbUpper === "IC(BM)" || abbUpper === "ICBM" || abbUpper === "IC-BM" || fnUpper === "IC BM";
            const cleanFullName = isES ? "EarthStar" : isEM ? "EM Battery" : isICBM ? "IC BM" : item.fullName;
            mergedMap.set(item.abbreviation.toUpperCase(), { ...item, fullName: cleanFullName });
          }
        });
        DEFAULT_ABBREVIATIONS.forEach(item => {
          mergedMap.set(item.abbreviation.toUpperCase(), item);
        });
        result = Array.from(mergedMap.values());
      } catch (e) {
        result = DEFAULT_ABBREVIATIONS;
      }
    }

    const allowedToolNames = new Set([
      "ADR", "AFR", "ALD", "BaseStar", "Battery", "EM Battery", "EM Batt", "EMBATT", "BST Caliper", "DGR", "Dummy Flex", 
      "EarthStar", "ES TX", "ES", "GWD", "GeoTap", "GyroStar", "ILS", "LCS", "LithoStar", "PWD", 
      "PulseStar", "ResiStar", "StrataStar", "UCS", "XBAT", "iCruise", "IC BM", "FDC", 
      "HCIM", "PCDC", "PCGK", "P4M", "DHS", "Pulsar", "FlowStar"
    ].map(name => name.toLowerCase()));

    return result.filter(item => item && item.fullName && allowedToolNames.has(item.fullName.toLowerCase()));
  });

  const [highlightedRowIds, setHighlightedRowIds] = useState<string[] | null>(null);
  const [focusOnMatches, setFocusOnMatches] = useState(false);
  const [activeTab, setActiveTab] = useState<"logs" | "excel-editor" | "rig-breakdown" | "gantt">("logs");
  const [selectedToolForRigPanel, setSelectedToolForRigPanel] = useState<string | null>(null);
  const [sidebarCollapsed, setSidebarCollapsed] = useState<boolean>(() => localStorage.getItem("sidebar_collapsed") === "true");
  const [isUploadOpen, setIsUploadOpen] = useState(false);
  const [isLightMode, setIsLightMode] = useState<boolean>(() => {
    return localStorage.getItem("theme_preference") === "light";
  });
  const [localLiveSyncLastUpdated, setLocalLiveSyncLastUpdated] = useState<string | null>(null);
  const [activeDataSource, setActiveDataSource] = useState<"excel_editor" | "uploaded_file">("excel_editor");
  const [syncNotification, setSyncNotification] = useState<string | null>(null);

  const dataRef = React.useRef(data);
  const lastLocalMutationTimeRef = React.useRef<number>(Date.now());
  const isSyncingLogsRef = React.useRef<boolean>(false);

  useEffect(() => {
    dataRef.current = data;
    lastLocalMutationTimeRef.current = Date.now();
  }, [data]);

  const [confirmModal, setConfirmModal] = useState<{
    isOpen: boolean;
    title: string;
    message: string;
    onConfirm: () => void;
  }>({
    isOpen: false,
    title: "",
    message: "",
    onConfirm: () => {}
  });

  const FORCED_LOGOUT_VERSION = "v_2026_07_30_reset";

  const [isLoggedIn, setIsLoggedIn] = useState<boolean>(() => {
    if (localStorage.getItem("auth_reset_version") !== FORCED_LOGOUT_VERSION) {
      localStorage.removeItem("is_logged_in");
      localStorage.removeItem("user_session");
      localStorage.removeItem("has_signed_in_before");
      localStorage.setItem("auth_reset_version", FORCED_LOGOUT_VERSION);
      return false;
    }
    return localStorage.getItem("is_logged_in") === "true";
  });
  const [currentUser, setCurrentUser] = useState<UserAccount>(() => {
    const saved = localStorage.getItem("user_session");
    if (saved && localStorage.getItem("auth_reset_version") === FORCED_LOGOUT_VERSION) {
      try {
        return JSON.parse(saved);
      } catch (e) {}
    }
    return {
      email: "admin@drillops.com",
      name: "Admin",
      role: "admin",
      permissions: "both",
      status: "approved",
      title: "Administrator (Full Edit Rights)"
    };
  });
  const currentUserRef = React.useRef(currentUser);
  useEffect(() => {
    currentUserRef.current = currentUser;
  }, [currentUser]);
  const [isAuthModalOpen, setIsAuthModalOpen] = useState(false);
  const [isAdminPortalOpen, setIsAdminPortalOpen] = useState(false);

  // Admin Active Users Presence State
  const [activeUserCount, setActiveUserCount] = useState<number>(1);
  const [activeUsersList, setActiveUsersList] = useState<any[]>([]);
  const [showActiveUsersPopover, setShowActiveUsersPopover] = useState(false);
  const sessionIdRef = React.useRef<string>(Math.random().toString(36).substring(2) + Date.now().toString(36));

  const handleSelectUser = (newUser: UserAccount) => {
    setCurrentUser(newUser);
    localStorage.setItem("user_session", JSON.stringify(newUser));
  };

  const handleOnLoginSuccess = (userAcc: UserAccount) => {
    setCurrentUser(userAcc);
    localStorage.setItem("user_session", JSON.stringify(userAcc));
    localStorage.setItem("is_logged_in", "true");

    const hasSignedInBefore = localStorage.getItem("has_signed_in_before") === "true";
    if (!hasSignedInBefore) {
      // First time signing in in this browser: clear all operational logs for a fresh workspace
      setData([]);
      localStorage.setItem("internship_rig_data", "[]");
      localStorage.setItem("has_signed_in_before", "true");
      fetch("/api/logs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ logs: [] })
      }).catch((e) => console.error("Error clearing initial logs on server:", e));

      setSyncNotification("First sign-in complete: Workspace cleared fresh. Your progress will be saved automatically.");
      setTimeout(() => setSyncNotification(null), 6000);
    }

    setIsLoggedIn(true);
  };

  const [loginError, setLoginError] = useState("");
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const account = await entraHandleReturn();
        if (!account) return;

        const profile = getEntraProfile();
        if (!profile) return;

        // Ask the server for our stored profile (role/status). First call
        // ever from this identity triggers server-side bootstrap: first
        // user becomes admin+approved, every subsequent one starts
        // viewer+pending until an admin approves them.
        let uData: any = null;
        try {
          const apiRes = await authFetch(`/api/users/${profile.oid}`);
          if (apiRes.ok) {
            uData = await apiRes.json();
          } else if (apiRes.status === 404) {
            const createRes = await authFetch(`/api/users/row`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ user: { id: profile.oid, name: profile.name } })
            });
            if (createRes.ok) {
              const created = await createRes.json();
              uData = created.user;
            }
          }
        } catch (e) {
          console.error("Error fetching/creating user profile:", e);
        }

        if (uData) {
          const userStatus: UserStatus = (uData.status as UserStatus) || "pending";
          const loggedInUser: UserAccount = {
            id: profile.oid,
            email: profile.email,
            name: uData.name || profile.name,
            role: uData.role || "viewer",
            permissions: uData.permissions || (uData.role === "admin" || uData.role === "editor" ? "both" : "read"),
            status: userStatus,
            title: uData.title || (userStatus === "pending" ? "Pending Access Approval" : uData.role === "admin" ? "Administrator" : uData.role === "editor" ? "Data Editor" : "Operations Viewer")
          };
          setCurrentUser(loggedInUser);
          setIsLoggedIn(true);
          localStorage.setItem("user_session", JSON.stringify(loggedInUser));
          localStorage.setItem("is_logged_in", "true");
        }
      } catch (err) {
        console.error("Error completing Entra ID sign-in:", err);
      }
    })();
  }, []);

  // Periodic sync for current user's profile status (approve/reject/role
  // changes an admin makes elsewhere show up here without a refresh).
  useEffect(() => {
    if (!currentUser?.id) return;

    const pollInterval = setInterval(async () => {
      try {
        const res = await authFetch(`/api/users/${currentUser.id}`);
        if (res.ok) {
          const freshData = await res.json();
          if (freshData?.status && freshData.status !== currentUser.status) {
            const updated: UserAccount = {
              ...currentUser,
              status: freshData.status as UserStatus,
              role: (freshData.role as UserRole) || currentUser.role,
              permissions: freshData.permissions || currentUser.permissions
            };
            setCurrentUser(updated);
            localStorage.setItem("user_session", JSON.stringify(updated));
          }
        }
      } catch (e) {}
    }, currentUser.status === "pending" ? 3000 : 12000);

    return () => clearInterval(pollInterval);
  }, [currentUser?.id, currentUser?.status]);

  const handleSignOut = async () => {
    localStorage.removeItem("is_logged_in");
    localStorage.removeItem("user_session");
    setIsLoggedIn(false);
    try {
      await entraSignOut();
    } catch (e) {}
  };

  const handlePerformLogin = async () => {
    setIsLoggingIn(true);
    setLoginError("");
    try {
      await entraSignInRedirect();
    } catch (err: any) {
      console.error("Entra sign-in error:", err);
      setLoginError(err?.message || "Sign-in failed. Please try again.");
      setIsLoggingIn(false);
    }
  };

  useLayoutEffect(() => {
    if (isLightMode) {
      document.documentElement.classList.add("light");
      localStorage.setItem("theme_preference", "light");
    } else {
      document.documentElement.classList.remove("light");
      localStorage.setItem("theme_preference", "dark");
    }
  }, [isLightMode]);

  const toggleTheme = useCallback(() => {
    document.documentElement.classList.add("theme-transitioning");
    setIsLightMode((prev) => {
      const next = !prev;
      if (next) {
        document.documentElement.classList.add("light");
        localStorage.setItem("theme_preference", "light");
      } else {
        document.documentElement.classList.remove("light");
        localStorage.setItem("theme_preference", "dark");
      }
      return next;
    });
    window.requestAnimationFrame(() => {
      window.requestAnimationFrame(() => {
        document.documentElement.classList.remove("theme-transitioning");
      });
    });
  }, []);

  // Load and sync data with server-side database endpoints on mount
  const [isServerLoading, setIsServerLoading] = useState(true);

  useEffect(() => {
    const fetchServerData = async () => {
      try {
        // Fetch logs from central server store
        const logsRes = await fetch("/api/logs");
        if (logsRes.ok) {
          const logsData = await logsRes.json();
          if (logsData.exists) {
            // Server file logs.json exists! This is the master database for ALL platform users.
            setData(correctAndSanitizeRows(logsData.logs || []));
          } else {
            // First time server run: if local storage has data, populate server
            const saved = localStorage.getItem("internship_rig_data");
            if (saved) {
              try {
                const parsed = JSON.parse(saved);
                if (Array.isArray(parsed) && parsed.length > 0) {
                  await fetch("/api/logs", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ logs: parsed })
                  });
                  setData(correctAndSanitizeRows(parsed));
                }
              } catch (e) {
                console.error(e);
              }
            }
          }
        }

        // Fetch abbreviations from central server store
        const abbRes = await fetch("/api/abbreviations");
        if (abbRes.ok) {
          const abbData = await abbRes.json();
          if (abbData.exists && Array.isArray(abbData.abbreviations) && abbData.abbreviations.length > 0) {
            const mergedMap = new Map<string, Abbreviation>();
            abbData.abbreviations.forEach((a: Abbreviation) => {
              if (a && a.abbreviation) {
                const abbUpper = (a.abbreviation || "").trim().toUpperCase();
                const fnUpper = (a.fullName || "").trim().toUpperCase();
                const isES = abbUpper.startsWith("ES") || abbUpper.startsWith("EARTHSTAR") || fnUpper.startsWith("ES") || fnUpper.startsWith("EARTHSTAR");
                const isEM = abbUpper.startsWith("EMBATT") || abbUpper.startsWith("EM BATT") || fnUpper === "EM BATTERY" || fnUpper === "EM BATT";
                const isICBM = abbUpper === "IC BM" || abbUpper === "IC(BM)" || abbUpper === "ICBM" || abbUpper === "IC-BM" || fnUpper === "IC BM";
                const cleanFullName = isES ? "EarthStar" : isEM ? "EM Battery" : isICBM ? "IC BM" : a.fullName;
                mergedMap.set(a.abbreviation.toUpperCase(), { ...a, fullName: cleanFullName });
              }
            });
            // Ensure all DEFAULT_ABBREVIATIONS are always included
            DEFAULT_ABBREVIATIONS.forEach((item) => {
              if (!mergedMap.has(item.abbreviation.toUpperCase())) {
                mergedMap.set(item.abbreviation.toUpperCase(), item);
              }
            });

            const allowedToolNames = new Set([
              "ADR", "AFR", "ALD", "BaseStar", "Battery", "EM Battery", "EM Batt", "EMBATT", "BST Caliper", "DGR", "Dummy Flex", 
              "EarthStar", "ES TX", "ES", "GWD", "GeoTap", "GyroStar", "ILS", "LCS", "LithoStar", "PWD", 
              "PulseStar", "ResiStar", "StrataStar", "UCS", "XBAT", "iCruise", "IC BM", "FDC", 
              "HCIM", "PCDC", "PCGK", "P4M", "DHS", "Pulsar", "FlowStar"
            ].map(name => name.toLowerCase()));

            const normalizedAbbs = Array.from(mergedMap.values()).filter(item => item && item.fullName && allowedToolNames.has(item.fullName.toLowerCase()));
            setAbbreviations(normalizedAbbs);
          } else {
            try {
              await fetch("/api/abbreviations", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ abbreviations: DEFAULT_ABBREVIATIONS })
              });
              setAbbreviations(DEFAULT_ABBREVIATIONS);
            } catch (e) {
              console.warn("Could not initial sync abbreviations to server API:", e);
            }
          }
        }
      } catch (err) {
        console.warn("Failed to sync with server-side API:", err);
      } finally {
        setIsServerLoading(false);
      }
    };

    fetchServerData();
  }, []);

  // Poll for live-pushed Excel data from /api/latest-excel every 15 seconds
  useEffect(() => {
    let active = true;
    const pollLatestExcel = async () => {
      try {
        const response = await fetch("/api/latest-excel", { cache: "no-store" });
        if (response.status === 404) {
          // Silent skip if no file pushed yet
          return;
        }
        if (!response.ok) {
          return;
        }
        
        const contentType = response.headers.get("content-type") || "";
        if (contentType.includes("text/html") || contentType.includes("application/json")) {
          // Received HTML fallback page or JSON response instead of raw excel binary
          return;
        }

        const arrayBuffer = await response.arrayBuffer();
        if (!active || !arrayBuffer || arrayBuffer.byteLength < 4) return;

        // Extra guard: verify first bytes are not HTML/JSON markup
        const u8 = new Uint8Array(arrayBuffer.slice(0, 100));
        const headStr = new TextDecoder().decode(u8).trim().toLowerCase();
        if (headStr.startsWith("<") || headStr.includes("<!doctype") || headStr.includes("<html") || headStr.startsWith("{")) {
          return;
        }

        const workbook = safeReadExcelWorkbook(arrayBuffer);
        const sheets = workbook.SheetNames.map(name => {
          return analyzeWorksheet(workbook.Sheets[name], name);
        }).filter(s => s.rawRowsCount > 0);

        if (sheets.length === 0) return;

        // Take sheet with highest match score
        sheets.sort((a, b) => b.matchScore - a.matchScore);
        const bestSheet = sheets[0];
        const parsedRows = correctAndSanitizeRows(bestSheet.rows);

        const isSame = (a: RowData[], b: RowData[]): boolean => {
          if (a.length !== b.length) return false;
          for (let i = 0; i < a.length; i++) {
            const rowA = a[i];
            const rowB = b[i];
            if (
              rowA.id !== rowB.id ||
              rowA.client !== rowB.client ||
              rowA.rig !== rowB.rig ||
              rowA.well !== rowB.well ||
              rowA.section !== rowB.section ||
              rowA.qty !== rowB.qty ||
              rowA.configuration !== rowB.configuration ||
              rowA.footage !== rowB.footage ||
              rowA.date !== rowB.date ||
              rowA.month !== rowB.month ||
              rowA.shopDate !== rowB.shopDate ||
              rowA.comments !== rowB.comments
            ) {
              return false;
            }
          }
          return true;
        };

        if (!isSame(dataRef.current, parsedRows)) {
          // Skip overwriting if local edits/deletions occurred within the last 30 seconds
          if (Date.now() - lastLocalMutationTimeRef.current < 30000) {
            return;
          }
          setData(parsedRows);
          setHighlightedRowIds(null);
          setFocusOnMatches(false);
        }

        const nowStr = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
        setLocalLiveSyncLastUpdated(nowStr);
      } catch (err: any) {
        if (err?.name === "AbortError" || err?.message?.includes("Failed to fetch") || err?.message?.includes("NetworkError")) {
          // Softly ignore transient network fetch failures during background polling
          return;
        }
        console.warn("Live Excel polling issue:", err?.message || err);
      }
    };

    pollLatestExcel();
    const intervalId = setInterval(pollLatestExcel, 15000);

    return () => {
      active = false;
      clearInterval(intervalId);
    };
  }, []);

  // 2. Dropdown Filter Specifications States
  const [selectedClients, setSelectedClients] = useState<string[]>([]);
  const [isClientDropdownOpen, setIsClientDropdownOpen] = useState(false);
  
  const [selectedRigs, setSelectedRigs] = useState<string[]>([]);
  const [isRigDropdownOpen, setIsRigDropdownOpen] = useState(false);
  
  const [selectedSection, setSelectedSection] = useState<string[]>([]);
  const [isSectionDropdownOpen, setIsSectionDropdownOpen] = useState(false);
  
  const [selectedShopDates, setSelectedShopDates] = useState<string[]>([]);
  const [isShopDateDropdownOpen, setIsShopDateDropdownOpen] = useState(false);
  
  const [selectedTools, setSelectedTools] = useState<string[]>([]);
  const [isToolDropdownOpen, setIsToolDropdownOpen] = useState(false);
  
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  const MONTHS_MAP: { [key: string]: string } = useMemo(() => ({
    jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06",
    jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12"
  }), []);

  const parseDateToISO = (dateStr: string): string => {
    if (!dateStr || dateStr.toLowerCase() === "n/a") return "";
    
    const trimmed = dateStr.trim();

    // 1. If it's a numeric string representing an Excel serial date
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

    // 2. If it's already in ISO format YYYY-MM-DD
    const isoRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (isoRegex.test(trimmed)) {
      return trimmed;
    }

    // 3. Try parsing standard DD-MMM-YY e.g., "22-Jun-26" or "22-Jun-2026"
    const parts = trimmed.split("-");
    if (parts.length === 3) {
      const monthStr = parts[1].toLowerCase();
      const month = MONTHS_MAP[monthStr];
      if (month) {
        let day = parts[0];
        if (day.length === 1) day = "0" + day;
        let year = parts[2];
        if (year.length === 2) {
          year = "20" + year;
        }
        return `${year}-${month}-${day}`;
      }
    }

    // 4. If it's a timestamp or any other standard JS date string format
    const parsed = Date.parse(trimmed);
    if (!isNaN(parsed)) {
      const d = new Date(parsed);
      if (d.getFullYear() < 3000) {
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, "0");
        const day = String(d.getDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
      }
    }

    return "";
  };

  const dateRangeBounds = useMemo(() => {
    if (!data || data.length === 0) return { min: "", max: "" };
    const isoDates = data.map(r => parseDateToISO(r.date)).filter(Boolean).sort();
    return {
      min: isoDates[0] || "",
      max: isoDates[isoDates.length - 1] || ""
    };
  }, [data]);

  // Compute dynamic months from the dataset (latest 2 months + full list of available months)
  const { latestDataMonths, allDataMonths } = useMemo(() => {
    const monthsSet = new Set<string>();
    const FULL_MONTH_NAMES = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];

    data.forEach((r) => {
      let iso = parseDateToISO(r.date);
      if (!iso && r.month) {
        const parsedM = parseAnyDate(r.month);
        if (parsedM.date) {
          iso = parseDateToISO(parsedM.date);
        }
      }
      if (iso && iso.length >= 7) {
        monthsSet.add(iso.substring(0, 7));
      }
    });

    const sorted = Array.from(monthsSet).sort();
    
    const allMonthsList = sorted.map((ym) => {
      const [yearStr, monthStr] = ym.split("-");
      const year = parseInt(yearStr, 10);
      const monthNum = parseInt(monthStr, 10);
      const monthName = FULL_MONTH_NAMES[monthNum - 1] || "Month";
      const start = `${yearStr}-${monthStr.padStart(2, "0")}-01`;
      const lastDay = new Date(year, monthNum, 0).getDate();
      const end = `${yearStr}-${monthStr.padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
      return {
        key: ym,
        label: `${monthName} ${year}`,
        shortLabel: `${monthName.slice(0, 3)} '${String(year).slice(-2)}`,
        start,
        end
      };
    });

    const targetYM = sorted.length > 0 ? sorted.slice(-2) : ["2026-06", "2026-07"];
    const latestList = targetYM.map((ym) => {
      const [yearStr, monthStr] = ym.split("-");
      const year = parseInt(yearStr, 10);
      const monthNum = parseInt(monthStr, 10);
      const monthName = FULL_MONTH_NAMES[monthNum - 1] || "Month";
      const start = `${yearStr}-${monthStr.padStart(2, "0")}-01`;
      const lastDay = new Date(year, monthNum, 0).getDate();
      const end = `${yearStr}-${monthStr.padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
      return {
        key: ym,
        label: `${monthName} ${year}`,
        start,
        end
      };
    });

    return {
      latestDataMonths: latestList,
      allDataMonths: allMonthsList
    };
  }, [data, MONTHS_MAP]);

  // unique tool names list for Select Tool dropdown
  const uniqueFullNames = useMemo(() => {
    const names = new Set<string>();
    abbreviations.forEach((abb) => {
      if (abb.fullName) names.add(abb.fullName);
    });
    return Array.from(names).sort();
  }, [abbreviations]);

  // Save states to local storage and sync with server datastore
  useEffect(() => {
    localStorage.setItem("internship_rig_data", JSON.stringify(data));
    dataRef.current = data;
    lastLocalMutationTimeRef.current = Date.now();

    if (!isServerLoading) {
      isSyncingLogsRef.current = true;
      const syncLogsToServer = async () => {
        try {
          const res = await fetch("/api/logs", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ logs: data })
          });
          if (!res.ok) {
            console.warn("Server response for logs sync:", res.status);
          }
        } catch (e) {
          console.warn("Could not sync logs to server API:", e);
        } finally {
          isSyncingLogsRef.current = false;
        }
      };
      syncLogsToServer();
    }
  }, [data, isServerLoading]);

  useEffect(() => {
    localStorage.setItem("tool_abbreviations", JSON.stringify(abbreviations));
    if (!isServerLoading) {
      const syncAbbsToServer = async () => {
        try {
          const res = await fetch("/api/abbreviations", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ abbreviations })
          });
          if (!res.ok) {
            console.warn("Server response for abbreviations sync:", res.status);
          }
        } catch (e) {
          console.warn("Could not sync abbreviations to server API:", e);
        }
      };
      syncAbbsToServer();
    }
  }, [abbreviations, isServerLoading]);

  // Real-time EventSource (SSE) + fast backup polling so data changes made by any user sync across all connected sessions instantly
  useEffect(() => {
    if (isServerLoading) return;

    const syncLatestData = async () => {
      // If user is currently typing in an input field or active element, skip polling overwrite
      const activeEl = document.activeElement;
      if (activeEl && (activeEl.tagName === "INPUT" || activeEl.tagName === "TEXTAREA" || activeEl.getAttribute("contenteditable") === "true")) {
        return;
      }

      // Do not overwrite local data if a local edit/deletion occurred recently (within 3 seconds) or a save is in progress
      if (Date.now() - lastLocalMutationTimeRef.current < 3000 || isSyncingLogsRef.current) {
        return;
      }

      try {
        const logsRes = await fetch("/api/logs");
        if (logsRes.ok) {
          const logsData = await logsRes.json();
          if (logsData.exists && Array.isArray(logsData.logs)) {
            const serverJson = JSON.stringify(logsData.logs);
            const currentJson = JSON.stringify(dataRef.current);
            if (serverJson !== currentJson) {
              setData(correctAndSanitizeRows(logsData.logs));
            }
          }
        }

        const abbRes = await fetch("/api/abbreviations");
        if (abbRes.ok) {
          const abbData = await abbRes.json();
          if (abbData.exists && Array.isArray(abbData.abbreviations)) {
            const serverAbbJson = JSON.stringify(abbData.abbreviations);
            const currentAbbJson = JSON.stringify(abbreviations);
            if (serverAbbJson !== currentAbbJson) {
              setAbbreviations(abbData.abbreviations);
            }
          }
        }

        const usersRes = await fetch("/api/users");
        if (usersRes.ok) {
          const usersData = await usersRes.json();
          if (Array.isArray(usersData.users)) {
            const curr = currentUserRef.current;
            const myAccount = usersData.users.find((u: any) =>
              (u.id && curr.id && u.id === curr.id) ||
              (u.email && curr.email && u.email.toLowerCase() === curr.email.toLowerCase()) ||
              (u.username && curr.username && u.username.toLowerCase() === curr.username.toLowerCase())
            );
            if (myAccount) {
              const prevStr = JSON.stringify(curr);
              const merged = { ...curr, ...myAccount };
              const nextStr = JSON.stringify(merged);
              if (prevStr !== nextStr) {
                setCurrentUser(merged);
                localStorage.setItem("user_session", JSON.stringify(merged));
              }
            }
          }
        }
      } catch (e) {
        // Silently ignore sync errors
      }
    };

    // 1. Establish SSE real-time stream
    let eventSource: EventSource | null = null;
    try {
      eventSource = new EventSource("/api/events");
      eventSource.onmessage = (event) => {
        try {
          const parsed = JSON.parse(event.data);
          if (parsed && (parsed.topic === "logs" || parsed.topic === "abbreviations" || parsed.topic === "users")) {
            syncLatestData();
          } else if (parsed && parsed.topic === "presence" && parsed.payload) {
            if (typeof parsed.payload.activeCount === "number") {
              setActiveUserCount(parsed.payload.activeCount);
              setActiveUsersList(parsed.payload.activeUsers || []);
            }
          }
        } catch (e) {}
      };
    } catch (e) {
      console.warn("SSE EventSource not supported or failed, falling back to polling", e);
    }

    // 2. High-speed backup polling (every 3 seconds)
    const intervalId = setInterval(syncLatestData, 3000);

    return () => {
      if (eventSource) {
        eventSource.close();
      }
      clearInterval(intervalId);
    };
  }, [isServerLoading, abbreviations]);

  // Presence Heartbeat Effect
  useEffect(() => {
    if (isServerLoading) return;

    const sendPresenceHeartbeat = async () => {
      try {
        const activeTabName =
          activeTab === "logs"
            ? "MWD/LWD Master Logs"
            : activeTab === "excel-editor"
            ? "Excel Data Grid Editor"
            : activeTab === "rig-breakdown"
            ? "Rig Breakdown Analysis"
            : activeTab === "gantt"
            ? "Rig & Well Gantt Schedule"
            : "Dashboard Overview";

        const res = await fetch("/api/presence/heartbeat", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            sessionId: sessionIdRef.current,
            userId: currentUser?.id,
            username: currentUser?.username || (currentUser?.email ? currentUser.email.split("@")[0] : "user"),
            name: currentUser?.name || "User",
            role: currentUser?.role || "viewer",
            activeTab: activeTab,
            activeTabName: activeTabName,
            activeUrl: `/#${activeTab}`
          })
        });
        if (res.ok) {
          const data = await res.json();
          if (typeof data.activeCount === "number") {
            setActiveUserCount(data.activeCount);
            setActiveUsersList(data.activeUsers || []);
          }
        }
      } catch (e) {
        // Silently ignore presence heartbeat errors
      }
    };

    sendPresenceHeartbeat();
    const presenceInterval = setInterval(sendPresenceHeartbeat, 5000);
    return () => clearInterval(presenceInterval);
  }, [currentUser, isServerLoading, activeTab]);

  // 3. Dynamic Filter Lists
  const clientOptions = useMemo(() => {
    const clients = new Set<string>();
    data.forEach((row) => {
      if (row.client) {
        const cTrim = row.client.trim();
        if (cTrim && cTrim !== "N/A") {
          clients.add(cTrim);
        }
      }
    });
    return Array.from(clients).sort();
  }, [data]);

  const rigOptions = useMemo(() => {
    const rigs = new Set<string>();
    const selectedClientsLower = selectedClients.map(c => c.trim().toLowerCase());
    data.forEach((row) => {
      const rowClientClean = row.client ? row.client.trim().toLowerCase() : "";
      if (selectedClients.length === 0 || selectedClientsLower.includes(rowClientClean)) {
        if (row.rig) {
          rigs.add(row.rig.trim());
        }
      }
    });
    return Array.from(rigs).sort();
  }, [data, selectedClients]);

  const sectionOptions = useMemo(() => {
    const sections = new Set<string>();
    data.forEach((row) => {
      const raw = row.section || row.holeSize || "";
      if (raw) {
        const converted = getSectionFromHoleSize(raw);
        if (converted) {
          sections.add(converted);
        }
      }
    });
    const priority = ["4.75in", "6.75in", "8in", "9.5in"];
    const ordered = Array.from(sections);
    ordered.sort((a, b) => {
      const indexA = priority.indexOf(a);
      const indexB = priority.indexOf(b);
      if (indexA !== -1 && indexB !== -1) return indexA - indexB;
      if (indexA !== -1) return -1;
      if (indexB !== -1) return 1;
      return a.localeCompare(b);
    });
    return ordered.length > 0 ? ordered : priority;
  }, [data]);

  // Reset Rig filter if clients change and current Rigs are not in option list
  useEffect(() => {
    const standardizedRigOptions = rigOptions.map(r => r.toLowerCase());
    const validSelectedRigs = selectedRigs.filter(r => standardizedRigOptions.includes(r.toLowerCase()));
    if (validSelectedRigs.length !== selectedRigs.length) {
      setSelectedRigs(validSelectedRigs);
    }
  }, [rigOptions, selectedRigs]);

  // 4. Data Processing with Memoization and Pre-calculated Tool Specifications
  const filteredData = useMemo(() => {
    // A. Start with base data
    let result = data;

    // B. If AI Focus Mode is active and we have highlighted IDs, restrict data to those IDs
    if (focusOnMatches && highlightedRowIds && highlightedRowIds.length > 0) {
      const highlightedSet = new Set(highlightedRowIds);
      result = result.filter((row) => highlightedSet.has(row.id));
    }

    const selectedClientsLower = selectedClients.map((c) => c.trim().toLowerCase());
    const selectedRigsLower = selectedRigs.map((r) => r.trim().toLowerCase());
    const selectedSectionsLower = selectedSection.map((s) => s.trim().toLowerCase());
    const selectedShopDatesLower = selectedShopDates.map((s) => s.toLowerCase());

    // Pre-build tool specs to avoid repetitive allocations inside the row loop
    const toolSpecs = selectedTools.map((tool) => {
      const toolLower = tool.toLowerCase();
      const isEarthStar = toolLower === "earthstar" || toolLower === "earth star" || toolLower.startsWith("es");
      const isEMBattery = toolLower === "em battery" || toolLower === "em batt" || toolLower === "em batteries" || toolLower.startsWith("embatt");
      const isStandardBattery = (toolLower === "battery" || toolLower === "batt") && !isEMBattery;
      const isBSTCaliper = toolLower === "bst caliper" || toolLower === "caliper" || toolLower === "bstcaliper";
      const isStrataStar = toolLower === "stratastar" || toolLower === "strata star" || toolLower === "ss" || toolLower.startsWith("stratastar");
      const isFlowStar = toolLower === "flowstar" || toolLower === "flow star" || toolLower === "flow-star" || toolLower === "fst" || toolLower.startsWith("flowstar");
      const isICBM = toolLower === "ic bm" || toolLower === "ic(bm)" || toolLower === "icbm" || toolLower === "ic-bm" || toolLower === "ic (bm)";
      const isICruise = (toolLower === "icruise" || toolLower === "ic") && !isICBM;

      const toolTerms = [
        toolLower,
        ...abbreviations
          .filter((a) => a.fullName.toLowerCase() === toolLower)
          .map((a) => a.abbreviation.toLowerCase())
      ];
      return { isEarthStar, isEMBattery, isStandardBattery, isBSTCaliper, isStrataStar, isFlowStar, isICBM, isICruise, toolTerms };
    });

    const standardBatteryTerms = [
      "battery", "batteries", "batt", "batt1", "batt2", "istarbatt", "istarbatt1", "istarbatt2", "btt", "qbat"
    ];

    // D. Apply specifications dropdown and date range filters
    return result.filter((row) => {
      const rowClientClean = row.client ? row.client.trim().toLowerCase() : "";
      const matchClient = selectedClients.length === 0 || selectedClientsLower.includes(rowClientClean);
      if (!matchClient) return false;

      const rowRigClean = row.rig ? row.rig.trim().toLowerCase() : "";
      const matchRig = selectedRigs.length === 0 || selectedRigsLower.includes(rowRigClean);
      if (!matchRig) return false;

      const rowToolSize = getSectionFromHoleSize(row.section || row.holeSize || "").toLowerCase();
      const rowSectionClean = row.section ? row.section.trim().toLowerCase() : "";
      const rowHoleSizeClean = row.holeSize ? row.holeSize.trim().toLowerCase() : "";
      const matchSection = selectedSection.length === 0 ||
        selectedSectionsLower.includes(rowToolSize) ||
        selectedSectionsLower.includes(rowSectionClean) ||
        selectedSectionsLower.includes(rowHoleSizeClean);
      if (!matchSection) return false;

      if (selectedShopDates.length > 0) {
        const rowShopClean = row.shopDate ? row.shopDate.trim().toLowerCase() : "";
        const isDelivered = rowShopClean === "delivered";
        const isPending = !rowShopClean || rowShopClean !== "delivered";
        const matchShop = (selectedShopDatesLower.includes("delivered") && isDelivered) || (selectedShopDatesLower.includes("pending") && isPending);
        if (!matchShop) return false;
      }

      // Check date range
      if (row.date) {
        const rowISODate = parseDateToISO(row.date);
        if (startDate && rowISODate < startDate) return false;
        if (endDate && rowISODate > endDate) return false;
      } else if (startDate || endDate) {
        return false;
      }

      // Check tool terms in configuration and comments (case insensitive)
      if (toolSpecs.length > 0) {
        const configLower = (row.configuration || "").toLowerCase();
        const commentsLower = (row.comments || "").toLowerCase();

        if (!configLower && !commentsLower) return false;

        const fullTextLower = `${configLower} ${commentsLower}`;
        const words = fullTextLower.split(/[^a-z0-9]+/);

        const matchTool = toolSpecs.some((spec) => {
          if (spec.isICBM) {
            return /IC[\s\-_]*\(?BM\)?\b|ICBM\b/i.test(fullTextLower);
          }
          if (spec.isICruise) {
            // Strip out IC BM instances completely so they never match iCruise
            const textWithoutICBM = fullTextLower.replace(/IC[\s\-_]*\(?BM\)?\b|ICBM\b/gi, " ");
            const wordsWithoutICBM = textWithoutICBM.split(/[^a-z0-9]+/);
            return wordsWithoutICBM.includes("ic") || wordsWithoutICBM.includes("icruise") || /\bicruise\b/i.test(textWithoutICBM);
          }
          if (spec.isEarthStar) {
            return words.some((w) => w.startsWith("es") || w.startsWith("exrx") || w.startsWith("earthstar"));
          }
          if (spec.isEMBattery) {
            return /em[\s\-_]*(?:batteries|battery|batt\d*|battpack|batt)/i.test(fullTextLower) || words.some((w) => w.startsWith("embatt"));
          }
          if (spec.isStandardBattery) {
            const EM_BATTERY_EXCLUSION_REGEX = /em[\s\-_]*(?:batteries|battery|batt\d*|battpack|batt)|smart\s+battery\s+bulkhead|battery\s+relief\s+bulk\s+head/gi;
            const cleanedText = fullTextLower.replace(EM_BATTERY_EXCLUSION_REGEX, " ");
            const cleanedWords = cleanedText.split(/[^a-z0-9]+/);
            return standardBatteryTerms.some((term) => cleanedWords.includes(term));
          }
          if (spec.isBSTCaliper) {
            return /bst[\s\-_]*caliper|\bcaliper\b|\bbstcaliper\b/i.test(fullTextLower);
          }
          if (spec.isStrataStar) {
            return /\bstratastar\b|\bstrata[\s\-_]*star\b|\bss\d*\b/i.test(fullTextLower);
          }
          if (spec.isFlowStar) {
            return /flow[\s\-_]*star|fst\d*|flowstar\d*/i.test(fullTextLower) || words.some((w) => w.startsWith("fst") || w.startsWith("flowstar") || w === "flow");
          }

          return spec.toolTerms.some((term) => words.includes(term) || fullTextLower.includes(term));
        });

        if (!matchTool) return false;
      }

      return true;
    });
  }, [data, selectedClients, selectedRigs, selectedSection, selectedShopDates, selectedTools, startDate, endDate, abbreviations, highlightedRowIds, focusOnMatches]);

  // 5. Handlers wrapped in useCallback for enterprise responsiveness
  const handleUpdateFromExcelEditor = useCallback(() => {
    try {
      const saved = localStorage.getItem("internship_rig_data");
      let rawData: RowData[] = [];
      if (saved) {
        try {
          rawData = JSON.parse(saved);
        } catch (e) {
          rawData = dataRef.current;
        }
      } else {
        rawData = dataRef.current;
      }
      const corrected = correctAndSanitizeRows(rawData);
      setData(corrected);
      setHighlightedRowIds(null);
      setFocusOnMatches(false);

      const timeStr = new Date().toLocaleTimeString();
      setLocalLiveSyncLastUpdated(timeStr);
      setActiveDataSource("excel_editor");

      setSyncNotification(`Entire dashboard connected & updated from Excel Sheet Editor (${corrected.length} operational logs).`);
      setTimeout(() => setSyncNotification(null), 5000);
    } catch (e) {
      console.error("Error updating from Excel Sheet Editor:", e);
    }
  }, []);

  const handleDataLoaded = useCallback((
    newData: RowData[],
    source: "excel_editor" | "uploaded_file" = "uploaded_file"
  ) => {
    const sanitized = correctAndSanitizeRows(newData);
    setData(sanitized);

    // Save to localStorage immediately
    try {
      localStorage.setItem("internship_rig_data", JSON.stringify(sanitized));
    } catch (e) {
      console.error("Error saving logs to localStorage:", e);
    }

    // POST updated logs to server API immediately to keep central storage in sync
    fetch("/api/logs", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ logs: sanitized })
    }).catch((e) => console.error("Error persisting updated logs to server API:", e));

    lastLocalMutationTimeRef.current = Date.now();
    setHighlightedRowIds(null);
    setFocusOnMatches(false);
    setActiveDataSource(source);
    if (source === "excel_editor") {
      setLocalLiveSyncLastUpdated(new Date().toLocaleTimeString());
    }
  }, []);

  const handleRestoreDefault = useCallback(() => {
    setConfirmModal({
      isOpen: true,
      title: "Restore Default Dataset",
      message: "Are you sure you want to restore the default screenshot internship log dataset? Any current uploads and syncs will be reset.",
      onConfirm: () => {
        setData([]);
        setAbbreviations(DEFAULT_ABBREVIATIONS);
        setHighlightedRowIds(null);
        setFocusOnMatches(false);
        setSelectedClients([]);
        setSelectedRigs([]);
        setSelectedSection([]);
        setSelectedShopDates([]);
        setSelectedTools([]);
      }
    });
  }, []);

  const handleClearAllData = useCallback(() => {
    setConfirmModal({
      isOpen: true,
      title: "Clear All Log Data",
      message: "Are you sure you want to completely clear all data? This will empty the dashboard so you can upload your custom sheet fresh.",
      onConfirm: () => {
        setData([]);
        setHighlightedRowIds(null);
        setFocusOnMatches(false);
        setSelectedClients([]);
        setSelectedRigs([]);
        setSelectedSection([]);
        setSelectedShopDates([]);
        setSelectedTools([]);
        setStartDate("");
        setEndDate("");
        localStorage.setItem("internship_rig_data", "[]");
      }
    });
  }, []);

  const handleAddAbbreviation = useCallback((newAbb: Abbreviation) => {
    setAbbreviations((prev) => [...prev, newAbb]);
  }, []);

  const handleDeleteAbbreviation = useCallback((abbToDelete: string) => {
    setAbbreviations((prev) => prev.filter((a) => a.abbreviation !== abbToDelete));
  }, []);

  const handleUpdateGroup = useCallback((
    oldFullName: string,
    newFullName: string,
    newAbbreviations: string[],
    category: string,
    description: string
  ) => {
    setAbbreviations((prev) => {
      const upperNewAbbs = newAbbreviations.map(a => a.trim().toUpperCase()).filter(Boolean);
      const filtered = prev.filter(
        (a) => a.fullName.toLowerCase() !== oldFullName.toLowerCase() &&
               !upperNewAbbs.includes(a.abbreviation.toUpperCase())
      );
      const newEntries = upperNewAbbs.map((abb) => ({
        abbreviation: abb,
        fullName: newFullName.trim(),
        category: category,
        description: description
      }));
      return [...filtered, ...newEntries];
    });
  }, []);

  const handleAbbreviationClick = useCallback((_abb: string) => {}, []);

  const handleSetHighlightedRows = useCallback((ids: string[] | null) => {
    setHighlightedRowIds(ids);
    if (!ids) {
      setFocusOnMatches(false);
    }
  }, []);

  // Reset specification filters
  const resetFilters = useCallback(() => {
    setSelectedClients([]);
    setSelectedRigs([]);
    setSelectedSection([]);
    setSelectedShopDates([]);
    setSelectedTools([]);
    setStartDate("");
    setEndDate("");
    setHighlightedRowIds(null);
    setFocusOnMatches(false);
  }, []);

  const isFiltered = selectedClients.length > 0 || selectedRigs.length > 0 || selectedSection.length > 0 || selectedShopDates.length > 0 || selectedTools.length > 0 || startDate !== "" || endDate !== "" || focusOnMatches;

  if (!isLoggedIn) {
    return (
      <AnimatePresence mode="wait">
        <motion.div
          key="login"
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          className="min-h-screen bg-[#060709] font-sans text-[#F5F7FA] antialiased flex items-center justify-center relative overflow-hidden p-4 selection:bg-[#3B82F6]/30 selection:text-[#F5F7FA] w-full"
        >
          {/* Animated Background Visuals */}
          <div className="absolute inset-0 z-0 overflow-hidden pointer-events-none">
            {/* Grid Overlay */}
            <div className="absolute inset-0 bg-[linear-gradient(to_right,#141822_1px,transparent_1px),linear-gradient(to_bottom,#141822_1px,transparent_1px)] bg-[size:4rem_4rem] [mask-image:radial-gradient(ellipse_60%_50%_at_50%_50%,#000_70%,transparent_100%)] opacity-30" />
            
            {/* Radial glow effects */}
            <div className="absolute top-[-10%] left-[-10%] w-[50%] h-[50%] rounded-full bg-blue-500/10 blur-[120px]" />
            <div className="absolute bottom-[-10%] right-[-10%] w-[50%] h-[50%] rounded-full bg-emerald-500/5 blur-[120px]" />
            
            {/* Animated slow floating circles */}
            <motion.div
              animate={{
                y: [0, -20, 0],
                x: [0, 10, 0],
              }}
              transition={{
                duration: 10,
                repeat: Infinity,
                ease: "easeInOut"
              }}
              className="absolute top-1/4 left-1/3 w-72 h-72 rounded-full bg-blue-600/5 blur-[100px]"
            />
            <motion.div
              animate={{
                y: [0, 30, 0],
                x: [0, -15, 0],
              }}
              transition={{
                duration: 12,
                repeat: Infinity,
                ease: "easeInOut"
              }}
              className="absolute bottom-1/3 right-1/4 w-96 h-96 rounded-full bg-indigo-500/5 blur-[110px]"
            />
          </div>

          <div className="relative z-10 max-w-md w-full">
            <BorderGlow
              borderRadius={24}
              backgroundColor="#0A0D14"
              glowColor="220 80 50"
              glowIntensity={1.2}
              edgeSensitivity={20}
              className="w-full shadow-2xl border border-white/[0.03]"
            >
              <div className="p-8 sm:p-10 flex flex-col items-center">
                
                {/* Microsoft / Entra ID sign-in */}
                <div className="w-full space-y-4">
                  {loginError && (
                    <motion.div 
                      initial={{ opacity: 0, y: -4 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="p-3 bg-rose-500/10 border border-rose-500/20 text-rose-400 text-xs rounded-xl flex items-center gap-2"
                    >
                      <span className="w-1.5 h-1.5 rounded-full bg-rose-500 shrink-0" />
                      <span>{loginError}</span>
                    </motion.div>
                  )}

                  <motion.div
                    initial={{ y: 10, opacity: 0 }}
                    animate={{ y: 0, opacity: 1 }}
                    transition={{ delay: 0.4 }}
                    className="pt-2 space-y-3"
                  >
                    <button
                      type="button"
                      onClick={handlePerformLogin}
                      disabled={isLoggingIn}
                      className="w-full h-10 rounded-xl bg-blue-600 hover:bg-blue-500 active:scale-[0.98] text-white text-xs font-bold transition-all shadow-lg shadow-blue-500/10 cursor-pointer flex items-center justify-center gap-2 select-none border border-blue-500/20"
                    >
                      {isLoggingIn ? (
                        <>
                          <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                          <span>Redirecting to Microsoft...</span>
                        </>
                      ) : (
                        <>
                          <span>Sign in with Microsoft</span>
                          <ArrowRight className="w-3.5 h-3.5" />
                        </>
                      )}
                    </button>

                    <div className="flex items-center justify-center text-[10px] text-slate-500 pt-1 font-mono">
                      <span>Sign in with your Halliburton account</span>
                    </div>
                  </motion.div>
                </div>

                {/* Additional footer-like details inside the login card */}
                <motion.div
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  transition={{ delay: 0.6 }}
                  className="mt-8 pt-6 border-t border-white/[0.03] w-full flex items-center justify-center gap-2.5 text-[10px] text-slate-600 font-mono"
                >
                  <ShieldCheck className="w-3.5 h-3.5 text-emerald-500/60" />
                  <span>256-bit AES Sandboxed Shell</span>
                </motion.div>

              </div>
            </BorderGlow>
          </div>
        </motion.div>
      </AnimatePresence>
    );
  }

  // Access-approval gate: Block pending users from reaching any dashboard views or routes
  if (currentUser?.status === "pending") {
    return (
      <div 
        className="min-h-screen bg-[#060709] font-sans text-[#F5F7FA] antialiased flex items-center justify-center relative overflow-hidden p-6 selection:bg-[#3B82F6]/30 selection:text-[#F5F7FA] w-full" 
        id="access-approval-gate"
      >
        {/* Animated Background Visuals */}
        <div className="absolute inset-0 z-0 overflow-hidden pointer-events-none">
          <div className="absolute inset-0 bg-[linear-gradient(to_right,#141822_1px,transparent_1px),linear-gradient(to_bottom,#141822_1px,transparent_1px)] bg-[size:4rem_4rem] [mask-image:radial-gradient(ellipse_60%_50%_at_50%_50%,#000_70%,transparent_100%)] opacity-30" />
          <div className="absolute top-1/4 left-1/2 -translate-x-1/2 w-96 h-96 rounded-full bg-amber-500/10 blur-[120px]" />
          <div className="absolute bottom-1/4 right-1/4 w-80 h-80 rounded-full bg-blue-500/5 blur-[100px]" />
        </div>

        <div className="relative z-10 max-w-md w-full">
          <BorderGlow
            borderRadius={24}
            backgroundColor="#0A0D14"
            glowColor="45 100 50"
            glowIntensity={1.2}
            edgeSensitivity={20}
            className="w-full shadow-2xl border border-amber-500/20"
          >
            <div className="p-8 sm:p-10 flex flex-col items-center text-center">
              {/* Pulsing Status Icon */}
              <div className="w-16 h-16 rounded-2xl bg-amber-500/15 border border-amber-500/30 flex items-center justify-center mb-6 shadow-inner shadow-amber-500/10">
                <Clock className="w-8 h-8 text-amber-400 animate-pulse" />
              </div>

              {/* Status Badge */}
              <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-amber-500/10 border border-amber-500/30 text-amber-300 text-xs font-semibold mb-4 font-mono">
                <span className="w-2 h-2 rounded-full bg-amber-400 animate-ping" />
                <span>Pending Approval</span>
              </div>

              {/* Headline */}
              <h2 className="text-xl font-bold tracking-tight text-white mb-3 font-sans">
                Access Request Submitted
              </h2>

              {/* Exact Message Required by Specification */}
              <p className="text-sm text-slate-300 leading-relaxed mb-6 font-normal">
                Your access request is awaiting admin approval. You'll be notified once an admin reviews your account.
              </p>

              {/* User Account Info */}
              <div className="w-full p-3 rounded-xl bg-slate-900/80 border border-white/5 mb-6 text-xs text-slate-400 flex items-center justify-between font-mono">
                <span className="text-slate-500">Signed in as:</span>
                <span className="text-slate-200 font-semibold truncate max-w-[200px]" title={currentUser.email || currentUser.name}>
                  {currentUser.email || currentUser.name || "User"}
                </span>
              </div>

              {/* Sign out button */}
              <button
                onClick={handleSignOut}
                className="w-full h-10 rounded-xl bg-slate-800 hover:bg-slate-700 active:scale-[0.98] text-slate-200 hover:text-white text-xs font-bold transition-all border border-slate-700/80 cursor-pointer flex items-center justify-center gap-2 shadow-sm"
                id="pending-sign-out-btn"
              >
                <LogOut className="w-4 h-4 text-slate-400" />
                <span>Sign out</span>
              </button>
            </div>
          </BorderGlow>
        </div>
      </div>
    );
  }

  return (
    <div className={`min-h-screen font-sans antialiased flex transition-colors duration-200 selection:bg-[#3B82F6]/30 ${
      isLightMode ? "bg-[#525252] text-slate-900 selection:text-slate-900" : "bg-[#0B0B0D] text-[#F5F7FA] selection:text-[#F5F7FA]"
    }`} id="app-root">
      
      {/* Left Sidebar Layout */}
      <aside 
        className={`border-r flex flex-col justify-between shrink-0 sticky top-0 h-screen z-40 transition-all duration-300 ${
          isLightMode ? "bg-[#636363] border-[#3F3F3F] text-slate-900" : "bg-[#0B0B0D] border-white/5 text-[#F5F7FA]"
        } ${
          sidebarCollapsed ? "w-[72px] p-3" : "w-[240px] p-5"
        }`} 
        id="app-sidebar"
      >
        <div>
          {/* Logo & Platform Info & Toggle button */}
          {!sidebarCollapsed ? (
            <div className="flex items-center justify-between mb-8 gap-2">
              <div className="truncate">
                <h1 className={`text-sm font-bold tracking-tight leading-tight font-sans uppercase truncate ${
                  isLightMode ? "text-slate-900" : "text-[#F5F7FA]"
                }`}>
                  MLWD Gantt Chart
                </h1>
              </div>
              <button
                onClick={() => {
                  setSidebarCollapsed(true);
                  localStorage.setItem("sidebar_collapsed", "true");
                }}
                className={`p-1.5 rounded-md transition-all cursor-pointer shrink-0 ${
                  isLightMode ? "hover:bg-slate-100 text-slate-500 hover:text-slate-900" : "hover:bg-white/5 text-[#9AA0A6] hover:text-[#F5F7FA]"
                }`}
                title="Collapse Sidebar"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
            </div>
          ) : (
            <div className="flex flex-col items-center gap-4 mb-8">
              <button
                onClick={() => {
                  setSidebarCollapsed(false);
                  localStorage.setItem("sidebar_collapsed", "false");
                }}
                className={`p-1.5 rounded-md transition-all cursor-pointer ${
                  isLightMode ? "hover:bg-slate-100 text-slate-500 hover:text-slate-900" : "hover:bg-white/5 text-[#9AA0A6] hover:text-[#F5F7FA]"
                }`}
                title="Expand Sidebar - MLWD Gantt Chart"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          )}

          {/* Navigation Items */}
          <nav className="space-y-1" id="sidebar-navigation">
            <button
              onClick={() => setActiveTab("logs")}
              className={`w-full flex items-center rounded-lg text-xs font-semibold transition-all cursor-pointer ${
                sidebarCollapsed ? "justify-center p-2.5" : "gap-3 px-3 py-2.5"
              } ${
                activeTab === "logs"
                  ? (isLightMode ? "bg-white text-blue-800 font-bold border border-slate-300 shadow-sm" : "bg-white/5 text-[#F5F7FA]")
                  : (isLightMode ? "text-slate-100 hover:text-white hover:bg-white/10 bg-transparent" : "text-[#9AA0A6] hover:text-[#F5F7FA] bg-transparent")
              }`}
              title="Operational Logs"
            >
              <FileSpreadsheet className="w-4 h-4 text-[#3B82F6] shrink-0" />
              {!sidebarCollapsed && <span className="truncate">Operational Logs</span>}
            </button>
            <button
              onClick={() => setActiveTab("excel-editor")}
              className={`w-full flex items-center rounded-lg text-xs font-semibold transition-all cursor-pointer ${
                sidebarCollapsed ? "justify-center p-2.5" : "gap-3 px-3 py-2.5"
              } ${
                activeTab === "excel-editor"
                  ? (isLightMode ? "bg-white text-purple-800 font-bold border border-slate-300 shadow-sm" : "bg-white/5 text-[#F5F7FA]")
                  : (isLightMode ? "text-slate-100 hover:text-white hover:bg-white/10 bg-transparent" : "text-[#9AA0A6] hover:text-[#F5F7FA] bg-transparent")
              }`}
              title="Excel Sheet Editor"
            >
              <Table className="w-4 h-4 text-purple-400 shrink-0" />
              {!sidebarCollapsed && <span className="truncate flex items-center gap-1.5">Excel Sheet Editor <span className={`px-1.5 py-0.2 text-[9px] rounded font-mono ${isLightMode ? "bg-purple-100 text-purple-800 font-bold" : "bg-purple-500/20 text-purple-300"}`}>EDIT</span></span>}
            </button>
            <button
              onClick={() => setActiveTab("rig-breakdown")}
              className={`w-full flex items-center rounded-lg text-xs font-semibold transition-all cursor-pointer ${
                sidebarCollapsed ? "justify-center p-2.5" : "gap-3 px-3 py-2.5"
              } ${
                activeTab === "rig-breakdown"
                  ? (isLightMode ? "bg-white text-emerald-800 font-bold border border-slate-300 shadow-sm" : "bg-white/5 text-[#F5F7FA]")
                  : (isLightMode ? "text-slate-100 hover:text-white hover:bg-white/10 bg-transparent" : "text-[#9AA0A6] hover:text-[#F5F7FA] bg-transparent")
              }`}
              title="Rig Usage Breakdown"
            >
              <Layers className="w-4 h-4 text-[#10B981] shrink-0" />
              {!sidebarCollapsed && <span className="truncate">Rig Usage Breakdown</span>}
            </button>
            <button
              onClick={() => setActiveTab("gantt")}
              className={`w-full flex items-center rounded-lg text-xs font-semibold transition-all cursor-pointer ${
                sidebarCollapsed ? "justify-center p-2.5" : "gap-3 px-3 py-2.5"
              } ${
                activeTab === "gantt"
                  ? (isLightMode ? "bg-white text-amber-800 font-bold border border-slate-300 shadow-sm" : "bg-white/5 text-[#F5F7FA]")
                  : (isLightMode ? "text-slate-100 hover:text-white hover:bg-white/10 bg-transparent" : "text-[#9AA0A6] hover:text-[#F5F7FA] bg-transparent")
              }`}
              title="Rig & Well Gantt Schedule"
            >
              <CalendarRange className="w-4 h-4 text-amber-500 shrink-0" />
              {!sidebarCollapsed && <span className="truncate">Gantt Chart (Rigs & Wells)</span>}
            </button>
          </nav>
        </div>

        {/* Sidebar Footer Details */}
        <div className={`pt-4 border-t text-[11px] font-mono leading-relaxed ${
          isLightMode ? "border-slate-200 text-slate-500" : "border-white/5 text-[#9AA0A6]"
        } ${sidebarCollapsed ? "flex flex-col items-center gap-3 justify-center" : ""}`}>
          {sidebarCollapsed ? (
            <>
              <button
                onClick={handleSignOut}
                className="p-1.5 rounded-lg text-rose-500 hover:bg-rose-100 hover:text-rose-700 transition-all cursor-pointer"
                title="Sign Out Session"
              >
                <LogOut className="w-4 h-4" />
              </button>
            </>
          ) : (
            <>
              {/* System Operational Status */}

              <div className="flex items-center justify-between text-[11px]">
                <div className="text-[#10B981] flex items-center gap-1.5 font-sans font-semibold truncate">
                  <span className="w-1.5 h-1.5 rounded-full bg-[#10B981] animate-pulse shrink-0" />
                  <span className="truncate">System Operational</span>
                </div>
              </div>

              <button
                onClick={handleSignOut}
                className={`mt-2 w-full flex items-center gap-2.5 px-2.5 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer font-sans border ${
                  isLightMode
                    ? "text-rose-600 bg-rose-50 border-rose-200 hover:bg-rose-100"
                    : "text-rose-400 border-transparent hover:bg-rose-950/20 hover:text-rose-300 hover:border-rose-900/30"
                }`}
              >
                <LogOut className="w-4 h-4 shrink-0" />
                <span>Sign Out Session</span>
              </button>
            </>
          )}
        </div>
      </aside>

      {/* Right Column Layout */}
      <div className={`flex-1 flex flex-col min-h-screen transition-colors duration-200 ${isLightMode ? "bg-[#525252]" : "bg-[#0B0B0D]"}`}>
        
        {/* Top Bar Navigation & Status */}
        <header className={`h-12 border-b flex items-center justify-between px-7 shrink-0 sticky top-0 z-40 transition-colors ${
          isLightMode ? "bg-[#5C5C5C]/95 backdrop-blur border-[#3F3F3F] text-slate-100 shadow-xs" : "bg-[#0B0B0D] border-white/5 text-[#9AA0A6]"
        }`} id="main-header">
          <div className={`text-xs font-black font-sans uppercase tracking-wider ${
            isLightMode ? "text-black" : "text-white"
          }`}>
            M / LWD TOOLS DASHBOARD
          </div>
          
          <div className="flex items-center gap-2">
            {/* Theme Toggle Button (Sun / Moon) */}
            <button
              type="button"
              onClick={toggleTheme}
              className={`flex items-center gap-2 px-3 py-1.5 h-8 rounded-lg text-xs font-bold border transition-all cursor-pointer shadow-sm ${
                isLightMode
                  ? "bg-slate-900 text-white border-slate-900 hover:bg-slate-800"
                  : "bg-amber-500/10 text-amber-300 border-amber-500/30 hover:bg-amber-500/20"
              }`}
              id="header-theme-toggle-btn"
              title={isLightMode ? "Switch to Dark Theme" : "Switch to Light Theme"}
            >
              {isLightMode ? (
                <>
                  <Moon className="w-3.5 h-3.5 text-amber-300 shrink-0" />
                  <span className="font-sans">Dark Mode</span>
                </>
              ) : (
                <>
                  <Sun className="w-3.5 h-3.5 text-amber-400 shrink-0" />
                  <span className="font-sans text-amber-300">Light Mode</span>
                </>
              )}
            </button>

            {/* Local Live-pushed Excel Status Indicator */}
            {localLiveSyncLastUpdated && (
              <div className={`hidden sm:flex items-center gap-1.5 px-2.5 py-1 h-8 rounded-lg text-[10px] font-bold ${
                isLightMode
                  ? "bg-emerald-50 border border-emerald-200 text-emerald-700"
                  : "bg-emerald-950/20 border border-emerald-900/30 text-emerald-400"
              }`}>
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse shrink-0" />
                <span>Live Sync: Connected ({localLiveSyncLastUpdated})</span>
              </div>
            )}

            {/* Admin Console & Active Users Indicator */}
            {currentUser.role === "admin" && (
              <div className="flex items-center gap-2">
                <div className="relative">
                  <button
                    type="button"
                    onClick={() => setShowActiveUsersPopover((prev) => !prev)}
                    className={`flex items-center gap-2 px-3.5 py-1.5 h-8 rounded-lg text-xs transition-all cursor-pointer shadow-md border ${
                      isLightMode
                        ? "bg-emerald-600 border-emerald-500 text-white hover:bg-emerald-700 shadow-emerald-600/30"
                        : "bg-emerald-500/20 border-2 border-emerald-400 text-emerald-300 hover:bg-emerald-500/30 shadow-emerald-500/25 ring-1 ring-emerald-400/50"
                    }`}
                    title="Admin Only: Currently Active Users on Dashboard"
                  >
                    <span className="relative flex h-2.5 w-2.5 shrink-0">
                      <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-80"></span>
                      <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-400"></span>
                    </span>
                    <Users className={`w-4 h-4 shrink-0 ${isLightMode ? "text-emerald-100" : "text-emerald-300"}`} />
                    <span className="font-sans font-normal tracking-wide">{activeUserCount} ACTIVE ONLINE</span>
                    <ChevronDown className={`w-3.5 h-3.5 ${isLightMode ? "text-emerald-100" : "text-emerald-400"}`} />
                  </button>

                  {showActiveUsersPopover && (
                    <div className={`absolute right-0 mt-2 w-80 sm:w-[380px] p-3.5 rounded-2xl border shadow-2xl z-50 text-xs ${
                      isLightMode
                        ? "bg-white border-slate-200 text-slate-800"
                        : "bg-[#121620] border-slate-700 text-slate-100"
                    }`}>
                      <div className={`flex items-center justify-between border-b pb-2 mb-2.5 ${
                        isLightMode ? "border-slate-100" : "border-slate-800"
                      }`}>
                        <div className="flex items-center gap-2">
                          <span className="relative flex h-2.5 w-2.5 shrink-0">
                            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-80"></span>
                            <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-400"></span>
                          </span>
                          <span className="font-extrabold font-sans text-xs uppercase tracking-wider flex items-center gap-1.5">
                            <Users className="w-4 h-4 text-emerald-500 shrink-0" />
                            Active People & Links
                          </span>
                        </div>
                        <span className="px-2.5 py-0.5 rounded-full text-[10px] bg-emerald-500/10 text-emerald-400 border border-emerald-500/30 font-mono font-bold">
                          {activeUserCount} Active Online
                        </span>
                      </div>

                      <div className="max-h-72 overflow-y-auto space-y-2 pr-1 custom-scrollbar">
                        {activeUsersList.length === 0 ? (
                          <div className="text-slate-400 text-[11px] py-2 text-center">1 active user online (You)</div>
                        ) : (
                          activeUsersList.map((usr, i) => (
                            <div key={usr.sessionId || i} className={`p-2.5 rounded-xl border transition-all ${
                              isLightMode ? "bg-slate-50 border-slate-200 hover:border-emerald-300" : "bg-slate-900/90 border-slate-800 hover:border-emerald-500/40"
                            }`}>
                              <div className="flex items-center justify-between gap-2 mb-1.5">
                                <div className="flex items-center gap-2 truncate min-w-0">
                                  <span className="w-2 h-2 rounded-full bg-emerald-500 shrink-0 animate-pulse" />
                                  <span className="font-bold text-xs truncate">{usr.name || usr.username || "User"}</span>
                                  <span className="text-[10px] text-slate-400 truncate">@{usr.username || "user"}</span>
                                </div>
                                <span className={`text-[9px] px-2 py-0.5 rounded-md font-mono font-bold uppercase shrink-0 ${
                                  usr.role === "admin"
                                    ? "bg-purple-950/90 text-purple-300 border border-purple-800"
                                    : usr.role === "editor"
                                    ? "bg-blue-950/90 text-blue-300 border border-blue-800"
                                    : "bg-amber-950/90 text-amber-300 border border-amber-800"
                                }`}>
                                  {usr.role || "viewer"}
                                </span>
                              </div>

                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-1.5 text-[10px] mt-2 pt-2 border-t border-slate-800/60">
                                <div className={`flex items-center gap-1.5 font-mono px-2 py-1 rounded border ${
                                  isLightMode ? "bg-slate-100 text-emerald-700 border-slate-200" : "bg-slate-950/60 text-emerald-400 border-emerald-500/20"
                                }`}>
                                  <Globe className="w-3 h-3 text-emerald-400 shrink-0" />
                                  <span className="truncate font-bold">IP: {usr.ip || "Direct Connected"}</span>
                                </div>

                                <button
                                  type="button"
                                  onClick={() => {
                                    if (usr.activeTab) setActiveTab(usr.activeTab as any);
                                  }}
                                  className={`flex items-center gap-1.5 px-2 py-1 rounded border text-left truncate cursor-pointer transition-colors ${
                                    isLightMode
                                      ? "bg-blue-50 border-blue-200 text-blue-700 hover:bg-blue-100"
                                      : "bg-blue-950/50 border-blue-500/30 text-blue-300 hover:bg-blue-900/60"
                                  }`}
                                  title={`Click to navigate to ${usr.activeTabName || usr.activeTab}`}
                                >
                                  <Link2 className="w-3 h-3 text-blue-400 shrink-0" />
                                  <span className="truncate font-semibold">{usr.activeTabName || usr.activeTab || "Logs"}</span>
                                </button>
                              </div>
                            </div>
                          ))
                        )}
                      </div>
                    </div>
                  )}
                </div>

                <button
                  type="button"
                  onClick={() => setIsAdminPortalOpen(true)}
                  className={`flex items-center gap-1.5 px-3 py-1.5 h-8 rounded-lg text-xs font-bold transition-all cursor-pointer shadow-xs ${
                    isLightMode
                      ? "bg-purple-50 border border-purple-200 text-purple-700 hover:bg-purple-100"
                      : "bg-purple-600/20 border border-purple-500/40 text-purple-300 hover:bg-purple-600/30 shadow-purple-600/10"
                  }`}
                  title="Open Admin Permission & Request Manager"
                >
                  <Shield className={`w-3.5 h-3.5 ${isLightMode ? "text-purple-600" : "text-purple-400"}`} />
                  <span className="hidden sm:inline font-sans">Admin Console</span>
                </button>
              </div>
            )}

            {currentUser?.role !== "admin" && (
              <div className={`flex items-center gap-2 px-3 py-1.5 h-8 rounded-lg text-xs font-semibold border ${
                isLightMode
                  ? "bg-emerald-50 text-emerald-800 border-emerald-200"
                  : "bg-emerald-950/30 text-emerald-300 border-emerald-500/30"
              }`} title="Data changes automatically sync in real-time across all sessions">
                <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse shrink-0" />
                <span className="font-sans">Live Auto-Synced</span>
              </div>
            )}

            {activeTab === "logs" && currentUser?.role === "admin" && (
              <button
                type="button"
                onClick={handleUpdateFromExcelEditor}
                className="flex items-center gap-1.5 px-3 py-1.5 h-8 rounded-lg text-xs font-bold bg-blue-600 hover:bg-blue-500 text-white transition-all cursor-pointer shadow-md shadow-blue-600/20 border border-blue-400/20"
                id="header-update-from-excel-editor-btn"
                title="Connect entire dashboard to the live Excel Sheet Editor"
              >
                <RefreshCw className="w-3.5 h-3.5 text-blue-100" />
                <span className="hidden lg:inline font-sans">Update from Excel Sheet Editor</span>
                <span className="lg:hidden font-sans">Update Sync</span>
              </button>
            )}

            {currentUser?.role === "admin" && (
              <button
                onClick={() => setIsUploadOpen(true)}
                className={`flex items-center gap-1.5 px-3 py-1.5 h-8 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                  isLightMode
                    ? "bg-white text-blue-600 border border-slate-200 hover:bg-slate-50 shadow-xs"
                    : "bg-[#1A1D21] text-[#3B82F6] border border-white/10 hover:bg-[#3B82F6]/10 hover:text-[#3B82F6]"
                }`}
                id="header-excel-sync-btn"
              >
                <FileSpreadsheet className="w-3.5 h-3.5" />
                <span>Upload & Sync</span>
              </button>
            )}

            {data.length > 0 && currentUser?.role === "admin" && (
              <button
                onClick={handleClearAllData}
                className={`flex items-center gap-1.5 px-3 py-1.5 h-8 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                  isLightMode
                    ? "bg-rose-50 text-rose-700 border border-rose-200 hover:bg-rose-100"
                    : "bg-[#1A1D21] text-rose-400 border border-white/10 hover:bg-rose-950/20 hover:border-rose-900/40"
                }`}
                id="header-clear-all-btn"
                title="Clear all uploaded data and start fresh"
              >
                <Trash2 className="w-3.5 h-3.5 text-rose-500" />
                <span>Clear All Logs</span>
              </button>
            )}
          </div>
        </header>

        {/* Primary Content Wrapper with 28px padding on all sides */}
        <main className="flex-1 p-7 overflow-y-auto" id="dashboard-main-content">
        <AnimatePresence mode="wait">
          {/* Tab 1: Operational Logs Workspace */}
          {activeTab === "logs" && (
            <motion.div
              key="logs"
              initial={{ opacity: 0, y: 6 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -6 }}
              transition={{ duration: 0.15 }}
              className="space-y-6"
            >
              {/* Top Operations Panel: Global Specifications filtering */}
              <BorderGlow
                borderRadius={16}
                backgroundColor={isLightMode ? "#FFFFFF" : "#0F1420"}
                glowColor={isLightMode ? "220 80 50" : "210 100 70"}
                glowIntensity={isLightMode ? 0.6 : 1.2}
                edgeSensitivity={25}
                className="w-full relative z-30"
              >
                <div className={`p-5 flex flex-col justify-between ${isLightMode ? "text-slate-800" : "text-[#F5F7FA]"}`} id="specifications-filters-panel">
                  <div>
                    <div className="flex items-center justify-between gap-4 mb-4">
                      <div className="flex items-center gap-2">
                        <SlidersHorizontal className={`w-4 h-4 ${isLightMode ? "text-blue-600" : "text-[#3B82F6]"}`} />
                        <h3 className={`text-sm font-semibold font-sans ${isLightMode ? "text-slate-950" : "text-[#F5F7FA]"}`}>Filters</h3>
                      </div>

                      {isFiltered && (
                        <button
                          onClick={resetFilters}
                          className="px-3.5 py-1.5 text-xs font-bold bg-rose-800 text-white hover:bg-rose-900 border border-rose-900 rounded-lg shadow-md transition-all cursor-pointer flex items-center gap-1.5"
                          id="reset-filters-link"
                        >
                          <RotateCcw className="w-3.5 h-3.5" />
                          <span>Reset Filters</span>
                        </button>
                      )}
                    </div>

                    <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
                      {/* Client Multi-select Dropdown */}
                      <div className="relative">
                        <label className={`block text-xs font-medium mb-1.5 font-sans ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"}`}>Client</label>
                        <button
                          type="button"
                          onClick={() => setIsClientDropdownOpen(!isClientDropdownOpen)}
                          className={`w-full h-9 px-3 text-[13px] border rounded-lg focus:outline-none font-medium transition-all cursor-pointer flex items-center justify-between text-left ${
                            isLightMode 
                              ? "bg-slate-50/90 border-slate-300 text-slate-800 focus:border-blue-500 hover:bg-slate-100/90 shadow-2xs hover:shadow-xs" 
                              : "bg-[#1A1D21] border-white/10 text-[#E0E0E0] focus:border-[#3B82F6] hover:bg-[#1f2329]"
                          }`}
                          id="client-spec-dropdown-trigger"
                        >
                          <span className="truncate">
                            {selectedClients.length === 0 
                              ? "All Clients" 
                              : selectedClients.length === 1 
                                ? selectedClients[0] 
                                : `${selectedClients.length} selected`}
                          </span>
                          <ChevronDown className={`w-4 h-4 text-slate-400 transition-transform ${isClientDropdownOpen ? "rotate-180" : ""}`} />
                        </button>

                        {isClientDropdownOpen && (
                          <>
                            <div 
                              className="fixed inset-0 z-40 cursor-default" 
                              onClick={() => setIsClientDropdownOpen(false)} 
                            />
                            
                            <div className={`absolute left-0 right-0 mt-1 border rounded-lg shadow-xl z-50 py-1.5 max-h-60 overflow-y-auto select-none ${
                              isLightMode 
                                ? "bg-white border-slate-200 text-slate-800 shadow-lg" 
                                : "bg-[#1A1D21] border-white/10 text-white shadow-xl"
                            }`}>
                              <div
                                onClick={() => {
                                  setSelectedClients([]);
                                }}
                                className={`flex items-center gap-2 px-3 py-1.5 cursor-pointer text-xs font-semibold ${isLightMode ? "hover:bg-slate-100" : "hover:bg-white/5"}`}
                              >
                                <div className={`w-3.5 h-3.5 rounded border flex items-center justify-center shrink-0 transition-all ${
                                  selectedClients.length === 0 
                                    ? "bg-blue-600 border-blue-600 text-white" 
                                    : (isLightMode ? "border-slate-300" : "border-slate-600")
                                }`}>
                                  {selectedClients.length === 0 && <Check className="w-2.5 h-2.5 stroke-[3]" />}
                                </div>
                                <span className={selectedClients.length === 0 ? "text-blue-600" : (isLightMode ? "text-slate-600" : "text-slate-400")}>
                                  All Clients
                                </span>
                              </div>

                              <div className={`h-px my-1 ${isLightMode ? "bg-slate-100" : "bg-white/5"}`} />

                              {clientOptions.map((cli) => {
                                const isSelected = selectedClients.includes(cli);
                                return (
                                  <div
                                    key={cli}
                                    onClick={() => {
                                      if (isSelected) {
                                        setSelectedClients(selectedClients.filter(item => item !== cli));
                                      } else {
                                        setSelectedClients([...selectedClients, cli]);
                                      }
                                    }}
                                    className={`flex items-center gap-2 px-3 py-1.5 cursor-pointer text-xs ${isLightMode ? "hover:bg-slate-50" : "hover:bg-white/5"}`}
                                  >
                                    <div className={`w-3.5 h-3.5 rounded border flex items-center justify-center shrink-0 transition-all ${
                                      isSelected 
                                        ? "bg-blue-600 border-blue-600 text-white" 
                                        : (isLightMode ? "border-slate-300" : "border-slate-600")
                                    }`}>
                                      {isSelected && <Check className="w-2.5 h-2.5 stroke-[3]" />}
                                    </div>
                                    <span className={isSelected ? (isLightMode ? "text-slate-900 font-semibold" : "text-white font-medium") : (isLightMode ? "text-slate-600" : "text-slate-400")}>
                                      {cli}
                                    </span>
                                  </div>
                                );
                              })}
                            </div>
                          </>
                        )}
                      </div>

                      {/* Rig Name Multi-select Dropdown */}
                      <div className="relative">
                        <label className={`block text-xs font-medium mb-1.5 font-sans ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"}`}>Rig Name</label>
                        <button
                          type="button"
                          onClick={() => setIsRigDropdownOpen(!isRigDropdownOpen)}
                          className={`w-full h-9 px-3 text-[13px] border rounded-lg focus:outline-none font-medium transition-all cursor-pointer flex items-center justify-between text-left ${
                            isLightMode 
                              ? "bg-slate-50/90 border-slate-300 text-slate-800 focus:border-blue-500 hover:bg-slate-100/90 shadow-2xs hover:shadow-xs" 
                              : "bg-[#1A1D21] border-white/10 text-[#E0E0E0] focus:border-[#3B82F6] hover:bg-[#1f2329]"
                          }`}
                          id="rig-spec-dropdown-trigger"
                        >
                          <span className="truncate">
                            {selectedRigs.length === 0 
                              ? "All Rigs" 
                              : selectedRigs.length === 1 
                                ? selectedRigs[0] 
                                : `${selectedRigs.length} selected`}
                          </span>
                          <ChevronDown className={`w-4 h-4 text-slate-400 transition-transform ${isRigDropdownOpen ? "rotate-180" : ""}`} />
                        </button>

                        {isRigDropdownOpen && (
                          <>
                            <div 
                              className="fixed inset-0 z-40 cursor-default" 
                              onClick={() => setIsRigDropdownOpen(false)} 
                            />
                            
                            <div className={`absolute left-0 right-0 mt-1 border rounded-lg shadow-xl z-50 py-1.5 max-h-60 overflow-y-auto select-none ${
                              isLightMode 
                                ? "bg-white border-slate-200 text-slate-800 shadow-lg" 
                                : "bg-[#1A1D21] border-white/10 text-white shadow-xl"
                            }`}>
                              <div
                                onClick={() => {
                                  setSelectedRigs([]);
                                }}
                                className={`flex items-center gap-2 px-3 py-1.5 cursor-pointer text-xs font-semibold ${isLightMode ? "hover:bg-slate-100" : "hover:bg-white/5"}`}
                              >
                                <div className={`w-3.5 h-3.5 rounded border flex items-center justify-center shrink-0 transition-all ${
                                  selectedRigs.length === 0 
                                    ? "bg-blue-600 border-blue-600 text-white" 
                                    : (isLightMode ? "border-slate-300" : "border-slate-600")
                                }`}>
                                  {selectedRigs.length === 0 && <Check className="w-2.5 h-2.5 stroke-[3]" />}
                                </div>
                                <span className={selectedRigs.length === 0 ? "text-blue-600" : (isLightMode ? "text-slate-600" : "text-slate-400")}>
                                  All Rigs
                                </span>
                              </div>

                              <div className={`h-px my-1 ${isLightMode ? "bg-slate-100" : "bg-white/5"}`} />

                              {rigOptions.map((rig) => {
                                const isSelected = selectedRigs.includes(rig);
                                return (
                                  <div
                                    key={rig}
                                    onClick={() => {
                                      if (isSelected) {
                                        setSelectedRigs(selectedRigs.filter(item => item !== rig));
                                      } else {
                                        setSelectedRigs([...selectedRigs, rig]);
                                      }
                                    }}
                                    className={`flex items-center gap-2 px-3 py-1.5 cursor-pointer text-xs ${isLightMode ? "hover:bg-slate-50" : "hover:bg-white/5"}`}
                                  >
                                    <div className={`w-3.5 h-3.5 rounded border flex items-center justify-center shrink-0 transition-all ${
                                      isSelected 
                                        ? "bg-blue-600 border-blue-600 text-white" 
                                        : (isLightMode ? "border-slate-300" : "border-slate-600")
                                    }`}>
                                      {isSelected && <Check className="w-2.5 h-2.5 stroke-[3]" />}
                                    </div>
                                    <span className={isSelected ? (isLightMode ? "text-slate-900 font-semibold" : "text-white font-medium") : (isLightMode ? "text-slate-600" : "text-slate-400")}>
                                      {rig}
                                    </span>
                                  </div>
                                );
                              })}
                            </div>
                          </>
                        )}
                      </div>

                      {/* Tool Sizes Multi-select Dropdown */}
                      <div className="relative">
                        <label className={`block text-xs font-medium mb-1.5 font-sans ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"}`}>Tool Sizes</label>
                        <button
                          type="button"
                          onClick={() => setIsSectionDropdownOpen(!isSectionDropdownOpen)}
                          className={`w-full h-9 px-3 text-[13px] border rounded-lg focus:outline-none font-medium transition-all cursor-pointer flex items-center justify-between text-left ${
                            isLightMode 
                              ? "bg-slate-50/90 border-slate-300 text-slate-800 focus:border-blue-500 hover:bg-slate-100/90 shadow-2xs hover:shadow-xs" 
                              : "bg-[#1A1D21] border-white/10 text-[#E0E0E0] focus:border-[#3B82F6] hover:bg-[#1f2329]"
                          }`}
                          id="section-spec-dropdown-trigger"
                        >
                          <span className="truncate">
                            {selectedSection.length === 0 
                              ? "All Tool Sizes" 
                              : selectedSection.length === 1 
                                ? selectedSection[0] 
                                : `${selectedSection.length} selected`}
                          </span>
                          <ChevronDown className={`w-4 h-4 text-slate-400 transition-transform ${isSectionDropdownOpen ? "rotate-180" : ""}`} />
                        </button>

                        {isSectionDropdownOpen && (
                          <>
                            {/* Backdrop to close dropdown */}
                            <div 
                              className="fixed inset-0 z-40 cursor-default" 
                              onClick={() => setIsSectionDropdownOpen(false)} 
                            />
                            
                            {/* Dropdown Options List */}
                            <div className={`absolute left-0 right-0 mt-1 border rounded-lg shadow-xl z-50 py-1.5 max-h-60 overflow-y-auto select-none ${
                              isLightMode 
                                ? "bg-white border-slate-200 text-slate-800 shadow-lg" 
                                : "bg-[#1A1D21] border-white/10 text-white shadow-xl"
                            }`}>
                              {/* "All" Select / Clear Button option */}
                              <div
                                onClick={() => {
                                  setSelectedSection([]);
                                }}
                                className={`flex items-center gap-2 px-3 py-1.5 cursor-pointer text-xs font-semibold ${isLightMode ? "hover:bg-slate-100" : "hover:bg-white/5"}`}
                              >
                                <div className={`w-3.5 h-3.5 rounded border flex items-center justify-center shrink-0 transition-all ${
                                  selectedSection.length === 0 
                                    ? "bg-blue-600 border-blue-600 text-white" 
                                    : (isLightMode ? "border-slate-300" : "border-slate-600")
                                }`}>
                                  {selectedSection.length === 0 && <Check className="w-2.5 h-2.5 stroke-[3]" />}
                                </div>
                                <span className={selectedSection.length === 0 ? "text-blue-600" : (isLightMode ? "text-slate-600" : "text-slate-400")}>
                                  All Tool Sizes
                                </span>
                              </div>

                              <div className={`h-px my-1 ${isLightMode ? "bg-slate-100" : "bg-white/5"}`} />

                              {/* Section Options */}
                              {sectionOptions.map((sec) => {
                                const isSelected = selectedSection.includes(sec);
                                return (
                                  <div
                                    key={sec}
                                    onClick={() => {
                                      if (isSelected) {
                                        setSelectedSection(selectedSection.filter(item => item !== sec));
                                      } else {
                                        setSelectedSection([...selectedSection, sec]);
                                      }
                                    }}
                                    className={`flex items-center gap-2 px-3 py-1.5 cursor-pointer text-xs ${isLightMode ? "hover:bg-slate-50" : "hover:bg-white/5"}`}
                                  >
                                    <div className={`w-3.5 h-3.5 rounded border flex items-center justify-center shrink-0 transition-all ${
                                      isSelected 
                                        ? "bg-blue-600 border-blue-600 text-white" 
                                        : (isLightMode ? "border-slate-300" : "border-slate-600")
                                    }`}>
                                      {isSelected && <Check className="w-2.5 h-2.5 stroke-[3]" />}
                                    </div>
                                    <span className={isSelected ? (isLightMode ? "text-slate-900 font-semibold" : "text-white font-medium") : (isLightMode ? "text-slate-600" : "text-slate-400")}>
                                      {sec}
                                    </span>
                                  </div>
                                );
                              })}
                            </div>
                          </>
                        )}
                      </div>

                      {/* Tool Class Multi-select Dropdown */}
                      <div className="relative">
                        <label className={`block text-xs font-medium mb-1.5 font-sans ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"}`}>Tools</label>
                        <button
                          type="button"
                          onClick={() => setIsToolDropdownOpen(!isToolDropdownOpen)}
                          className={`w-full h-9 px-3 text-[13px] border rounded-lg focus:outline-none font-medium transition-all cursor-pointer flex items-center justify-between text-left ${
                            isLightMode 
                              ? "bg-slate-50/90 border-slate-300 text-slate-800 focus:border-blue-500 hover:bg-slate-100/90 shadow-2xs hover:shadow-xs" 
                              : "bg-[#1A1D21] border-white/10 text-[#E0E0E0] focus:border-[#3B82F6] hover:bg-[#1f2329]"
                          }`}
                          id="tool-spec-dropdown-trigger"
                        >
                          <span className="truncate">
                            {selectedTools.length === 0 
                              ? "All Tools" 
                              : selectedTools.length === 1 
                                ? selectedTools[0].replace(/\s*\(.*?\)\s*/g, " ").trim() 
                                : `${selectedTools.length} selected`}
                          </span>
                          <ChevronDown className={`w-4 h-4 text-slate-400 transition-transform ${isToolDropdownOpen ? "rotate-180" : ""}`} />
                        </button>

                        {isToolDropdownOpen && (
                          <>
                            <div 
                              className="fixed inset-0 z-40 cursor-default" 
                              onClick={() => setIsToolDropdownOpen(false)} 
                            />
                            
                            <div className={`absolute left-0 right-0 mt-1 border rounded-lg shadow-xl z-50 py-1.5 max-h-60 overflow-y-auto select-none ${
                              isLightMode 
                                ? "bg-white border-slate-200 text-slate-800 shadow-lg" 
                                : "bg-[#1A1D21] border-white/10 text-white shadow-xl"
                            }`}>
                              <div
                                onClick={() => {
                                  setSelectedTools([]);
                                }}
                                className={`flex items-center gap-2 px-3 py-1.5 cursor-pointer text-xs font-semibold ${isLightMode ? "hover:bg-slate-100" : "hover:bg-white/5"}`}
                              >
                                <div className={`w-3.5 h-3.5 rounded border flex items-center justify-center shrink-0 transition-all ${
                                  selectedTools.length === 0 
                                    ? "bg-blue-600 border-blue-600 text-white" 
                                    : (isLightMode ? "border-slate-300" : "border-slate-600")
                                }`}>
                                  {selectedTools.length === 0 && <Check className="w-2.5 h-2.5 stroke-[3]" />}
                                </div>
                                <span className={selectedTools.length === 0 ? "text-blue-600" : (isLightMode ? "text-slate-600" : "text-slate-400")}>
                                  All Tools
                                </span>
                              </div>

                              <div className={`h-px my-1 ${isLightMode ? "bg-slate-100" : "bg-white/5"}`} />

                              {uniqueFullNames.map((toolName) => {
                                const isSelected = selectedTools.includes(toolName);
                                const displayLabel = toolName.replace(/\s*\(.*?\)\s*/g, " ").trim();
                                return (
                                  <div
                                    key={toolName}
                                    onClick={() => {
                                      if (isSelected) {
                                        setSelectedTools(selectedTools.filter(item => item !== toolName));
                                      } else {
                                        setSelectedTools([...selectedTools, toolName]);
                                      }
                                    }}
                                    className={`flex items-center gap-2 px-3 py-1.5 cursor-pointer text-xs ${isLightMode ? "hover:bg-slate-50" : "hover:bg-white/5"}`}
                                  >
                                    <div className={`w-3.5 h-3.5 rounded border flex items-center justify-center shrink-0 transition-all ${
                                      isSelected 
                                        ? "bg-blue-600 border-blue-600 text-white" 
                                        : (isLightMode ? "border-slate-300" : "border-slate-600")
                                    }`}>
                                      {isSelected && <Check className="w-2.5 h-2.5 stroke-[3]" />}
                                    </div>
                                    <span className={isSelected ? (isLightMode ? "text-slate-900 font-semibold" : "text-white font-medium") : (isLightMode ? "text-slate-600" : "text-slate-400")}>
                                      {displayLabel}
                                    </span>
                                  </div>
                                );
                              })}
                            </div>
                          </>
                        )}
                      </div>

                      {/* Shop Status Multi-select Dropdown */}
                      <div className="relative">
                        <label className={`block text-xs font-medium mb-1.5 font-sans ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"}`}>Shop Status</label>
                        <button
                          type="button"
                          onClick={() => setIsShopDateDropdownOpen(!isShopDateDropdownOpen)}
                          className={`w-full h-9 px-3 text-[13px] border rounded-lg focus:outline-none font-medium transition-all cursor-pointer flex items-center justify-between text-left ${
                            isLightMode 
                              ? "bg-slate-50/90 border-slate-300 text-slate-800 focus:border-blue-500 hover:bg-slate-100/90 shadow-2xs hover:shadow-xs" 
                              : "bg-[#1A1D21] border-white/10 text-[#E0E0E0] focus:border-[#3B82F6] hover:bg-[#1f2329]"
                          }`}
                          id="shopdate-spec-dropdown-trigger"
                        >
                          <span className="truncate">
                            {selectedShopDates.length === 0 
                              ? "All Statuses" 
                              : selectedShopDates.length === 1 
                                ? (selectedShopDates[0] === "Delivered" ? "Delivered" : "Pending / In-shop") 
                                : `${selectedShopDates.length} selected`}
                          </span>
                          <ChevronDown className={`w-4 h-4 text-slate-400 transition-transform ${isShopDateDropdownOpen ? "rotate-180" : ""}`} />
                        </button>

                        {isShopDateDropdownOpen && (
                          <>
                            <div 
                              className="fixed inset-0 z-40 cursor-default" 
                              onClick={() => setIsShopDateDropdownOpen(false)} 
                            />
                            
                            <div className={`absolute left-0 right-0 mt-1 border rounded-lg shadow-xl z-50 py-1.5 max-h-60 overflow-y-auto select-none ${
                              isLightMode 
                                ? "bg-white border-slate-200 text-slate-800 shadow-lg" 
                                : "bg-[#1A1D21] border-white/10 text-white shadow-xl"
                            }`}>
                              <div
                                onClick={() => {
                                  setSelectedShopDates([]);
                                }}
                                className={`flex items-center gap-2 px-3 py-1.5 cursor-pointer text-xs font-semibold ${isLightMode ? "hover:bg-slate-100" : "hover:bg-white/5"}`}
                              >
                                <div className={`w-3.5 h-3.5 rounded border flex items-center justify-center shrink-0 transition-all ${
                                  selectedShopDates.length === 0 
                                    ? "bg-blue-600 border-blue-600 text-white" 
                                    : (isLightMode ? "border-slate-300" : "border-slate-600")
                                }`}>
                                  {selectedShopDates.length === 0 && <Check className="w-2.5 h-2.5 stroke-[3]" />}
                                </div>
                                <span className={selectedShopDates.length === 0 ? "text-blue-600" : (isLightMode ? "text-slate-600" : "text-slate-400")}>
                                  All Statuses
                                </span>
                              </div>

                              <div className={`h-px my-1 ${isLightMode ? "bg-slate-100" : "bg-white/5"}`} />

                              {[
                                { value: "Delivered", label: "Delivered" },
                                { value: "Pending", label: "Pending / In-shop" }
                              ].map((opt) => {
                                const isSelected = selectedShopDates.includes(opt.value);
                                return (
                                  <div
                                    key={opt.value}
                                    onClick={() => {
                                      if (isSelected) {
                                        setSelectedShopDates(selectedShopDates.filter(item => item !== opt.value));
                                      } else {
                                        setSelectedShopDates([...selectedShopDates, opt.value]);
                                      }
                                    }}
                                    className={`flex items-center gap-2 px-3 py-1.5 cursor-pointer text-xs ${isLightMode ? "hover:bg-slate-50" : "hover:bg-white/5"}`}
                                  >
                                    <div className={`w-3.5 h-3.5 rounded border flex items-center justify-center shrink-0 transition-all ${
                                      isSelected 
                                        ? "bg-blue-600 border-blue-600 text-white" 
                                        : (isLightMode ? "border-slate-300" : "border-slate-600")
                                    }`}>
                                      {isSelected && <Check className="w-2.5 h-2.5 stroke-[3]" />}
                                    </div>
                                    <span className={isSelected ? (isLightMode ? "text-slate-900 font-semibold" : "text-white font-medium") : (isLightMode ? "text-slate-600" : "text-slate-400")}>
                                      {opt.label}
                                    </span>
                                  </div>
                                );
                              })}
                            </div>
                          </>
                        )}
                      </div>
                    </div>

                    {/* Date Range Filters Row */}
                    <div className={`mt-4 pt-4 border-t flex flex-col md:flex-row md:items-center justify-between gap-4 ${isLightMode ? "border-slate-100" : "border-white/5"}`}>
                      <div className="flex flex-col sm:flex-row sm:items-center gap-3 w-full md:w-auto">
                        <span className={`text-xs font-medium font-sans ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"}`}>Date Range Filter:</span>
                        <div className="flex items-center gap-2.5 w-full sm:w-auto">
                          <div className="relative flex-1 sm:flex-none">
                            <input
                              type="date"
                              value={startDate}
                              min={dateRangeBounds.min}
                              max={dateRangeBounds.max}
                              onChange={(e) => setStartDate(e.target.value)}
                              className={`w-full sm:w-auto h-9 px-3 text-[13px] border rounded-lg focus:outline-none font-medium transition-all cursor-pointer ${
                                isLightMode
                                  ? "bg-slate-50/90 border-slate-300 text-slate-800 focus:border-blue-500 hover:bg-slate-100/90 shadow-2xs hover:shadow-xs"
                                  : "bg-[#1A1D21] border-white/10 text-[#E0E0E0] focus:border-[#3B82F6] hover:bg-[#1f2329]"
                              }`}
                              id="start-date-filter"
                            />
                          </div>
                          <span className={`text-xs font-medium ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"}`}>to</span>
                          <div className="relative flex-1 sm:flex-none">
                            <input
                              type="date"
                              value={endDate}
                              min={dateRangeBounds.min}
                              max={dateRangeBounds.max}
                              onChange={(e) => setEndDate(e.target.value)}
                              className={`w-full sm:w-auto h-9 px-3 text-[13px] border rounded-lg focus:outline-none font-medium transition-all cursor-pointer ${
                                isLightMode
                                  ? "bg-slate-50/90 border-slate-300 text-slate-800 focus:border-blue-500 hover:bg-slate-100/90 shadow-2xs hover:shadow-xs"
                                  : "bg-[#1A1D21] border-white/10 text-[#E0E0E0] focus:border-[#3B82F6] hover:bg-[#1f2329]"
                              }`}
                              id="end-date-filter"
                            />
                          </div>
                        </div>
                      </div>

                      {/* Presets - Segmented Pill Control */}
                      <div className="flex items-center gap-1.5 overflow-x-auto pb-1 md:pb-0">
                        <div className={`flex items-center border rounded-lg p-0.5 h-8 ${
                          isLightMode ? "bg-slate-100 border-slate-200" : "bg-[#1A1D21] border-white/10"
                        }`}>
                          <button
                            onClick={() => {
                              setStartDate("");
                              setEndDate("");
                            }}
                            className={`h-full px-3 text-xs font-semibold rounded-[6px] transition-all cursor-pointer whitespace-nowrap ${
                              !startDate && !endDate
                                ? (isLightMode ? "bg-white text-slate-900 shadow-sm font-bold" : "bg-[#2D3139]/50 text-[#F5F7FA] font-bold")
                                : (isLightMode ? "text-slate-500 hover:text-slate-800" : "text-[#9AA0A6] hover:text-[#F5F7FA]")
                            }`}
                          >
                            All Time
                          </button>
                          {latestDataMonths.map((m) => {
                            const isSelected = startDate === m.start && endDate === m.end;
                            return (
                              <button
                                key={m.label}
                                onClick={() => {
                                  setStartDate(m.start);
                                  setEndDate(m.end);
                                }}
                                className={`h-full px-3 text-xs font-semibold rounded-[6px] transition-all cursor-pointer whitespace-nowrap ${
                                  isSelected
                                    ? (isLightMode ? "bg-white text-slate-900 shadow-sm font-bold" : "bg-[#2D3139]/50 text-[#F5F7FA] font-bold")
                                    : (isLightMode ? "text-slate-500 hover:text-slate-800" : "text-[#9AA0A6] hover:text-[#F5F7FA]")
                                }`}
                              >
                                {m.label}
                              </button>
                            );
                          })}

                          {allDataMonths.length > 0 && (
                            <select
                              value={
                                allDataMonths.find((m) => m.start === startDate && m.end === endDate)?.key || ""
                              }
                              onChange={(e) => {
                                const selectedKey = e.target.value;
                                if (!selectedKey) return;
                                const m = allDataMonths.find((item) => item.key === selectedKey);
                                if (m) {
                                  setStartDate(m.start);
                                  setEndDate(m.end);
                                }
                              }}
                              className={`h-full px-2 text-xs font-medium rounded-[6px] border-l outline-hidden cursor-pointer transition-all ${
                                isLightMode
                                  ? "bg-slate-100 text-slate-700 border-slate-200 hover:bg-slate-200"
                                  : "bg-[#1A1D21] text-slate-300 border-white/10 hover:bg-[#252C3D]"
                              }`}
                              title="Select any month from dataset"
                            >
                              <option value="">More Months...</option>
                              {allDataMonths.map((m) => (
                                <option key={m.key} value={m.key}>
                                  {m.label}
                                </option>
                              ))}
                            </select>
                          )}
                        </div>
                        {dateRangeBounds.min && dateRangeBounds.max && (
                          <span className={`text-[10px] font-mono pl-1.5 border-l whitespace-nowrap hidden lg:inline ${
                            isLightMode ? "text-slate-400 border-slate-200" : "text-[#9AA0A6] border-white/5"
                          }`}>
                            Available Range: {dateRangeBounds.min} to {dateRangeBounds.max}
                          </span>
                        )}
                      </div>
                    </div>

                    {/* Active Filter Badges */}
                    {isFiltered && (
                      <div className={`mt-4 pt-4 border-t flex flex-wrap items-center gap-1.5 ${isLightMode ? "border-slate-100" : "border-white/5"}`}>
                        <span className={`text-xs font-medium mr-1 ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"}`}>Active Filters:</span>
                        
                        {selectedClients.map((cli) => (
                          <span key={cli} className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-medium border ${
                            isLightMode ? "bg-blue-50 text-blue-700 border-blue-100" : "bg-[#1A1D21] text-[#3B82F6] border-white/5"
                          }`}>
                            <span>Client: {cli}</span>
                            <button onClick={() => setSelectedClients(selectedClients.filter(c => c !== cli))} className={`ml-0.5 cursor-pointer text-xs font-bold font-mono ${isLightMode ? "hover:text-blue-900 text-blue-400" : "hover:text-[#F5F7FA] text-[#3B82F6]"}`}>×</button>
                          </span>
                        ))}

                        {selectedRigs.map((rig) => (
                          <span key={rig} className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-medium border ${
                            isLightMode ? "bg-blue-50 text-blue-700 border-blue-100" : "bg-[#1A1D21] text-[#3B82F6] border-white/5"
                          }`}>
                            <span>Rig: {rig}</span>
                            <button onClick={() => setSelectedRigs(selectedRigs.filter(r => r !== rig))} className={`ml-0.5 cursor-pointer text-xs font-bold font-mono ${isLightMode ? "hover:text-blue-900 text-blue-400" : "hover:text-[#F5F7FA] text-[#3B82F6]"}`}>×</button>
                          </span>
                        ))}

                        {selectedSection.map((sec) => (
                          <span key={sec} className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-medium border ${
                            isLightMode ? "bg-blue-50 text-blue-700 border-blue-100" : "bg-[#1A1D21] text-[#3B82F6] border-white/5"
                          }`}>
                            <span>Section: {sec}</span>
                            <button 
                              onClick={() => setSelectedSection(selectedSection.filter(s => s !== sec))} 
                              className={`ml-0.5 cursor-pointer text-xs font-bold font-mono ${isLightMode ? "hover:text-blue-900 text-blue-400" : "hover:text-[#F5F7FA] text-[#3B82F6]"}`}
                              title={`Remove ${sec}`}
                            >
                              ×
                            </button>
                          </span>
                        ))}

                        {selectedTools.map((tool) => (
                          <span key={tool} className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-md text-[11px] font-bold border ${
                            isLightMode ? "bg-blue-600 text-white border-blue-700 shadow-2xs" : "bg-[#1A1D21] text-[#3B82F6] border-white/5"
                          }`}>
                            <span>Tool: {tool.replace(/\s*\(.*?\)\s*/g, " ").trim()}</span>
                            <button onClick={() => setSelectedTools(selectedTools.filter(t => t !== tool))} className={`ml-0.5 cursor-pointer text-xs font-bold font-mono ${isLightMode ? "hover:text-slate-200 text-white" : "hover:text-[#F5F7FA] text-[#3B82F6]"}`}>×</button>
                          </span>
                        ))}

                        {selectedShopDates.map((status) => (
                          <span key={status} className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-medium border ${
                            isLightMode ? "bg-slate-200 text-slate-800 border-slate-300 font-semibold" : "bg-[#1A1D21] text-[#3B82F6] border-white/5"
                          }`}>
                            <span>Status: {status === "Pending" ? "Pending / In-shop" : status}</span>
                            <button onClick={() => setSelectedShopDates(selectedShopDates.filter(s => s !== status))} className={`ml-0.5 cursor-pointer text-xs font-bold font-mono ${isLightMode ? "hover:text-slate-900 text-slate-600" : "hover:text-[#F5F7FA] text-[#3B82F6]"}`}>×</button>
                          </span>
                        ))}

                        {startDate && (
                          <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-medium border ${
                            isLightMode ? "bg-slate-200 text-slate-800 border-slate-300 font-semibold" : "bg-[#1A1D21] text-[#3B82F6] border-white/5"
                          }`}>
                            <span>Start: {startDate}</span>
                            <button onClick={() => setStartDate("")} className={`ml-0.5 cursor-pointer text-xs font-bold font-mono ${isLightMode ? "hover:text-slate-900 text-slate-600" : "hover:text-[#F5F7FA] text-[#3B82F6]"}`}>×</button>
                          </span>
                        )}

                        {endDate && (
                          <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-medium border ${
                            isLightMode ? "bg-slate-200 text-slate-800 border-slate-300 font-semibold" : "bg-[#1A1D21] text-[#3B82F6] border-white/5"
                          }`}>
                            <span>End: {endDate}</span>
                            <button onClick={() => setEndDate("")} className={`ml-0.5 cursor-pointer text-xs font-bold font-mono ${isLightMode ? "hover:text-slate-900 text-slate-600" : "hover:text-[#F5F7FA] text-[#3B82F6]"}`}>×</button>
                          </span>
                        )}

                        <button
                          onClick={resetFilters}
                          className="inline-flex items-center justify-center px-3 py-1 text-[11px] font-bold bg-rose-800 text-white hover:bg-rose-900 border border-rose-900 rounded-md shadow-sm transition-all cursor-pointer"
                        >
                          Clear All
                        </button>
                      </div>
                    )}
                  </div>
                </div>

                  {/* Focus status badge if active */}
                  {focusOnMatches && highlightedRowIds && (
                    <div className="mt-3.5 flex items-center justify-between p-2 bg-amber-500/10 border border-amber-500/20 rounded-xl text-xs text-amber-300">
                      <div className="flex items-center gap-2 font-medium">
                        <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse" />
                        Focus Filter: Showing only {highlightedRowIds.length} evaluated configurations.
                      </div>
                      <button
                        onClick={() => {
                          setFocusOnMatches(false);
                          setHighlightedRowIds(null);
                        }}
                        className="text-[10px] font-bold bg-[#16181F] text-amber-300 px-2 py-1 rounded-lg border border-amber-500/20 hover:bg-[#1c1e26] transition-all cursor-pointer"
                      >
                        Reset Focus
                      </button>
                    </div>
                  )}
              </BorderGlow>

            {/* KPI Performance Metrics Dashboard & Breakdown Charts */}
            <StatsDashboard
              data={filteredData}
              abbreviations={abbreviations}
              onAbbreviationClick={handleAbbreviationClick}
              selectedTool={selectedTools.join(", ") || "All"}
              isLightMode={isLightMode}
              view="profile"
              selectedToolForRigPanel={selectedToolForRigPanel}
              onSelectToolForRigPanel={(toolName) => {
                setSelectedToolForRigPanel(toolName);
                if (toolName) {
                  setActiveTab("rig-breakdown");
                }
              }}
            />

            {/* Main Interactive Spreadsheet Grid */}
            <SpreadsheetTable
              data={filteredData}
              abbreviations={abbreviations}
              highlightedRowIds={highlightedRowIds || []}
              onAbbreviationClick={handleAbbreviationClick}
              isLightMode={isLightMode}
            />

            {/* Allocation Metrics Breakdown Panel */}
            <div className={`mt-8 border-t ${isLightMode ? "border-slate-150" : "border-[#1E293B]/40"} pt-8`}>
              <div className="mb-4">
                <h2 className={`text-sm font-bold ${isLightMode ? "text-slate-800" : "text-[#F5F7FA]"} font-sans flex items-center gap-2`}>
                  <Layers className="w-4 h-4 text-blue-500" />
                  Fleet Configuration Allocations
                </h2>
                <p className="text-[11px] text-[#9AA0A6] mt-1 font-sans">
                  Detailed distribution of specialized LWD/MWD configuration runs across active client operators and drilling rigs.
                </p>
              </div>
              <StatsDashboard
                data={filteredData}
                abbreviations={abbreviations}
                onAbbreviationClick={handleAbbreviationClick}
                selectedTool={selectedTools.join(", ") || "All"}
                isLightMode={isLightMode}
                view="allocations"
                selectedToolForRigPanel={selectedToolForRigPanel}
                onSelectToolForRigPanel={(toolName) => {
                  setSelectedToolForRigPanel(toolName);
                  if (toolName) {
                    setActiveTab("rig-breakdown");
                  }
                }}
              />
            </div>

          </motion.div>
        )}

        {/* Tab 2: Interactive Excel Sheet Editor */}
        {activeTab === "excel-editor" && (
          <motion.div
            key="excel-editor"
            initial={{ opacity: 0, y: 6 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -6 }}
            transition={{ duration: 0.15 }}
            className="space-y-6"
          >
            <ExcelGridEditor
              data={data}
              abbreviations={abbreviations}
              onUpdateData={(newData) => handleDataLoaded(newData)}
              isLightMode={isLightMode}
              userRole={currentUser.role}
              canEdit={currentUser.role === "editor" || currentUser.role === "admin"}
              onRequestAuthModal={() => setIsAuthModalOpen(true)}
            />
          </motion.div>
        )}

        {/* Tab 3: Rig Usage Breakdown Workspace */}
        {activeTab === "rig-breakdown" && (
          <motion.div
            key="rig-breakdown"
            initial={{ opacity: 0, y: 6 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -6 }}
            transition={{ duration: 0.15 }}
            className="space-y-6"
          >
            <StatsDashboard
              data={filteredData}
              abbreviations={abbreviations}
              onAbbreviationClick={handleAbbreviationClick}
              selectedTool={selectedTools.join(", ") || "All"}
              isLightMode={isLightMode}
              view="rig-breakdown"
              selectedToolForRigPanel={selectedToolForRigPanel}
              onSelectToolForRigPanel={(toolName) => {
                setSelectedToolForRigPanel(toolName);
              }}
            />
          </motion.div>
        )}

        {/* Tab 4: Rig & Well Gantt Schedule Chart */}
        {activeTab === "gantt" && (
          <motion.div
            key="gantt"
            initial={{ opacity: 0, y: 6 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -6 }}
            transition={{ duration: 0.15 }}
            className="space-y-6"
          >
            <RigWellGanttChart
              data={data}
              abbreviations={abbreviations}
              isLightMode={isLightMode}
              onSelectTool={(toolName) => {
                setSelectedToolForRigPanel(toolName);
                setActiveTab("rig-breakdown");
              }}
            />
          </motion.div>
        )}
        </AnimatePresence>

      </main>
      </div> {/* Close the Right Column Layout */}

      {/* Upload & Sync Excel Modal */}
      <AnimatePresence>
        {isUploadOpen && currentUser?.role === "admin" && (
          <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center z-50 p-4" onClick={() => setIsUploadOpen(false)}>
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 10 }}
              className="relative max-w-xl w-full bg-[#0F1420] rounded-2xl border border-slate-800/80 shadow-2xl overflow-hidden"
              onClick={(e) => e.stopPropagation()}
            >
              {/* Header inside modal */}
              <div className="flex items-center justify-between px-5 py-3.5 border-b border-slate-800/60 bg-[#121622]">
                <div className="flex items-center gap-2">
                  <FileSpreadsheet className="w-5 h-5 text-blue-400" />
                  <h3 className="text-xs font-bold uppercase tracking-wider text-slate-200 font-display">
                    Excel Sync & Data Loader
                  </h3>
                </div>
                <button
                  onClick={() => setIsUploadOpen(false)}
                  className="p-1.5 px-2 rounded-lg bg-slate-900 text-slate-400 hover:text-white border border-slate-800 transition-colors cursor-pointer text-xs font-bold"
                >
                  ✕
                </button>
              </div>

              {/* UploadSync component container */}
              <div className="p-4 bg-[#0A0D14]">
                <UploadSync
                  onDataLoaded={(newData) => {
                    handleDataLoaded(newData);
                    setIsUploadOpen(false);
                  }}
                  onRestoreDefault={() => {
                    handleRestoreDefault();
                    setIsUploadOpen(false);
                  }}
                  onClearAllData={() => {
                    handleClearAllData();
                  }}
                  dataLength={data.length}
                  isLightMode={isLightMode}
                />
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Auth Modal & Role Switcher */}
      <AuthModal
        isOpen={isAuthModalOpen}
        onClose={() => setIsAuthModalOpen(false)}
        currentUser={currentUser}
        onSelectUser={handleSelectUser}
        isLightMode={isLightMode}
      />

      {/* Admin Permission & User Management Console */}
      <AdminPortal
        isOpen={isAdminPortalOpen}
        onClose={() => setIsAdminPortalOpen(false)}
        currentUser={currentUser}
        isLightMode={isLightMode}
        activeUserCount={activeUserCount}
        activeUsersList={activeUsersList}
      />

      {/* Custom Confirmation Modal */}
      <AnimatePresence>
        {confirmModal.isOpen && (
          <div className="fixed inset-0 bg-black/70 backdrop-blur-sm flex items-center justify-center z-[999] p-4" id="custom-confirm-modal">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-[#121316] border border-white/10 rounded-2xl max-w-md w-full p-6 shadow-2xl relative"
            >
              <h3 className="text-base font-bold text-white mb-2">{confirmModal.title}</h3>
              <p className="text-xs text-slate-400 mb-6 leading-relaxed">{confirmModal.message}</p>
              <div className="flex gap-3 justify-end">
                <button
                  type="button"
                  onClick={() => setConfirmModal(prev => ({ ...prev, isOpen: false }))}
                  className="px-4 py-2 rounded-lg text-xs font-semibold bg-[#1A1D21] border border-white/10 text-slate-300 hover:bg-[#252830] transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={() => {
                    confirmModal.onConfirm();
                    setConfirmModal(prev => ({ ...prev, isOpen: false }));
                  }}
                  className="px-4 py-2 rounded-lg text-xs font-semibold bg-rose-600 hover:bg-rose-500 text-white transition-colors cursor-pointer shadow-lg shadow-rose-900/20"
                >
                  Confirm Action
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Sync Notification Toast */}
      <AnimatePresence>
        {syncNotification && (
          <motion.div
            initial={{ opacity: 0, y: 20, scale: 0.95 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: 20, scale: 0.95 }}
            className="fixed bottom-6 right-6 z-[9999] flex items-center gap-3 px-5 py-3.5 bg-emerald-950/90 text-emerald-100 border border-emerald-500/40 rounded-2xl shadow-2xl backdrop-blur-md"
            id="sync-notification-toast"
          >
            <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0" />
            <span className="text-xs font-bold font-sans tracking-tight">{syncNotification}</span>
            <button
              onClick={() => setSyncNotification(null)}
              className="ml-3 text-emerald-400 hover:text-white text-xs font-bold font-mono p-1 rounded-md hover:bg-emerald-800/40 cursor-pointer"
            >
              ✕
            </button>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
