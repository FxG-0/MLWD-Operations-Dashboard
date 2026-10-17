import React, { useState, useMemo, useRef, useEffect } from "react";
import { motion, AnimatePresence } from "motion/react";
import {
  Calendar,
  CalendarDays,
  CalendarRange,
  ChevronDown,
  ChevronRight,
  ChevronLeft,
  Search,
  SlidersHorizontal,
  Layers,
  Building2,
  Clock,
  Filter,
  Check,
  X,
  RotateCcw,
  Sparkles
} from "lucide-react";
import { RowData, Abbreviation, getSectionFromHoleSize, parseAnyDate } from "../types";

interface RigWellGanttChartProps {
  data: RowData[];
  abbreviations?: Abbreviation[];
  isLightMode?: boolean;
  onSelectTool?: (toolName: string) => void;
}

interface WellRunEvent {
  id: string;
  rig: string;
  well: string;
  section: string;
  holeSize: string;
  toolConfig: string;
  qty: number;
  footage: number | null;
  startDate: Date;
  endDate: Date;
  startDateStr: string;
  endDateStr: string;
  client: string;
  shopDate: string;
  comments: string;
  rawRow: RowData;
}

interface WellGroup {
  wellName: string;
  rigName: string;
  client: string;
  earliestDate: Date;
  latestDate: Date;
  totalFootage: number;
  totalRuns: number;
  runs: WellRunEvent[];
}

interface RigGroup {
  rigName: string;
  client: string;
  earliestDate: Date;
  latestDate: Date;
  totalWells: number;
  totalRuns: number;
  totalFootage: number;
  wells: WellGroup[];
  allRuns: WellRunEvent[];
}

const MONTHS_MAP: { [key: string]: number } = {
  jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
  jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
};

function parseDateRobust(dateStr: string): Date | null {
  if (!dateStr || dateStr.toLowerCase() === "n/a") return null;
  const trimmed = dateStr.trim();

  // Excel serial number
  if (/^\d+(\.\d+)?$/.test(trimmed)) {
    const num = Number(trimmed);
    if (num > 20000 && num < 60000) {
      const utcDays = num - 25569;
      const d = new Date(utcDays * 86400 * 1000);
      if (!isNaN(d.getTime())) return d;
    }
  }

  // ISO YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    const [y, m, d] = trimmed.split("-").map(Number);
    const date = new Date(y, m - 1, d);
    if (!isNaN(date.getTime())) return date;
  }

  // DD-MMM-YY e.g. 22-Jun-26 or 22-Jun-2026
  const parts = trimmed.split("-");
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const monthKey = parts[1].toLowerCase().slice(0, 3);
    let year = parseInt(parts[2], 10);
    if (year < 100) year += 2000;
    if (MONTHS_MAP[monthKey] !== undefined && !isNaN(day) && !isNaN(year)) {
      const date = new Date(year, MONTHS_MAP[monthKey], day);
      if (!isNaN(date.getTime())) return date;
    }
  }

  // Use parseAnyDate from central type utilities (handles DD/MM, MM/DD, formats with slashes, etc.)
  const anyParsed = parseAnyDate(trimmed);
  if (anyParsed && anyParsed.date) {
    const p = anyParsed.date.split("-");
    if (p.length === 3) {
      const d = parseInt(p[0], 10);
      const mKey = p[1].toLowerCase().slice(0, 3);
      let y = parseInt(p[2], 10);
      if (y < 100) y += 2000;
      if (MONTHS_MAP[mKey] !== undefined && !isNaN(d) && !isNaN(y)) {
        const date = new Date(y, MONTHS_MAP[mKey], d);
        if (!isNaN(date.getTime())) return date;
      }
    }
  }

  // Standard JS Date parsing fallback
  const parsed = Date.parse(trimmed);
  if (!isNaN(parsed)) {
    const d = new Date(parsed);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  return null;
}

function formatDateDisplay(d: Date): string {
  const day = d.getDate();
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const month = monthNames[d.getMonth()];
  const yr = String(d.getFullYear()).slice(-2);
  return `${day}-${month}-${yr}`;
}

function formatMonthYearDisplay(d: Date): string {
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  return `${monthNames[d.getMonth()]} ${d.getFullYear()}`;
}

function toISODateInput(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

// Tool Sizes definition with hole size conversion mapping
const TOOL_SIZE_DEFINITIONS = [
  { id: "4.75in", label: '4.75"', holeSize: '6" / 6.125" / 5 7/8"', desc: 'Slim Hole' },
  { id: "6.75in", label: '6.75"', holeSize: '8.5" / 8 1/2"', desc: 'Intermediate' },
  { id: "8in", label: '8"', holeSize: '12.25" / 12 1/4"', desc: 'Production' },
  { id: "9.5in", label: '9.5"', holeSize: '16" / 17.5" / 20"', desc: 'Surface' }
];

export const RigWellGanttChart: React.FC<RigWellGanttChartProps> = ({
  data,
  isLightMode = false
}) => {
  // Search & Filter State
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedRigFilter, setSelectedRigFilter] = useState<string>("all");
  
  // Multi-Select Filters
  const [selectedSizes, setSelectedSizes] = useState<string[]>([]); // empty = all sizes
  const [selectedMonths, setSelectedMonths] = useState<string[]>([]); // empty = all months
  
  // Custom Date Range State
  const [customStartDate, setCustomStartDate] = useState<string>("");
  const [customEndDate, setCustomEndDate] = useState<string>("");
  
  // Popover Toggles
  const [isDatePickerOpen, setIsDatePickerOpen] = useState(false);
  const [isSizePickerOpen, setIsSizePickerOpen] = useState(false);
  const [isMonthPickerOpen, setIsMonthPickerOpen] = useState(false);

  // Popover DOM Refs for Click-Outside
  const datePickerRef = useRef<HTMLDivElement>(null);
  const sizePickerRef = useRef<HTMLDivElement>(null);
  const monthPickerRef = useRef<HTMLDivElement>(null);

  const [expandedRigs, setExpandedRigs] = useState<{ [rigName: string]: boolean }>({});
  const [selectedRunDetail, setSelectedRunDetail] = useState<WellRunEvent | null>(null);
  const [viewMode, setViewMode] = useState<"gantt" | "table">("gantt");

  const timelineContainerRef = useRef<HTMLDivElement>(null);

  // Close popovers on click outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (datePickerRef.current && !datePickerRef.current.contains(e.target as Node)) {
        setIsDatePickerOpen(false);
      }
      if (sizePickerRef.current && !sizePickerRef.current.contains(e.target as Node)) {
        setIsSizePickerOpen(false);
      }
      if (monthPickerRef.current && !monthPickerRef.current.contains(e.target as Node)) {
        setIsMonthPickerOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  // 1. Process data into structured Rig -> Well -> Runs hierarchy
  const { rigGroups, allRuns, minDate, maxDate, totalHorizonDays, uniqueRigs, sizeCounts } = useMemo(() => {
    const runsList: WellRunEvent[] = [];
    const secSet = new Set<string>();
    const counts: { [size: string]: number } = {};

    data.forEach((row) => {
      const rig = (row.rig || "").trim();
      const well = (row.well || "").trim();
      if (!rig || rig.toUpperCase() === "UNKNOWN" || rig.toUpperCase() === "N/A") return;
      if (!well || well.toUpperCase() === "UNKNOWN" || well.toUpperCase() === "N/A") return;

      const parsedDate = parseDateRobust(row.date);
      if (!parsedDate) return;

      const normalizedSection = getSectionFromHoleSize(row.section || row.holeSize || "8in");
      secSet.add(normalizedSection);
      counts[normalizedSection] = (counts[normalizedSection] || 0) + 1;

      // Default duration for visualization: 5 days per section run or calculated duration
      const startDate = new Date(parsedDate.getTime());
      const endDate = new Date(parsedDate.getTime() + 5 * 24 * 60 * 60 * 1000);

      runsList.push({
        id: row.id || `${rig}-${well}-${normalizedSection}-${parsedDate.getTime()}`,
        rig,
        well,
        section: normalizedSection,
        holeSize: row.holeSize || "",
        toolConfig: row.configuration || "Standard Run",
        qty: row.qty || 1,
        footage: row.footage || null,
        startDate,
        endDate,
        startDateStr: formatDateDisplay(startDate),
        endDateStr: formatDateDisplay(endDate),
        client: row.client || "Client",
        shopDate: row.shopDate || "",
        comments: row.comments || "",
        rawRow: row
      });
    });

    if (runsList.length === 0) {
      const now = new Date();
      const defaultStart = new Date(now.getFullYear(), now.getMonth(), 1);
      const defaultEnd = new Date(now.getFullYear(), now.getMonth() + 3, 0);
      return {
        rigGroups: [],
        allRuns: [],
        minDate: defaultStart,
        maxDate: defaultEnd,
        totalHorizonDays: 90,
        uniqueRigs: [],
        uniqueSections: [],
        sizeCounts: {}
      };
    }

    // Determine timeline boundaries with padding
    const minTime = Math.min(...runsList.map((r) => r.startDate.getTime()));
    const maxTime = Math.max(...runsList.map((r) => r.endDate.getTime()));

    // Pad by 4 days before & 7 days after for clean edge spacing
    const minPadded = new Date(minTime - 4 * 24 * 60 * 60 * 1000);
    const maxPadded = new Date(maxTime + 7 * 24 * 60 * 60 * 1000);

    const totalDays = Math.max(
      1,
      Math.ceil((maxPadded.getTime() - minPadded.getTime()) / (24 * 60 * 60 * 1000))
    );

    // Group by Rig -> Well
    const rigMap = new Map<string, Map<string, WellRunEvent[]>>();
    runsList.forEach((run) => {
      if (!rigMap.has(run.rig)) {
        rigMap.set(run.rig, new Map());
      }
      const wellMap = rigMap.get(run.rig)!;
      if (!wellMap.has(run.well)) {
        wellMap.set(run.well, []);
      }
      wellMap.get(run.well)!.push(run);
    });

    const rigsArray: RigGroup[] = [];
    const rigNamesList: string[] = [];

    rigMap.forEach((wellsMap, rigName) => {
      rigNamesList.push(rigName);
      const wellGroups: WellGroup[] = [];
      const rigRunsAll: WellRunEvent[] = [];

      wellsMap.forEach((runs, wellName) => {
        runs.sort((a, b) => a.startDate.getTime() - b.startDate.getTime());
        rigRunsAll.push(...runs);

        const wMin = new Date(Math.min(...runs.map((r) => r.startDate.getTime())));
        const wMax = new Date(Math.max(...runs.map((r) => r.endDate.getTime())));
        const wFootage = runs.reduce((acc, curr) => acc + (curr.footage || 0), 0);

        wellGroups.push({
          wellName,
          rigName,
          client: runs[0]?.client || "Client",
          earliestDate: wMin,
          latestDate: wMax,
          totalFootage: wFootage,
          totalRuns: runs.length,
          runs
        });
      });

      wellGroups.sort((a, b) => a.earliestDate.getTime() - b.earliestDate.getTime());

      const rMin = new Date(Math.min(...rigRunsAll.map((r) => r.startDate.getTime())));
      const rMax = new Date(Math.max(...rigRunsAll.map((r) => r.endDate.getTime())));
      const rFootage = rigRunsAll.reduce((acc, curr) => acc + (curr.footage || 0), 0);

      rigsArray.push({
        rigName,
        client: rigRunsAll[0]?.client || "Client",
        earliestDate: rMin,
        latestDate: rMax,
        totalWells: wellGroups.length,
        totalRuns: rigRunsAll.length,
        totalFootage: rFootage,
        wells: wellGroups,
        allRuns: rigRunsAll
      });
    });

    rigsArray.sort((a, b) => a.rigName.localeCompare(b.rigName, undefined, { numeric: true }));

    return {
      rigGroups: rigsArray,
      allRuns: runsList,
      minDate: minPadded,
      maxDate: maxPadded,
      totalHorizonDays: totalDays,
      uniqueRigs: rigNamesList.sort((a, b) => a.localeCompare(b, undefined, { numeric: true })),
      uniqueSections: Array.from(secSet).sort(),
      sizeCounts: counts
    };
  }, [data]);

  // Keep selectedRunDetail modal updated if data is edited in the Excel editor
  useEffect(() => {
    if (selectedRunDetail) {
      const matchingRun = allRuns.find(r => r.id === selectedRunDetail.id);
      if (matchingRun) {
        setSelectedRunDetail(matchingRun);
      }
    }
  }, [allRuns]);

  // Extract all distinct available months from the dataset
  const availableMonths = useMemo(() => {
    const monthsMap = new Map<string, { key: string; label: string; shortLabel: string; year: number; month: number; count: number }>();
    const monthNames = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    const shortMonthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    allRuns.forEach((run) => {
      const y = run.startDate.getFullYear();
      const m = run.startDate.getMonth();
      const key = `${y}-${String(m + 1).padStart(2, "0")}`;
      if (!monthsMap.has(key)) {
        monthsMap.set(key, {
          key,
          label: `${monthNames[m]} ${y}`,
          shortLabel: `${shortMonthNames[m]} '${String(y).slice(-2)}`,
          year: y,
          month: m,
          count: 0
        });
      }
      monthsMap.get(key)!.count += 1;
    });

    return Array.from(monthsMap.values()).sort((a, b) => a.key.localeCompare(b.key));
  }, [allRuns]);

  // Filtered Rigs based on user controls (Search, Rig, Multiple Sizes, Multiple Months, Custom Date Range)
  const filteredRigGroups = useMemo(() => {
    const q = searchTerm.trim().toLowerCase();
    const hasCustomRange = !!(customStartDate && customEndDate);
    const rangeStart = hasCustomRange ? parseDateRobust(customStartDate)?.getTime() : null;
    const rangeEnd = hasCustomRange ? parseDateRobust(customEndDate)?.getTime() : null;

    return rigGroups
      .map((rig) => {
        // Filter by Rig selector
        if (selectedRigFilter !== "all" && rig.rigName !== selectedRigFilter) {
          return null;
        }

        // Filter wells
        const matchingWells = rig.wells
          .map((well) => {
            let matchingRuns = well.runs;

            // 1. Filter runs by Multiple Tool Sizes
            if (selectedSizes.length > 0) {
              matchingRuns = matchingRuns.filter((r) => selectedSizes.includes(r.section));
            }

            // 2. Filter runs by Custom Date Range OR Multiple Selected Months
            if (hasCustomRange && rangeStart && rangeEnd) {
              const rEndExtended = rangeEnd + 24 * 60 * 60 * 1000 - 1;
              matchingRuns = matchingRuns.filter(
                (r) => r.startDate.getTime() <= rEndExtended && r.endDate.getTime() >= rangeStart
              );
            } else if (selectedMonths.length > 0) {
              matchingRuns = matchingRuns.filter((r) => {
                const runKey = `${r.startDate.getFullYear()}-${String(r.startDate.getMonth() + 1).padStart(2, "0")}`;
                return selectedMonths.includes(runKey);
              });
            }

            // 3. Filter runs by search query (rig, well, config, client, date, comments)
            if (q) {
              matchingRuns = matchingRuns.filter((r) => {
                return (
                  r.rig.toLowerCase().includes(q) ||
                  r.well.toLowerCase().includes(q) ||
                  r.toolConfig.toLowerCase().includes(q) ||
                  r.client.toLowerCase().includes(q) ||
                  r.section.toLowerCase().includes(q) ||
                  r.startDateStr.toLowerCase().includes(q) ||
                  r.comments.toLowerCase().includes(q)
                );
              });
            }

            if (matchingRuns.length === 0) return null;

            return {
              ...well,
              runs: matchingRuns,
              totalRuns: matchingRuns.length
            };
          })
          .filter(Boolean) as WellGroup[];

        if (matchingWells.length === 0) return null;

        const allFilteredRuns = matchingWells.flatMap((w) => w.runs);

        return {
          ...rig,
          wells: matchingWells,
          totalWells: matchingWells.length,
          allRuns: allFilteredRuns,
          totalRuns: allFilteredRuns.length
        };
      })
      .filter(Boolean) as RigGroup[];
  }, [
    rigGroups,
    searchTerm,
    selectedRigFilter,
    selectedSizes,
    selectedMonths,
    customStartDate,
    customEndDate
  ]);

  // Dynamic Timeline Day Cells, Month Segments, and Day Map
  const {
    timelineDays,
    timelineMonths,
    dateKeyToDayMap,
    isFocusedView,
    activeMinDate,
    activeMaxDate,
    activeHorizonDays
  } = useMemo(() => {
    interface DayCell {
      dateKey: string; // YYYY-MM-DD
      date: Date;
      dayNum: number;
      dayOfWeek: string;
      dateStr: string;
      monthKey: string;
      startPercent: number;
      widthPercent: number;
      isWeekend: boolean;
      hasRun: boolean;
      runsStartingCount: number;
      runsSummaryText: string;
    }

    interface MonthSegment {
      monthKey: string;
      label: string;
      startPercent: number;
      widthPercent: number;
      daysCount: number;
    }

    const weekdays = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];
    const monthNames = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];

    const allFilteredRuns = filteredRigGroups.flatMap((r) => r.allRuns);

    // Build raw list of dates to include on timeline
    const rawDates: Date[] = [];
    const monthSegmentKeys: string[] = [];

    // Case 1: Custom Date Range (continuous from start to end)
    if (customStartDate && customEndDate) {
      const s = parseDateRobust(customStartDate);
      const e = parseDateRobust(customEndDate);
      if (s && e && e.getTime() >= s.getTime()) {
        let cur = new Date(s.getFullYear(), s.getMonth(), s.getDate(), 0, 0, 0);
        const endDay = new Date(e.getFullYear(), e.getMonth(), e.getDate(), 23, 59, 59);
        while (cur.getTime() <= endDay.getTime()) {
          rawDates.push(new Date(cur.getTime()));
          const mKey = `${cur.getFullYear()}-${String(cur.getMonth() + 1).padStart(2, "0")}`;
          if (!monthSegmentKeys.includes(mKey)) {
            monthSegmentKeys.push(mKey);
          }
          cur.setDate(cur.getDate() + 1);
        }
      }
    }
    // Case 2: Specific Selected Months (ONLY include selected months, hiding any unselected months in between!)
    else if (selectedMonths.length > 0) {
      const sortedMonths = [...selectedMonths].sort();
      sortedMonths.forEach((mKey) => {
        monthSegmentKeys.push(mKey);
        const [y, m] = mKey.split("-").map(Number);
        const daysInMonth = new Date(y, m, 0).getDate();
        for (let d = 1; d <= daysInMonth; d++) {
          rawDates.push(new Date(y, m - 1, d, 0, 0, 0));
        }
      });
    }
    // Case 3: All Timeline (dataset months in order)
    else if (availableMonths.length > 0) {
      availableMonths.forEach((mObj) => {
        monthSegmentKeys.push(mObj.key);
        const [y, m] = mObj.key.split("-").map(Number);
        const daysInMonth = new Date(y, m, 0).getDate();
        for (let d = 1; d <= daysInMonth; d++) {
          rawDates.push(new Date(y, m - 1, d, 0, 0, 0));
        }
      });
    } else {
      // Fallback
      const now = new Date();
      for (let d = 1; d <= 30; d++) {
        rawDates.push(new Date(now.getFullYear(), now.getMonth(), d, 0, 0, 0));
      }
    }

    const totalDays = Math.max(1, rawDates.length);
    const dayCells: DayCell[] = [];
    const dateMap = new Map<string, DayCell>();

    rawDates.forEach((d, idx) => {
      const curYear = d.getFullYear();
      const curMonth = d.getMonth();
      const curDate = d.getDate();
      const dateKey = `${curYear}-${String(curMonth + 1).padStart(2, "0")}-${String(curDate).padStart(2, "0")}`;
      const monthKey = `${curYear}-${String(curMonth + 1).padStart(2, "0")}`;
      const dayOfWeekIdx = d.getDay();

      const startPercent = (idx / totalDays) * 100;
      const widthPercent = (1 / totalDays) * 100;

      // Find matching runs starting on this date
      const matchingRuns = allFilteredRuns.filter((r) => {
        const rStart = r.startDate;
        return (
          rStart.getFullYear() === curYear &&
          rStart.getMonth() === curMonth &&
          rStart.getDate() === curDate
        );
      });

      const summaryText =
        matchingRuns.length > 0
          ? `${formatDateDisplay(d)}: ${matchingRuns.length} run${matchingRuns.length > 1 ? "s" : ""} (${matchingRuns.map((r) => `${r.rig} - ${r.well}`).join(", ")})`
          : formatDateDisplay(d);

      const cell: DayCell = {
        dateKey,
        date: d,
        dayNum: curDate,
        dayOfWeek: weekdays[dayOfWeekIdx],
        dateStr: formatDateDisplay(d),
        monthKey,
        startPercent,
        widthPercent,
        isWeekend: dayOfWeekIdx === 0 || dayOfWeekIdx === 6,
        hasRun: matchingRuns.length > 0,
        runsStartingCount: matchingRuns.length,
        runsSummaryText: summaryText
      };

      dayCells.push(cell);
      dateMap.set(dateKey, cell);
    });

    // Build Month Segments
    const monthSegments: MonthSegment[] = [];
    monthSegmentKeys.forEach((mKey) => {
      const [y, m] = mKey.split("-").map(Number);
      const daysInThisSegment = dayCells.filter((c) => c.monthKey === mKey);
      if (daysInThisSegment.length > 0) {
        const firstIdx = dayCells.indexOf(daysInThisSegment[0]);
        const startPercent = (firstIdx / totalDays) * 100;
        const widthPercent = (daysInThisSegment.length / totalDays) * 100;

        monthSegments.push({
          monthKey: mKey,
          label: `${monthNames[m - 1]} ${y}`,
          startPercent,
          widthPercent,
          daysCount: daysInThisSegment.length
        });
      }
    });

    const isFocused =
      selectedMonths.length > 0 ||
      !!(customStartDate && customEndDate) ||
      totalDays <= 90;

    const activeMin = rawDates[0] || minDate;
    const activeMax = rawDates[rawDates.length - 1] || maxDate;

    return {
      timelineDays: dayCells,
      timelineMonths: monthSegments,
      dateKeyToDayMap: dateMap,
      isFocusedView: isFocused,
      activeMinDate: activeMin,
      activeMaxDate: activeMax,
      activeHorizonDays: totalDays
    };
  }, [
    customStartDate,
    customEndDate,
    selectedMonths,
    availableMonths,
    filteredRigGroups,
    minDate,
    maxDate
  ]);

  // Multi-size toggle helper
  const handleToggleSize = (sizeId: string) => {
    setSelectedSizes((prev) =>
      prev.includes(sizeId) ? prev.filter((s) => s !== sizeId) : [...prev, sizeId]
    );
  };

  // Multi-month toggle helper
  const handleToggleMonth = (monthKey: string) => {
    // If a custom date range was active, clear it when user clicks month pills
    if (customStartDate || customEndDate) {
      setCustomStartDate("");
      setCustomEndDate("");
    }
    setSelectedMonths((prev) =>
      prev.includes(monthKey) ? prev.filter((m) => m !== monthKey) : [...prev, monthKey]
    );
  };

  // Month navigation stepper (< and >)
  const handleStepMonth = (direction: -1 | 1) => {
    if (availableMonths.length === 0) return;
    if (selectedMonths.length === 0) {
      if (direction === 1) {
        setSelectedMonths([availableMonths[0].key]);
      } else {
        setSelectedMonths([availableMonths[availableMonths.length - 1].key]);
      }
      return;
    }
    const currentKey = selectedMonths[0];
    const idx = availableMonths.findIndex((m) => m.key === currentKey);
    if (idx === -1) {
      setSelectedMonths([availableMonths[0].key]);
      return;
    }
    const nextIdx = idx + direction;
    if (nextIdx >= 0 && nextIdx < availableMonths.length) {
      setSelectedMonths([availableMonths[nextIdx].key]);
    } else {
      setSelectedMonths([]);
    }
  };

  // Generate Timeline Header Columns:
  const timelineHeader = useMemo(() => {
    return {
      months: timelineMonths,
      days: timelineDays,
      isFocusedView
    };
  }, [timelineMonths, timelineDays, isFocusedView]);

  // Color lookup per section
  const getSectionBadgeColor = (sec: string) => {
    const s = sec.toLowerCase();
    if (s.includes("4.75") || s.includes("6in") || s.includes("5 7/8") || s.includes("6.125")) {
      return {
        bg: isLightMode ? "bg-blue-50 text-blue-800 border-blue-200" : "bg-blue-950/80 text-blue-300 border-blue-800/80",
        bar: isLightMode ? "#2563EB" : "#3B82F6",
        barClass: "bg-blue-600 hover:bg-blue-500",
        dot: "bg-blue-500"
      };
    }
    if (s.includes("6.75") || s.includes("8.5") || s.includes("8 1/2")) {
      return {
        bg: isLightMode ? "bg-emerald-50 text-emerald-800 border-emerald-200" : "bg-emerald-950/80 text-emerald-300 border-emerald-800/80",
        bar: isLightMode ? "#059669" : "#10B981",
        barClass: "bg-emerald-600 hover:bg-emerald-500",
        dot: "bg-emerald-500"
      };
    }
    if (s.includes("8in") || s.includes("12.25") || s.includes("12 1/4")) {
      return {
        bg: isLightMode ? "bg-purple-50 text-purple-800 border-purple-200" : "bg-purple-950/80 text-purple-300 border-purple-800/80",
        bar: isLightMode ? "#9333EA" : "#A855F7",
        barClass: "bg-purple-600 hover:bg-purple-500",
        dot: "bg-purple-500"
      };
    }
    if (s.includes("9.5") || s.includes("16") || s.includes("17") || s.includes("20")) {
      return {
        bg: isLightMode ? "bg-amber-50 text-amber-800 border-amber-200" : "bg-amber-950/80 text-amber-300 border-amber-800/80",
        bar: isLightMode ? "#D97706" : "#F59E0B",
        barClass: "bg-amber-600 hover:bg-amber-500",
        dot: "bg-amber-500"
      };
    }
    return {
      bg: isLightMode ? "bg-slate-100 text-slate-700 border-slate-200" : "bg-slate-800 text-slate-300 border-slate-700",
      bar: isLightMode ? "#0891B2" : "#06B6D4",
      barClass: "bg-cyan-600 hover:bg-cyan-500",
      dot: "bg-cyan-500"
    };
  };

  // Calculate coordinates for a run on the timeline
  const getTimelineBarStyles = (start: Date, end: Date) => {
    const isManualDateSelection = Boolean(customStartDate && customEndDate);
    const startKey = `${start.getFullYear()}-${String(start.getMonth() + 1).padStart(2, "0")}-${String(start.getDate()).padStart(2, "0")}`;
    const startCell = dateKeyToDayMap.get(startKey);

    // 1. If manual specific date selection is active, fit the bar inside the single day column cell
    if (isManualDateSelection) {
      if (!startCell) {
        return {
          left: "0%",
          width: "0%",
          isVisible: false
        };
      }
      return {
        left: `${startCell.startPercent}%`,
        width: `${dayCellWidthOrPercent(startCell)}%`,
        isVisible: true
      };
    }

    // 2. Otherwise (month / multiple months / general overview): return to standard duration bar spanning
    if (startCell) {
      const spanDays = Math.max(
        1,
        Math.round((end.getTime() - start.getTime()) / (24 * 60 * 60 * 1000))
      );
      const rawWidth = startCell.widthPercent * spanDays;
      const finalWidth = Math.min(
        100 - startCell.startPercent,
        Math.max(startCell.widthPercent * 1.2, rawWidth)
      );

      return {
        left: `${startCell.startPercent}%`,
        width: `${finalWidth}%`,
        isVisible: true
      };
    }

    // If start is before the first visible day cell, check if the run spans into visible days
    if (timelineDays.length > 0) {
      const firstCell = timelineDays[0];
      const lastCell = timelineDays[timelineDays.length - 1];
      if (end.getTime() >= firstCell.date.getTime() && start.getTime() <= lastCell.date.getTime()) {
        const firstVisibleCell = timelineDays.find(
          (c) => c.date.getTime() >= start.getTime() && c.date.getTime() <= end.getTime()
        );
        if (firstVisibleCell) {
          const remainingDays = Math.max(
            1,
            Math.round((end.getTime() - firstVisibleCell.date.getTime()) / (24 * 60 * 60 * 1000))
          );
          const rawWidth = firstVisibleCell.widthPercent * remainingDays;
          const finalWidth = Math.min(
            100 - firstVisibleCell.startPercent,
            Math.max(firstVisibleCell.widthPercent * 1.2, rawWidth)
          );
          return {
            left: `${firstVisibleCell.startPercent}%`,
            width: `${finalWidth}%`,
            isVisible: true
          };
        }
      }
    }

    return {
      left: "0%",
      width: "0%",
      isVisible: false
    };
  };

  const dayCellWidthOrPercent = (cell: { widthPercent: number }) => cell.widthPercent;

  // Quick stats
  const totalStats = useMemo(() => {
    const totalRigs = filteredRigGroups.length;
    const totalWells = filteredRigGroups.reduce((acc, r) => acc + r.totalWells, 0);
    const totalRuns = filteredRigGroups.reduce((acc, r) => acc + r.totalRuns, 0);
    const totalFootage = filteredRigGroups.reduce((acc, r) => acc + r.totalFootage, 0);

    const horizonStr =
      selectedMonths.length > 0
        ? selectedMonths
            .slice()
            .sort()
            .map((m) => {
              const [y, mon] = m.split("-").map(Number);
              const mNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
              return `${mNames[mon - 1]} '${String(y).slice(-2)}`;
            })
            .join(", ")
        : `${formatDateDisplay(activeMinDate)} → ${formatDateDisplay(activeMaxDate)}`;

    return {
      totalRigs,
      totalWells,
      totalRuns,
      totalFootage,
      horizon: horizonStr
    };
  }, [filteredRigGroups, selectedMonths, activeMinDate, activeMaxDate]);

  const toggleAllRigsExpansion = (expand: boolean) => {
    const newState: { [rigName: string]: boolean } = {};
    rigGroups.forEach((r) => {
      newState[r.rigName] = expand;
    });
    setExpandedRigs(newState);
  };

  // Quick Preset Handlers for Calendar Range
  const applyDatePreset = (daysAhead: number) => {
    const start = minDate.getTime() > 0 ? new Date(minDate.getTime()) : new Date();
    const end = new Date(start.getTime() + daysAhead * 24 * 60 * 60 * 1000);
    setCustomStartDate(toISODateInput(start));
    setCustomEndDate(toISODateInput(end));
    setSelectedMonths([]);
  };

  const applyYearPreset = () => {
    const y = minDate.getFullYear() || new Date().getFullYear();
    setCustomStartDate(`${y}-01-01`);
    setCustomEndDate(`${y}-12-31`);
    setSelectedMonths([]);
  };

  const clearDateRange = () => {
    setCustomStartDate("");
    setCustomEndDate("");
  };

  const isCustomRangeActive = !!(customStartDate && customEndDate);

  return (
    <div className="space-y-6">
      {/* Top Banner Card */}
      <div
        className={`p-5 sm:p-6 rounded-2xl border transition-all ${
          isLightMode
            ? "bg-white border-slate-200 shadow-sm"
            : "bg-[#121620] border-slate-700/60 shadow-lg"
        }`}
      >
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="flex items-start gap-4">
            <div
              className={`p-3 rounded-xl border shrink-0 ${
                isLightMode
                  ? "bg-emerald-50 text-emerald-700 border-emerald-200"
                  : "bg-emerald-500/10 text-emerald-400 border-emerald-500/30"
              }`}
            >
              <CalendarRange className="w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2
                  className={`text-lg sm:text-xl font-bold font-sans tracking-tight ${
                    isLightMode ? "text-slate-900" : "text-white"
                  }`}
                >
                  Rig & Well Operational Schedule (Gantt Chart)
                </h2>
                <span
                  className={`px-2 py-0.5 rounded-full text-[10px] font-mono font-bold uppercase tracking-wider ${
                    isLightMode
                      ? "bg-emerald-100 text-emerald-800"
                      : "bg-emerald-500/20 text-emerald-300 border border-emerald-500/30"
                  }`}
                >
                  Live Timeline
                </span>
              </div>
              <p
                className={`text-xs sm:text-sm mt-1 font-sans ${
                  isLightMode ? "text-slate-600" : "text-[#9AA0A6]"
                }`}
              >
                Interactive timeline mapping drilling rigs, well mobilization schedules, and run dates. Select any month or date range to inspect exact daily milestones.
              </p>
            </div>
          </div>

          {/* Quick Metrics Bar */}
          <div className="flex flex-wrap items-center gap-2 sm:gap-3">
            <div
              className={`px-3 py-2 rounded-xl border flex items-center gap-2.5 ${
                isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#1A1F2C] border-slate-700/80"
              }`}
            >
              <Building2 className="w-4 h-4 text-emerald-500" />
              <div>
                <div className={`text-[10px] uppercase font-bold tracking-wider ${isLightMode ? "text-slate-500" : "text-slate-400"}`}>
                  Active Rigs
                </div>
                <div className={`text-sm font-bold font-mono ${isLightMode ? "text-slate-900" : "text-white"}`}>
                  {totalStats.totalRigs}
                </div>
              </div>
            </div>

            <div
              className={`px-3 py-2 rounded-xl border flex items-center gap-2.5 ${
                isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#1A1F2C] border-slate-700/80"
              }`}
            >
              <Layers className="w-4 h-4 text-blue-500" />
              <div>
                <div className={`text-[10px] uppercase font-bold tracking-wider ${isLightMode ? "text-slate-500" : "text-slate-400"}`}>
                  Scheduled Wells
                </div>
                <div className={`text-sm font-bold font-mono ${isLightMode ? "text-slate-900" : "text-white"}`}>
                  {totalStats.totalWells}
                </div>
              </div>
            </div>

            <div
              className={`px-3 py-2 rounded-xl border flex items-center gap-2.5 ${
                isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#1A1F2C] border-slate-700/80"
              }`}
            >
              <Clock className="w-4 h-4 text-purple-400" />
              <div>
                <div className={`text-[10px] uppercase font-bold tracking-wider ${isLightMode ? "text-slate-500" : "text-slate-400"}`}>
                  Total Runs
                </div>
                <div className={`text-sm font-bold font-mono ${isLightMode ? "text-slate-900" : "text-white"}`}>
                  {totalStats.totalRuns}
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Primary Filter Row: Search, Rig Dropdown, Multi-Size Selector, and Calendar Button */}
        <div className="mt-5 pt-5 border-t border-slate-200/60 dark:border-slate-800 flex flex-col xl:flex-row xl:items-center justify-between gap-4">
          <div className="flex flex-wrap items-center gap-2.5 flex-1">
            {/* Search Box */}
            <div className="relative min-w-[200px] max-w-xs flex-1">
              <Search
                className={`w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 ${
                  isLightMode ? "text-slate-400" : "text-slate-500"
                }`}
              />
              <input
                type="text"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                placeholder="Search Rig, Well, Tool, Client..."
                className={`w-full pl-9 pr-8 py-1.5 text-xs rounded-xl border font-sans outline-hidden transition-all ${
                  isLightMode
                    ? "bg-slate-50 border-slate-200 text-slate-800 placeholder:text-slate-400 focus:bg-white focus:border-blue-500"
                    : "bg-[#0B0D14] border-slate-700/80 text-slate-100 placeholder:text-slate-500 focus:border-emerald-500"
                }`}
              />
              {searchTerm && (
                <button
                  onClick={() => setSearchTerm("")}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-200 p-0.5"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              )}
            </div>

            {/* Rig Dropdown Filter */}
            <select
              value={selectedRigFilter}
              onChange={(e) => setSelectedRigFilter(e.target.value)}
              className={`px-3 py-1.5 text-xs rounded-xl border font-sans font-medium outline-hidden cursor-pointer ${
                isLightMode
                  ? "bg-slate-50 border-slate-200 text-slate-800"
                  : "bg-[#0B0D14] border-slate-700/80 text-slate-200"
              }`}
            >
              <option value="all">All Operating Rigs ({uniqueRigs.length})</option>
              {uniqueRigs.map((rig) => (
                <option key={rig} value={rig}>
                  Rig {rig}
                </option>
              ))}
            </select>

            {/* Multi-Size Selection Controls */}
            <div className="flex items-center gap-1.5 flex-wrap">
              {/* All Sizes Button */}
              <button
                onClick={() => setSelectedSizes([])}
                className={`px-2.5 py-1 text-xs font-semibold rounded-lg border transition-all cursor-pointer ${
                  selectedSizes.length === 0
                    ? isLightMode
                      ? "bg-slate-900 text-white border-slate-900 shadow-2xs"
                      : "bg-white text-slate-900 border-white shadow-2xs"
                    : isLightMode
                    ? "bg-slate-100 text-slate-600 border-slate-200 hover:bg-slate-200"
                    : "bg-[#1A1F2C] text-slate-400 border-slate-700/80 hover:text-white"
                }`}
              >
                All Sizes
              </button>

              {/* Individual Multi-Select Size Pills */}
              {TOOL_SIZE_DEFINITIONS.map((def) => {
                const colors = getSectionBadgeColor(def.id);
                const isSelected = selectedSizes.includes(def.id);
                const count = sizeCounts[def.id] || 0;

                return (
                  <button
                    key={def.id}
                    onClick={() => handleToggleSize(def.id)}
                    title={`${def.label} Tool Size → Hole Sizes: ${def.holeSize}`}
                    className={`px-2.5 py-1 text-xs font-semibold rounded-lg border flex items-center gap-1.5 transition-all cursor-pointer ${
                      isSelected
                        ? `${colors.bg} font-bold shadow-xs scale-105 ring-1 ring-emerald-500/40`
                        : isLightMode
                        ? "bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100"
                        : "bg-[#1A1F2C] text-slate-400 border-slate-700/80 hover:text-white"
                    }`}
                  >
                    <span className={`w-2 h-2 rounded-full ${colors.dot}`} />
                    <span>{def.label}</span>
                    <span className="text-[10px] opacity-75 font-mono">({count})</span>
                    {isSelected && <Check className="w-3 h-3 ml-0.5 text-emerald-500" />}
                  </button>
                );
              })}

              {/* Sizes Multi-Select Popover Dropdown */}
              <div className="relative" ref={sizePickerRef}>
                <button
                  onClick={() => setIsSizePickerOpen(!isSizePickerOpen)}
                  className={`p-1.5 rounded-lg border text-xs font-medium transition-all flex items-center gap-1 cursor-pointer ${
                    selectedSizes.length > 0
                      ? isLightMode
                        ? "bg-blue-50 border-blue-300 text-blue-700"
                        : "bg-emerald-950/80 border-emerald-700 text-emerald-300"
                      : isLightMode
                      ? "bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100"
                      : "bg-[#1A1F2C] border-slate-700/80 text-slate-300 hover:bg-[#252C3D]"
                  }`}
                  title="Granular Tool / Hole Size Multi-Selector"
                >
                  <Filter className="w-3.5 h-3.5" />
                  <span className="text-[11px]">
                    {selectedSizes.length > 0 ? `${selectedSizes.length} Selected` : "Sizes"}
                  </span>
                  <ChevronDown className="w-3 h-3 opacity-60" />
                </button>

                {isSizePickerOpen && (
                  <div
                    className={`absolute left-0 top-full mt-2 w-72 p-3 rounded-2xl border shadow-xl z-50 font-sans ${
                      isLightMode ? "bg-white border-slate-200" : "bg-[#161B26] border-slate-700"
                    }`}
                  >
                    <div className="flex items-center justify-between pb-2 border-b border-slate-200 dark:border-slate-800">
                      <span className="text-xs font-bold text-slate-800 dark:text-slate-200">
                        Select Tool / Hole Sizes
                      </span>
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => setSelectedSizes(TOOL_SIZE_DEFINITIONS.map((s) => s.id))}
                          className="text-[10px] font-semibold text-blue-600 dark:text-emerald-400 hover:underline"
                        >
                          Select All
                        </button>
                        <span className="text-slate-300 dark:text-slate-700">•</span>
                        <button
                          onClick={() => setSelectedSizes([])}
                          className="text-[10px] font-semibold text-slate-500 hover:underline"
                        >
                          Clear
                        </button>
                      </div>
                    </div>

                    <div className="divide-y divide-slate-100 dark:divide-slate-800/80 mt-2 max-h-56 overflow-y-auto custom-scrollbar">
                      {TOOL_SIZE_DEFINITIONS.map((def) => {
                        const isChecked = selectedSizes.includes(def.id);
                        const count = sizeCounts[def.id] || 0;
                        return (
                          <label
                            key={def.id}
                            className="flex items-start gap-2.5 py-2 px-1 hover:bg-slate-50 dark:hover:bg-slate-800/50 rounded-lg cursor-pointer transition-colors"
                          >
                            <input
                              type="checkbox"
                              checked={isChecked}
                              onChange={() => handleToggleSize(def.id)}
                              className="mt-0.5 rounded text-emerald-600 focus:ring-emerald-500"
                            />
                            <div className="flex-1 min-w-0">
                              <div className="flex items-center justify-between">
                                <span className="font-bold text-xs text-slate-800 dark:text-slate-200">
                                  {def.label} ({def.desc})
                                </span>
                                <span className="text-[10px] font-mono text-slate-500 font-semibold">
                                  {count} runs
                                </span>
                              </div>
                              <p className="text-[10px] text-slate-500 dark:text-slate-400 truncate">
                                Holes: {def.holeSize}
                              </p>
                            </div>
                          </label>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            </div>

            {/* Dedicated Calendar Button for Date Range Picker */}
            <div className="relative" ref={datePickerRef}>
              <button
                onClick={() => setIsDatePickerOpen(!isDatePickerOpen)}
                className={`px-3 py-1.5 rounded-xl border text-xs font-semibold flex items-center gap-2 transition-all cursor-pointer shadow-2xs ${
                  isCustomRangeActive
                    ? isLightMode
                      ? "bg-blue-600 text-white border-blue-600 shadow-sm"
                      : "bg-emerald-600 text-white border-emerald-600 shadow-sm"
                    : isLightMode
                    ? "bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700"
                    : "bg-[#1A1F2C] hover:bg-[#252C3D] border-slate-700/80 text-slate-200"
                }`}
                title="Select Custom Date Range Horizon"
              >
                <Calendar className="w-4 h-4" />
                <span>
                  {isCustomRangeActive
                    ? `${formatDateDisplay(parseDateRobust(customStartDate)!)} → ${formatDateDisplay(parseDateRobust(customEndDate)!)}`
                    : "Date Range"}
                </span>
                <ChevronDown className="w-3.5 h-3.5 opacity-70" />
              </button>

              {/* Date Range Picker Popover Card */}
              {isDatePickerOpen && (
                <div
                  className={`absolute left-0 top-full mt-2 w-80 p-4 rounded-2xl border shadow-2xl z-50 font-sans ${
                    isLightMode ? "bg-white border-slate-200" : "bg-[#161B26] border-slate-700"
                  }`}
                >
                  <div className="flex items-center justify-between pb-3 border-b border-slate-200 dark:border-slate-800">
                    <div className="flex items-center gap-1.5">
                      <CalendarDays className="w-4 h-4 text-emerald-500" />
                      <span className="text-xs font-bold text-slate-900 dark:text-white">
                        Custom Date Range
                      </span>
                    </div>
                    {isCustomRangeActive && (
                      <button
                        onClick={clearDateRange}
                        className="text-[10px] font-semibold text-rose-500 hover:underline flex items-center gap-1"
                      >
                        <RotateCcw className="w-3 h-3" />
                        Reset
                      </button>
                    )}
                  </div>

                  {/* Start & End Date Inputs */}
                  <div className="grid grid-cols-2 gap-2.5 my-3">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">
                        Start Date
                      </label>
                      <input
                        type="date"
                        value={customStartDate}
                        onChange={(e) => {
                          setCustomStartDate(e.target.value);
                          setSelectedMonths([]);
                        }}
                        className={`w-full px-2.5 py-1.5 text-xs rounded-xl border font-mono outline-hidden ${
                          isLightMode
                            ? "bg-slate-50 border-slate-200 text-slate-800 focus:border-blue-500"
                            : "bg-[#0B0D14] border-slate-700 text-slate-100 focus:border-emerald-500"
                        }`}
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">
                        End Date
                      </label>
                      <input
                        type="date"
                        value={customEndDate}
                        onChange={(e) => {
                          setCustomEndDate(e.target.value);
                          setSelectedMonths([]);
                        }}
                        className={`w-full px-2.5 py-1.5 text-xs rounded-xl border font-mono outline-hidden ${
                          isLightMode
                            ? "bg-slate-50 border-slate-200 text-slate-800 focus:border-blue-500"
                            : "bg-[#0B0D14] border-slate-700 text-slate-100 focus:border-emerald-500"
                        }`}
                      />
                    </div>
                  </div>

                  {/* Quick Horizon Presets */}
                  <div className="space-y-1.5 pt-2 border-t border-slate-100 dark:border-slate-800">
                    <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                      Quick Horizons
                    </span>
                    <div className="grid grid-cols-3 gap-1.5">
                      <button
                        onClick={() => applyDatePreset(30)}
                        className={`px-2 py-1 text-[10px] font-semibold rounded-lg border transition-all cursor-pointer ${
                          isLightMode ? "bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700" : "bg-slate-800/80 hover:bg-slate-700 border-slate-700 text-slate-300"
                        }`}
                      >
                        Next 30d
                      </button>
                      <button
                        onClick={() => applyDatePreset(60)}
                        className={`px-2 py-1 text-[10px] font-semibold rounded-lg border transition-all cursor-pointer ${
                          isLightMode ? "bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700" : "bg-slate-800/80 hover:bg-slate-700 border-slate-700 text-slate-300"
                        }`}
                      >
                        Next 60d
                      </button>
                      <button
                        onClick={() => applyDatePreset(90)}
                        className={`px-2 py-1 text-[10px] font-semibold rounded-lg border transition-all cursor-pointer ${
                          isLightMode ? "bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700" : "bg-slate-800/80 hover:bg-slate-700 border-slate-700 text-slate-300"
                        }`}
                      >
                        Next 90d
                      </button>
                      <button
                        onClick={applyYearPreset}
                        className={`px-2 py-1 text-[10px] font-semibold rounded-lg border transition-all cursor-pointer ${
                          isLightMode ? "bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700" : "bg-slate-800/80 hover:bg-slate-700 border-slate-700 text-slate-300"
                        }`}
                      >
                        Full Year
                      </button>
                      <button
                        onClick={() => {
                          setCustomStartDate(toISODateInput(minDate));
                          setCustomEndDate(toISODateInput(maxDate));
                          setSelectedMonths([]);
                        }}
                        className={`px-2 py-1 text-[10px] font-semibold rounded-lg border transition-all cursor-pointer col-span-2 ${
                          isLightMode ? "bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700" : "bg-slate-800/80 hover:bg-slate-700 border-slate-700 text-slate-300"
                        }`}
                      >
                        Full Dataset Range
                      </button>
                    </div>
                  </div>

                  {/* Popover Footer */}
                  <div className="mt-4 pt-3 border-t border-slate-200 dark:border-slate-800 flex justify-end gap-2">
                    <button
                      onClick={() => setIsDatePickerOpen(false)}
                      className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                        isLightMode
                          ? "bg-blue-600 text-white hover:bg-blue-700"
                          : "bg-emerald-600 text-white hover:bg-emerald-500"
                      }`}
                    >
                      Apply & View
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* View Toggles & Expand Controls */}
          <div className="flex items-center gap-2 shrink-0">
            {/* Expand / Collapse All */}
            <div className="flex items-center rounded-xl border border-slate-200 dark:border-slate-700/80 overflow-hidden">
              <button
                onClick={() => toggleAllRigsExpansion(true)}
                className={`px-2.5 py-1.5 text-[11px] font-semibold transition-all cursor-pointer ${
                  isLightMode
                    ? "bg-slate-50 hover:bg-slate-100 text-slate-700 border-r border-slate-200"
                    : "bg-[#1A1F2C] hover:bg-[#252C3D] text-slate-300 border-r border-slate-700/80"
                }`}
                title="Expand all rig sub-wells"
              >
                Expand All
              </button>
              <button
                onClick={() => toggleAllRigsExpansion(false)}
                className={`px-2.5 py-1.5 text-[11px] font-semibold transition-all cursor-pointer ${
                  isLightMode
                    ? "bg-slate-50 hover:bg-slate-100 text-slate-700"
                    : "bg-[#1A1F2C] hover:bg-[#252C3D] text-slate-300"
                }`}
                title="Collapse all rig sub-wells"
              >
                Collapse
              </button>
            </div>

            {/* View Mode (Gantt vs Table) */}
            <button
              onClick={() => setViewMode(viewMode === "gantt" ? "table" : "gantt")}
              className={`p-1.5 rounded-xl border transition-all cursor-pointer ${
                isLightMode
                  ? "bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700"
                  : "bg-[#1A1F2C] hover:bg-[#252C3D] border-slate-700/80 text-slate-300"
              }`}
              title={viewMode === "gantt" ? "Switch to Table View" : "Switch to Gantt View"}
            >
              {viewMode === "gantt" ? <SlidersHorizontal className="w-4 h-4" /> : <CalendarRange className="w-4 h-4" />}
            </button>
          </div>
        </div>

        {/* Feature: Dedicated Multi-Month Selection Bar */}
        {availableMonths.length > 0 && (
          <div
            className={`mt-4 pt-4 border-t flex flex-wrap items-center justify-between gap-3 ${
              isLightMode ? "border-slate-200/80" : "border-slate-800/80"
            }`}
          >
            <div className="flex flex-wrap items-center gap-2">
              <div className="flex items-center gap-1.5 text-xs font-bold text-slate-500 uppercase tracking-wider shrink-0 mr-1">
                <CalendarDays className="w-4 h-4 text-emerald-500" />
                <span>Filter Months:</span>
              </div>

              {/* Month Stepper (Prev / Multi-Month Dropdown / Next) */}
              <div className="flex items-center rounded-xl border border-slate-200 dark:border-slate-700/80 overflow-hidden shrink-0">
                <button
                  onClick={() => handleStepMonth(-1)}
                  className={`p-1.5 transition-all cursor-pointer ${
                    isLightMode
                      ? "bg-slate-50 hover:bg-slate-100 text-slate-600 border-r border-slate-200"
                      : "bg-[#1A1F2C] hover:bg-[#252C3D] text-slate-300 border-r border-slate-700/80"
                  }`}
                  title="Step Previous Month"
                >
                  <ChevronLeft className="w-3.5 h-3.5" />
                </button>

                {/* Multi-Month Dropdown Popover */}
                <div className="relative" ref={monthPickerRef}>
                  <button
                    onClick={() => setIsMonthPickerOpen(!isMonthPickerOpen)}
                    className={`px-3 py-1 text-xs font-semibold font-sans flex items-center gap-1.5 transition-all cursor-pointer ${
                      selectedMonths.length > 0
                        ? isLightMode
                          ? "bg-blue-50 text-blue-700 font-bold"
                          : "bg-emerald-950/80 text-emerald-300 font-bold"
                        : isLightMode
                        ? "bg-slate-50 text-slate-800 hover:bg-slate-100"
                        : "bg-[#1A1F2C] text-slate-200 hover:bg-[#252C3D]"
                    }`}
                  >
                    <span>
                      {selectedMonths.length === 0
                        ? "All Months"
                        : selectedMonths.length === 1
                        ? availableMonths.find((m) => m.key === selectedMonths[0])?.label || selectedMonths[0]
                        : `${selectedMonths.length} Months Selected`}
                    </span>
                    <ChevronDown className="w-3 h-3 opacity-60" />
                  </button>

                  {isMonthPickerOpen && (
                    <div
                      className={`absolute left-0 top-full mt-2 w-64 p-3 rounded-2xl border shadow-xl z-50 font-sans ${
                        isLightMode ? "bg-white border-slate-200" : "bg-[#161B26] border-slate-700"
                      }`}
                    >
                      <div className="flex items-center justify-between pb-2 border-b border-slate-200 dark:border-slate-800">
                        <span className="text-xs font-bold text-slate-800 dark:text-slate-200">
                          Select Months
                        </span>
                        <div className="flex items-center gap-2">
                          <button
                            onClick={() => {
                              setSelectedMonths(availableMonths.map((m) => m.key));
                              clearDateRange();
                            }}
                            className="text-[10px] font-semibold text-blue-600 dark:text-emerald-400 hover:underline"
                          >
                            All
                          </button>
                          <span className="text-slate-300 dark:text-slate-700">•</span>
                          <button
                            onClick={() => setSelectedMonths([])}
                            className="text-[10px] font-semibold text-slate-500 hover:underline"
                          >
                            Clear
                          </button>
                        </div>
                      </div>

                      <div className="divide-y divide-slate-100 dark:divide-slate-800/80 mt-2 max-h-60 overflow-y-auto custom-scrollbar">
                        {availableMonths.map((m) => {
                          const isChecked = selectedMonths.includes(m.key);
                          return (
                            <label
                              key={m.key}
                              className="flex items-center justify-between py-1.5 px-1 hover:bg-slate-50 dark:hover:bg-slate-800/50 rounded-lg cursor-pointer transition-colors"
                            >
                              <div className="flex items-center gap-2">
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => handleToggleMonth(m.key)}
                                  className="rounded text-emerald-600 focus:ring-emerald-500"
                                />
                                <span className="text-xs font-medium text-slate-800 dark:text-slate-200">
                                  {m.label}
                                </span>
                              </div>
                              <span className="text-[10px] font-mono text-slate-500">
                                {m.count} runs
                              </span>
                            </label>
                          );
                        })}
                      </div>
                    </div>
                  )}
                </div>

                <button
                  onClick={() => handleStepMonth(1)}
                  className={`p-1.5 transition-all cursor-pointer ${
                    isLightMode
                      ? "bg-slate-50 hover:bg-slate-100 text-slate-600 border-l border-slate-200"
                      : "bg-[#1A1F2C] hover:bg-[#252C3D] text-slate-300 border-l border-slate-700/80"
                  }`}
                  title="Step Next Month"
                >
                  <ChevronRight className="w-3.5 h-3.5" />
                </button>
              </div>

              {/* Quick Multi-Month Toggle Pills */}
              <div className="flex items-center gap-1.5 overflow-x-auto py-0.5 max-w-xl custom-scrollbar">
                <button
                  onClick={() => {
                    setSelectedMonths([]);
                    clearDateRange();
                  }}
                  className={`px-2.5 py-1 text-xs font-semibold rounded-lg border transition-all cursor-pointer shrink-0 ${
                    selectedMonths.length === 0 && !isCustomRangeActive
                      ? isLightMode
                        ? "bg-blue-600 text-white border-blue-600 shadow-xs"
                        : "bg-emerald-600 text-white border-emerald-600 shadow-xs"
                      : isLightMode
                      ? "bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100"
                      : "bg-[#1A1F2C] text-slate-400 border-slate-700/80 hover:text-white"
                  }`}
                >
                  All Timeline
                </button>

                {availableMonths.map((m) => {
                  const isSelected = selectedMonths.includes(m.key);
                  return (
                    <button
                      key={m.key}
                      onClick={() => handleToggleMonth(m.key)}
                      className={`px-2.5 py-1 text-xs font-medium rounded-lg border flex items-center gap-1.5 transition-all cursor-pointer shrink-0 ${
                        isSelected
                          ? isLightMode
                            ? "bg-blue-600 text-white border-blue-600 font-bold shadow-xs scale-105"
                            : "bg-emerald-600 text-white border-emerald-600 font-bold shadow-xs scale-105"
                          : isLightMode
                          ? "bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100 hover:border-slate-300"
                          : "bg-[#1A1F2C] text-slate-300 border-slate-700/80 hover:text-white hover:border-slate-600"
                      }`}
                    >
                      <span>{m.shortLabel}</span>
                      <span
                        className={`text-[10px] px-1 py-0.2 rounded-full font-mono ${
                          isSelected
                            ? "bg-white/20 text-white font-bold"
                            : isLightMode
                            ? "bg-slate-200 text-slate-600"
                            : "bg-slate-800 text-slate-400"
                        }`}
                      >
                        {m.count}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Active Status Badges & Quick Reset */}
            <div className="flex items-center gap-2 flex-wrap">
              {/* Custom Date Range Active Badge */}
              {isCustomRangeActive && (
                <div
                  className={`flex items-center gap-1.5 px-2.5 py-1 rounded-xl text-xs font-semibold border ${
                    isLightMode
                      ? "bg-blue-50 border-blue-200 text-blue-800"
                      : "bg-emerald-950/60 border-emerald-800/80 text-emerald-300"
                  }`}
                >
                  <Calendar className="w-3.5 h-3.5" />
                  <span>
                    Range: <strong>{formatDateDisplay(parseDateRobust(customStartDate)!)} → {formatDateDisplay(parseDateRobust(customEndDate)!)}</strong>
                  </span>
                  <button
                    onClick={clearDateRange}
                    className="p-0.5 rounded hover:bg-black/10 transition-colors cursor-pointer"
                    title="Clear Date Range"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              )}

              {/* Selected Multiple Months Badge */}
              {selectedMonths.length > 0 && !isCustomRangeActive && (
                <div
                  className={`flex items-center gap-1.5 px-2.5 py-1 rounded-xl text-xs font-semibold border ${
                    isLightMode
                      ? "bg-blue-50 border-blue-200 text-blue-800"
                      : "bg-emerald-950/60 border-emerald-800/80 text-emerald-300"
                  }`}
                >
                  <span>
                    Months ({selectedMonths.length}):{" "}
                    <strong>
                      {selectedMonths
                        .map((k) => availableMonths.find((m) => m.key === k)?.shortLabel || k)
                        .join(", ")}
                    </strong>
                  </span>
                  <button
                    onClick={() => setSelectedMonths([])}
                    className="p-0.5 rounded hover:bg-black/10 transition-colors cursor-pointer"
                    title="Clear Months Filter"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              )}

              {/* Selected Multiple Sizes Badge */}
              {selectedSizes.length > 0 && (
                <div
                  className={`flex items-center gap-1.5 px-2.5 py-1 rounded-xl text-xs font-semibold border ${
                    isLightMode
                      ? "bg-purple-50 border-purple-200 text-purple-800"
                      : "bg-purple-950/60 border-purple-800/80 text-purple-300"
                  }`}
                >
                  <span>
                    Sizes: <strong>{selectedSizes.join(", ")}</strong>
                  </span>
                  <button
                    onClick={() => setSelectedSizes([])}
                    className="p-0.5 rounded hover:bg-black/10 transition-colors cursor-pointer"
                    title="Clear Sizes Filter"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              )}

              {/* Horizon Stats if no active custom chips */}
              {!isCustomRangeActive && selectedMonths.length === 0 && selectedSizes.length === 0 && (
                <div className="text-[11px] text-slate-500 font-mono">
                  {totalStats.horizon} ({activeHorizonDays} days)
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* Main Gantt Timeline Workspace */}
      {viewMode === "gantt" ? (
        <div
          className={`rounded-2xl border overflow-hidden transition-all ${
            isLightMode
              ? "bg-white border-slate-200 shadow-sm"
              : "bg-[#121620] border-slate-700/60 shadow-xl"
          }`}
        >
          {/* Scrollable Gantt Canvas */}
          <div
            ref={timelineContainerRef}
            className="overflow-x-auto custom-scrollbar relative min-h-[450px] max-h-[750px]"
          >
            <div
              style={{
                minWidth: `${Math.max(
                  950,
                  timelineHeader.isFocusedView
                    ? timelineHeader.days.length * 36
                    : timelineHeader.months.length * 140
                )}px`
              }}
            >
              {/* Timeline Header Row (Months Tier + Optional Dynamic Days Sub-Header) */}
              <div
                className={`sticky top-0 z-30 border-b flex ${
                  isLightMode
                    ? "bg-slate-100/95 backdrop-blur-md border-slate-300 text-slate-700"
                    : "bg-[#0B0D14]/95 backdrop-blur-md border-slate-800 text-slate-300"
                }`}
              >
                {/* Fixed Rig / Well Name Column Header */}
                <div
                  className={`w-64 shrink-0 px-4 py-2.5 border-r font-sans font-bold text-xs uppercase tracking-wider flex items-center justify-between ${
                    isLightMode ? "border-slate-300 bg-slate-200/70" : "border-slate-800 bg-[#0F131D]"
                  } ${timelineHeader.isFocusedView ? "h-16" : "h-9"}`}
                >
                  <div className="flex items-center gap-2">
                    <Building2 className="w-4 h-4 text-emerald-500" />
                    <span>Rig / Well Units</span>
                  </div>
                  <span className="text-[10px] font-mono opacity-70">
                    {filteredRigGroups.length} rigs
                  </span>
                </div>

                {/* Dynamic Time Scale (Tiered: Month Row + Optional Day Dates Row) */}
                <div className={`flex-1 relative flex flex-col ${timelineHeader.isFocusedView ? "h-16" : "h-9"}`}>
                  {/* Top Tier: Month Milestone Scale */}
                  <div className={`w-full relative ${timelineHeader.isFocusedView ? "h-7 border-b border-slate-200 dark:border-slate-800/80" : "h-9"}`}>
                    {timelineHeader.months.map((m, idx) => (
                      <div
                        key={idx}
                        style={{
                          left: `${m.startPercent}%`,
                          width: `${m.widthPercent}%`
                        }}
                        className={`absolute top-0 bottom-0 px-3 font-mono font-bold text-xs truncate flex items-center border-l ${
                          isLightMode
                            ? "border-slate-300 text-slate-800 bg-slate-200/40"
                            : "border-slate-800 text-slate-200 bg-slate-900/50"
                        }`}
                      >
                        {m.label}
                      </div>
                    ))}
                  </div>

                  {/* Sub-Header Tier: Daily Dates for Selected Month(s) / Range */}
                  {timelineHeader.isFocusedView && (
                    <div className="w-full relative h-9 flex items-center">
                      {timelineHeader.days.map((d, idx) => (
                        <div
                          key={idx}
                          style={{
                            left: `${d.startPercent}%`,
                            width: `${d.widthPercent}%`
                          }}
                          title={d.runsSummaryText}
                          className={`absolute top-0 bottom-0 flex flex-col items-center justify-center border-l transition-colors select-none ${
                            d.hasRun
                              ? isLightMode
                                ? "bg-emerald-50 text-emerald-900 border-emerald-200 font-bold"
                                : "bg-emerald-950/60 text-emerald-300 border-emerald-800/80 font-bold"
                              : d.isWeekend
                              ? isLightMode
                                ? "bg-slate-200/50 text-slate-400 border-slate-200/80"
                                : "bg-slate-900/60 text-slate-500 border-slate-800/60"
                              : isLightMode
                              ? "border-slate-200/80 text-slate-600 hover:bg-slate-200/40"
                              : "border-slate-800/60 text-slate-400 hover:bg-slate-800/40"
                          }`}
                        >
                          {/* Day of Week */}
                          <span className="text-[8px] font-mono leading-none opacity-70">
                            {d.dayOfWeek}
                          </span>
                          
                          {/* Date Number with Run Indicator */}
                          <div className="flex items-center gap-0.5 mt-0.5">
                            <span className="text-[10px] font-mono font-bold leading-none">
                              {d.dayNum}
                            </span>
                            {d.hasRun && (
                              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 ring-1 ring-white/50" />
                            )}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              {/* Rows List (Grouped by Rig) */}
              <div className="divide-y divide-slate-200/70 dark:divide-slate-800/80 font-sans">
                {filteredRigGroups.map((rig) => {
                  const isExpanded = !!expandedRigs[rig.rigName];

                  return (
                    <div key={rig.rigName} className="group">
                      {/* Parent Rig Row Header */}
                      <div
                        className={`flex items-center min-h-[52px] transition-colors relative ${
                          isLightMode
                            ? "hover:bg-slate-50 bg-white"
                            : "hover:bg-[#161B26] bg-[#121620]"
                        }`}
                      >
                        {/* Rig Label Column */}
                        <div
                          className={`w-64 shrink-0 px-4 py-2 border-r flex items-center justify-between gap-2 select-none ${
                            isLightMode
                              ? "border-slate-200 bg-slate-50/50"
                              : "border-slate-800 bg-[#0F131D]/60"
                          }`}
                        >
                          <div className="flex items-center gap-2 min-w-0">
                            <button
                              onClick={() =>
                                setExpandedRigs((prev) => ({
                                  ...prev,
                                  [rig.rigName]: !prev[rig.rigName]
                                }))
                              }
                              className={`p-1 rounded-md transition-all cursor-pointer ${
                                isLightMode
                                  ? "hover:bg-slate-200 text-slate-600"
                                  : "hover:bg-slate-800 text-slate-400"
                              }`}
                              title={isExpanded ? "Collapse Wells" : "Expand Wells"}
                            >
                              <ChevronDown
                                className={`w-3.5 h-3.5 transition-transform duration-200 ${
                                  isExpanded ? "rotate-0" : "-rotate-90"
                                }`}
                              />
                            </button>
                            <div className="min-w-0">
                              <div className="flex items-center gap-1.5">
                                <span
                                  className={`font-bold text-xs truncate ${
                                    isLightMode ? "text-slate-900" : "text-white"
                                  }`}
                                  title={rig.rigName}
                                >
                                  {rig.rigName}
                                </span>
                                <span
                                  className={`px-1.5 py-0.2 rounded text-[9px] font-mono font-bold ${
                                    isLightMode
                                      ? "bg-slate-200 text-slate-700"
                                      : "bg-slate-800 text-slate-300"
                                  }`}
                                >
                                  {rig.totalWells} well{rig.totalWells > 1 ? "s" : ""}
                                </span>
                              </div>
                              <div className="text-[10px] text-slate-500 dark:text-slate-400 truncate">
                                {rig.client} • {rig.totalRuns} runs
                              </div>
                            </div>
                          </div>
                        </div>

                        {/* Rig Timeline Canvas Track (Aggregated Runs Preview) */}
                        <div className="flex-1 relative h-10 flex items-center">
                          {/* Rig Activity Bars for all Wells */}
                          {rig.allRuns.map((run) => {
                            const barStyles = getTimelineBarStyles(run.startDate, run.endDate);
                            if (!barStyles.isVisible) return null;
                            const colors = getSectionBadgeColor(run.section);

                            return (
                              <motion.div
                                key={run.id}
                                initial={{ opacity: 0, scaleY: 0.8 }}
                                animate={{ opacity: 1, scaleY: 1 }}
                                style={{
                                  left: barStyles.left,
                                  width: barStyles.width
                                }}
                                onClick={() => setSelectedRunDetail(run)}
                                className="absolute h-7 p-0.5 cursor-pointer z-10"
                                title={`${run.well} (${run.section}) - ${run.startDateStr}`}
                              >
                                <div
                                  className={`w-full h-full rounded-md shadow-xs transition-all hover:scale-105 hover:z-20 flex items-center justify-center px-1 text-[9px] font-bold text-white overflow-hidden ${colors.barClass}`}
                                >
                                  <span className="truncate drop-shadow-xs">
                                    {run.well} • {run.section}
                                  </span>
                                </div>
                              </motion.div>
                            );
                          })}
                        </div>
                      </div>

                      {/* Expandable Wells Sub-Rows */}
                      <AnimatePresence>
                        {isExpanded && (
                          <motion.div
                            initial={{ opacity: 0, height: 0 }}
                            animate={{ opacity: 1, height: "auto" }}
                            exit={{ opacity: 0, height: 0 }}
                            className={`divide-y divide-slate-200/50 dark:divide-slate-800/50 ${
                              isLightMode ? "bg-slate-50/80" : "bg-[#0B0E16]/80"
                            }`}
                          >
                            {rig.wells.map((well) => (
                              <div
                                key={well.wellName}
                                className={`flex items-center min-h-[44px] transition-colors ${
                                  isLightMode ? "hover:bg-slate-100/60" : "hover:bg-[#141924]"
                                }`}
                              >
                                {/* Well Label Column */}
                                <div
                                  className={`w-64 shrink-0 pl-10 pr-4 py-2 border-r flex items-center justify-between gap-2 ${
                                    isLightMode ? "border-slate-200" : "border-slate-800/80"
                                  }`}
                                >
                                  <div className="min-w-0">
                                    <div className="flex items-center gap-1.5">
                                      <span className="w-1.5 h-1.5 rounded-full bg-blue-500 shrink-0" />
                                      <span
                                        className={`font-semibold text-xs truncate ${
                                          isLightMode ? "text-slate-800" : "text-slate-200"
                                        }`}
                                        title={well.wellName}
                                      >
                                        {well.wellName}
                                      </span>
                                    </div>
                                    <div className="text-[10px] text-slate-500 font-mono">
                                      {well.runs.length} run{well.runs.length > 1 ? "s" : ""}
                                    </div>
                                  </div>

                                  <span className="text-[10px] font-mono text-slate-500 shrink-0">
                                    {formatDateDisplay(well.earliestDate)}
                                  </span>
                                </div>

                                {/* Well Timeline Canvas Track */}
                                <div className="flex-1 relative h-9 flex items-center">
                                  {/* Individual Section Runs on this Well */}
                                  {well.runs.map((run) => {
                                    const barStyles = getTimelineBarStyles(run.startDate, run.endDate);
                                    if (!barStyles.isVisible) return null;
                                    const colors = getSectionBadgeColor(run.section);

                                    return (
                                      <div
                                        key={run.id}
                                        style={{
                                          left: barStyles.left,
                                          width: barStyles.width
                                        }}
                                        onClick={() => setSelectedRunDetail(run)}
                                        className="absolute h-6 p-0.5 cursor-pointer z-10"
                                        title={`${run.well} (${run.section}): Mobilized ${run.startDateStr}`}
                                      >
                                        <div
                                          className={`w-full h-full rounded-md shadow-xs transition-all hover:scale-105 hover:z-20 flex items-center justify-between px-1.5 text-[9px] font-medium text-white overflow-hidden ${colors.barClass}`}
                                        >
                                          <span className="font-bold truncate drop-shadow-xs">
                                            {run.section}
                                          </span>
                                          <span className="font-mono text-[8px] opacity-90 truncate ml-0.5 hidden sm:inline">
                                            {run.startDate.getDate()}
                                          </span>
                                        </div>
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            ))}
                          </motion.div>
                        )}
                      </AnimatePresence>
                    </div>
                  );
                })}

                {filteredRigGroups.length === 0 && (
                  <div className="py-16 text-center">
                    <Building2 className="w-10 h-10 text-slate-400 mx-auto mb-2 opacity-50" />
                    <h3 className="font-bold text-sm text-slate-700 dark:text-slate-300">
                      No matching rig schedules found
                    </h3>
                    <p className="text-xs text-slate-500 mt-1">
                      Try adjusting your search query, size selection, or month filters.
                    </p>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      ) : (
        /* Alternative Tabular Schedule View */
        <div
          className={`rounded-2xl border overflow-hidden ${
            isLightMode ? "bg-white border-slate-200 shadow-sm" : "bg-[#121620] border-slate-700/60"
          }`}
        >
          <div className="overflow-x-auto custom-scrollbar">
            <table className="w-full text-left text-xs">
              <thead
                className={`border-b font-sans font-bold uppercase tracking-wider text-[10px] ${
                  isLightMode
                    ? "bg-slate-50 border-slate-200 text-slate-600"
                    : "bg-[#0B0D14] border-slate-800 text-slate-400"
                }`}
              >
                <tr>
                  <th className="py-3 px-4">Rig</th>
                  <th className="py-3 px-4">Well Name</th>
                  <th className="py-3 px-4">Tool Section</th>
                  <th className="py-3 px-4">Mobilization Date</th>
                  <th className="py-3 px-4">Configuration</th>
                  <th className="py-3 px-4">Qty</th>
                  <th className="py-3 px-4">Footage</th>
                  <th className="py-3 px-4">Client</th>
                  <th className="py-3 px-4">Shop Date</th>
                  <th className="py-3 px-4 text-right">Action</th>
                </tr>
              </thead>
              <tbody
                className={`divide-y ${
                  isLightMode
                    ? "divide-slate-200 text-slate-800"
                    : "divide-slate-800 text-slate-200"
                }`}
              >
                {filteredRigGroups.flatMap((rig) =>
                  rig.allRuns.map((run) => {
                    const colors = getSectionBadgeColor(run.section);
                    return (
                      <tr
                        key={run.id}
                        onClick={() => setSelectedRunDetail(run)}
                        className={`transition-colors cursor-pointer ${
                          isLightMode ? "hover:bg-slate-50" : "hover:bg-slate-800/40"
                        }`}
                      >
                        <td className="py-2.5 px-4 font-bold text-emerald-600 dark:text-emerald-400">
                          {run.rig}
                        </td>
                        <td className="py-2.5 px-4 font-semibold">{run.well}</td>
                        <td className="py-2.5 px-4">
                          <span
                            className={`px-2 py-0.5 rounded text-[10px] font-bold border ${colors.bg}`}
                          >
                            {run.section}
                          </span>
                        </td>
                        <td className="py-2.5 px-4 font-mono font-semibold text-blue-600 dark:text-blue-400">
                          {run.startDateStr}
                        </td>
                        <td className="py-2.5 px-4 font-mono text-[11px] max-w-[200px] truncate" title={run.toolConfig}>
                          {run.toolConfig}
                        </td>
                        <td className="py-2.5 px-4 font-mono">{run.qty}</td>
                        <td className="py-2.5 px-4 font-mono">
                          {run.footage ? `${run.footage.toLocaleString()} ft` : "—"}
                        </td>
                        <td className="py-2.5 px-4 text-slate-500">{run.client}</td>
                        <td className="py-2.5 px-4 text-slate-500">{run.shopDate || "—"}</td>
                        <td className="py-2.5 px-4 text-right">
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              setSelectedRunDetail(run);
                            }}
                            className={`px-2 py-1 rounded text-[10px] font-bold transition-all ${
                              isLightMode
                                ? "bg-slate-100 hover:bg-slate-200 text-slate-700"
                                : "bg-slate-800 hover:bg-slate-700 text-slate-200"
                            }`}
                          >
                            Details
                          </button>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Detailed Run Inspection Modal */}
      <AnimatePresence>
        {selectedRunDetail && (
          <div
            className="fixed inset-0 bg-slate-950/80 backdrop-blur-xs flex items-center justify-center z-50 p-4"
            onClick={() => setSelectedRunDetail(null)}
          >
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 10 }}
              onClick={(e) => e.stopPropagation()}
              className={`max-w-lg w-full rounded-2xl border shadow-2xl overflow-hidden font-sans ${
                isLightMode ? "bg-white border-slate-200" : "bg-[#121620] border-slate-700"
              }`}
            >
              {/* Modal Header */}
              <div
                className={`p-5 border-b flex items-center justify-between ${
                  isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#0B0D14] border-slate-800"
                }`}
              >
                <div className="flex items-center gap-3">
                  <div className="p-2.5 rounded-xl bg-emerald-500/10 text-emerald-500 border border-emerald-500/20">
                    <Building2 className="w-5 h-5" />
                  </div>
                  <div>
                    <h3
                      className={`text-base font-bold ${
                        isLightMode ? "text-slate-900" : "text-white"
                      }`}
                    >
                      {selectedRunDetail.rig} • {selectedRunDetail.well}
                    </h3>
                    <p className="text-xs text-slate-500">
                      Scheduled Operational Run Details
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => setSelectedRunDetail(null)}
                  className={`p-1.5 rounded-lg border transition-all cursor-pointer ${
                    isLightMode
                      ? "bg-white hover:bg-slate-100 border-slate-200 text-slate-500"
                      : "bg-slate-800 hover:bg-slate-700 border-slate-700 text-slate-400"
                  }`}
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              {/* Modal Content */}
              <div className="p-5 space-y-4 text-xs">
                <div className="grid grid-cols-2 gap-3">
                  <div
                    className={`p-3 rounded-xl border ${
                      isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#0B0D14] border-slate-800"
                    }`}
                  >
                    <div className="text-[10px] text-slate-500 uppercase font-bold">
                      Mobilization Date
                    </div>
                    <div className="text-sm font-mono font-bold text-emerald-600 dark:text-emerald-400 mt-0.5">
                      {selectedRunDetail.startDateStr}
                    </div>
                  </div>

                  <div
                    className={`p-3 rounded-xl border ${
                      isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#0B0D14] border-slate-800"
                    }`}
                  >
                    <div className="text-[10px] text-slate-500 uppercase font-bold">
                      Tool Section Size
                    </div>
                    <div className="text-sm font-mono font-bold text-blue-600 dark:text-blue-400 mt-0.5">
                      {selectedRunDetail.section} ({selectedRunDetail.holeSize || 'Standard'})
                    </div>
                  </div>

                  <div
                    className={`p-3 rounded-xl border ${
                      isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#0B0D14] border-slate-800"
                    }`}
                  >
                    <div className="text-[10px] text-slate-500 uppercase font-bold">
                      Planned Footage
                    </div>
                    <div className="text-sm font-mono font-bold text-amber-600 dark:text-amber-400 mt-0.5">
                      {selectedRunDetail.footage ? `${selectedRunDetail.footage.toLocaleString()} ft` : "Not specified"}
                    </div>
                  </div>

                  <div
                    className={`p-3 rounded-xl border ${
                      isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#0B0D14] border-slate-800"
                    }`}
                  >
                    <div className="text-[10px] text-slate-500 uppercase font-bold">
                      Operating Client
                    </div>
                    <div className="text-sm font-semibold text-slate-800 dark:text-slate-200 mt-0.5">
                      {selectedRunDetail.client}
                    </div>
                  </div>
                </div>

                {/* Tool Configuration */}
                <div
                  className={`p-3 rounded-xl border ${
                    isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#0B0D14] border-slate-800"
                  }`}
                >
                  <div className="text-[10px] text-slate-500 uppercase font-bold mb-1">
                    Tool Configuration & Payload
                  </div>
                  <div className="font-mono text-xs font-semibold text-slate-800 dark:text-slate-200 bg-white dark:bg-slate-900 p-2 rounded-lg border border-slate-200 dark:border-slate-800">
                    {selectedRunDetail.toolConfig} (Quantity: {selectedRunDetail.qty})
                  </div>
                </div>

                {/* Shop Date / Delivery Status */}
                {selectedRunDetail.shopDate && (
                  <div
                    className={`p-3 rounded-xl border ${
                      isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#0B0D14] border-slate-800"
                    }`}
                  >
                    <div className="text-[10px] text-slate-500 uppercase font-bold mb-1">
                      Shop Delivery Status
                    </div>
                    <div className="font-sans text-xs text-slate-700 dark:text-slate-300">
                      {selectedRunDetail.shopDate}
                    </div>
                  </div>
                )}

                {/* Comments */}
                {selectedRunDetail.comments && (
                  <div
                    className={`p-3 rounded-xl border ${
                      isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#0B0D14] border-slate-800"
                    }`}
                  >
                    <div className="text-[10px] text-slate-500 uppercase font-bold mb-1">
                      Operational Comments
                    </div>
                    <div className="font-sans text-xs text-slate-700 dark:text-slate-300 italic">
                      "{selectedRunDetail.comments}"
                    </div>
                  </div>
                )}
              </div>

              {/* Modal Footer */}
              <div
                className={`p-4 border-t flex justify-end gap-2 ${
                  isLightMode ? "bg-slate-50 border-slate-200" : "bg-[#0B0D14] border-slate-800"
                }`}
              >
                <button
                  onClick={() => setSelectedRunDetail(null)}
                  className={`px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                    isLightMode
                      ? "bg-slate-900 text-white hover:bg-slate-800"
                      : "bg-emerald-600 text-white hover:bg-emerald-500"
                  }`}
                >
                  Close Inspection
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
};

export default RigWellGanttChart;
