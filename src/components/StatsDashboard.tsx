import React, { useMemo, useState, useEffect } from "react";
import { motion } from "motion/react";
import { Compass, Footprints, Grid, Layers, ClipboardCheck, ArrowRight, Search, SlidersHorizontal, Info, ArrowUpDown, Check, ChevronDown, X, Calendar } from "lucide-react";
import { RowData, Abbreviation, getSectionFromHoleSize } from "../types";
import BorderGlow from "./BorderGlow";
import { EmptyState } from "./ui/empty-state";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Cell
} from "recharts";

interface StatsDashboardProps {
  data: RowData[];
  abbreviations: Abbreviation[];
  onAbbreviationClick: (abb: string) => void;
  selectedTool?: string;
  isLightMode?: boolean;
  view?: "all" | "profile" | "allocations" | "rig-breakdown";
  selectedToolForRigPanel?: string | null;
  onSelectToolForRigPanel?: (toolName: string | null) => void;
}

// Precise exact-token matching helper that isolates abbreviation codes to prevent substring overlaps
const matchesTool = (config: string, abb: Abbreviation): boolean => {
  if (!config) return false;
  
  const configUpper = config.toUpperCase();
  const abbUpper = abb.abbreviation.toUpperCase();
  const fullNameUpper = abb.fullName.toUpperCase();

  // Dedicated check for IC BM tool (matches only IC BM or IC(BM))
  if (fullNameUpper === "IC BM" || abbUpper === "IC BM" || abbUpper === "IC(BM)" || abbUpper === "ICBM" || abbUpper === "IC-BM") {
    return /IC[\s\-_]*\(?BM\)?\b|ICBM\b/i.test(configUpper);
  }

  // Dedicated check for iCruise tool (contains all types of IC except BM)
  if (fullNameUpper === "ICRUISE" || abbUpper === "IC" || abbUpper === "ICRUISE") {
    // Strip out IC BM / IC(BM) so it is not confused with iCruise
    const configWithoutICBM = configUpper.replace(/IC[\s\-_]*\(?BM\)?\b|ICBM\b/gi, " ");
    const tokensWithoutICBM = configWithoutICBM.replace(/[()\[\]\-\/\\,;]/g, " ").split(/\s+/).filter(Boolean);
    return tokensWithoutICBM.includes("IC") || tokensWithoutICBM.includes("ICRUISE") || /ICRUISE/i.test(configWithoutICBM);
  }
  
  // Dedicated check for FlowStar tool
  if (fullNameUpper === "FLOWSTAR" || abbUpper === "FST" || abbUpper === "FLOWSTAR" || abbUpper === "FLOW STAR" || abbUpper === "FLOW") {
    if (/FLOW[\s\-_]*STAR|FST\d*|FLOWSTAR\d*/i.test(configUpper)) return true;
    const tokens = configUpper.replace(/[()\[\]\-\/\\,;]/g, " ").split(/\s+/).filter(Boolean);
    if (tokens.some((t) => t.startsWith("FST") || t.startsWith("FLOWSTAR") || t === "FLOW")) return true;
  }

  // Dedicated check for EarthStar tool
  if (fullNameUpper === "EARTHSTAR" || abbUpper.startsWith("ES")) {
    const tokens = configUpper.replace(/[()\[\]\-\/\\,;]/g, " ").split(/\s+/).filter(Boolean);
    if (tokens.some((t) => t.startsWith("ES") || t.startsWith("EXRX") || t.startsWith("EARTHSTAR"))) return true;
  }

  // Dedicated check for EM Battery tool
  if (fullNameUpper === "EM BATTERY" || abbUpper.startsWith("EMBATT")) {
    const tokens = configUpper.replace(/[()\[\]\-\/\\,;]/g, " ").split(/\s+/).filter(Boolean);
    if (tokens.some((t) => t.startsWith("EMBATT") || t.startsWith("EM_BATT"))) return true;
  }

  // Replace brackets, parentheses, hyphens, and slashes with spaces to isolate tokens
  const normalizedConfig = configUpper.replace(/[()\[\]\-\/\\,;]/g, " ");
  
  // Split into individual tokens
  const tokens = normalizedConfig.split(/\s+/).filter(Boolean);
  
  // Check for exact match in tokens for abbreviation OR full name
  if (tokens.includes(abbUpper) || tokens.includes(fullNameUpper) || tokens.includes(fullNameUpper.replace(/\s+/g, ""))) {
    return true;
  }
  
  // Support multi-word abbreviations (e.g., "IC BUS MASTER")
  const abbTokens = abbUpper.split(/\s+/).filter(Boolean);
  if (abbTokens.length > 1) {
    for (let i = 0; i <= tokens.length - abbTokens.length; i++) {
      let sequenceMatch = true;
      for (let j = 0; j < abbTokens.length; j++) {
        if (tokens[i + j] !== abbTokens[j]) {
          sequenceMatch = false;
          break;
        }
      }
      if (sequenceMatch) return true;
    }
  }
  
  return false;
};

// Detailed proper names mapping for LWD/MWD tool descriptions
const PROPER_NAMES: { [key: string]: string } = {
  "IC": "iCruise RSS",
  "ICRUISE": "iCruise RSS",
  "IC BM": "iCruise Bus Master (IC BM)",
  "IC(BM)": "iCruise Bus Master (IC BM)",
  "ICBM": "iCruise Bus Master (IC BM)",
  "IC-BM": "iCruise Bus Master (IC BM)",
  "BST": "BaseStar Acoustic",
  "BASE": "BaseStar Acoustic",
  "BASESTAR": "BaseStar Acoustic",
  "RST": "ResiStar Resistivity",
  "RESISTAR": "ResiStar Resistivity",
  "RS": "ResiStar Resistivity",
  "STS": "StrataStar Mechanics",
  "LST": "LithoStar Density",
  "LITHO": "LithoStar Density",
  "LITHOSTAR": "LithoStar Density",
  "BATT": "Lithium Battery Pack",
  "BATTERY": "Lithium Battery Pack",
  "BATT1": "Lithium Battery Pack",
  "BATT2": "Lithium Battery Pack",
  "ISTARBATT": "Lithium Battery Pack",
  "ISTARBATT1": "Lithium Battery Pack",
  "ISTARBATT2": "Lithium Battery Pack",
  "EM BATT": "EM Battery Pack",
  "EMBATT": "EM Battery Pack",
  "EM BATTERY": "EM Battery Pack",
  "EMBATTERY": "EM Battery Pack",
  "EM BATT1": "EM Battery Pack",
  "EM BATT2": "EM Battery Pack",
  "EM BATT 1": "EM Battery Pack",
  "EM BATT 2": "EM Battery Pack",
  "EM-BATT": "EM Battery Pack",
  "EM-BATTERY": "EM Battery Pack",
  "ILS": "Integrated Logging System",
  "HAL50ILS": "Integrated Logging System",
  "XBAT": "XBAT Sonic Acoustic",
  "AFR": "Azimuthal Focus Resistivity",
  "LC": "Lithology Density Caliper",
  "LCS": "Lithology Density Caliper",
  "UC": "Ultrasonic Caliper (UCS)",
  "UCS": "Ultrasonic Caliper (UCS)",
  "USC": "Ultrasonic Caliper (UCS)",
  "GEOTAP": "GeoTap Formation Pressure",
  "GT": "GeoTap Formation Pressure",
  "ALD": "Azimuthal Lithodensity",
  "GWD": "Gyro While Drilling",
  "DGR": "Dual Gamma Ray",
  "PWD": "Pressure While Drilling",
  "DUMMYFLEX": "Dummy Flex Spacer",
  "DF": "Dummy Flex Spacer",
  "CALIPER": "BST Caliper",
  "BST CALIPER": "BST Caliper",
  "BSTCALIPER": "BST Caliper",
  "ADR": "Azimuthal Deep Resistivity",
  "PULSESTAR": "PulseStar Telemetry",
  "PST": "PulseStar Telemetry",
  "PULSAR": "Pulsar Telemetry",
  "FST": "FlowStar Flow Sensor",
  "FLOWSTAR": "FlowStar Flow Sensor",
  "FLOW STAR": "FlowStar Flow Sensor",
  "FLOW-STAR": "FlowStar Flow Sensor",
  "GYROSTAR": "GyroStar Directional",
  "EARTHSTAR": "EarthStar Telemetry",
  "ES": "EarthStar Telemetry",
  "ES TX": "EarthStar Telemetry",
  "ESTX": "EarthStar Telemetry",
  "ESTX1": "EarthStar Telemetry",
  "ESTX2": "EarthStar Telemetry",
  "ES RX": "EarthStar Telemetry",
  "ESRX": "EarthStar Telemetry",
  "ESRX1": "EarthStar Telemetry",
  "ESRX2": "EarthStar Telemetry",
  "FDC": "Formation Density Caliper",
  "HCIM": "Hydrocarbon Zone Evaluation",
  "PCDC": "Pulse Control & Gamma",
  "PCGK": "Power-Pulse & Gamma",
  "P4M": "Propagation Resistivity",
  "DHS": "Downhole Sensor Sub"
};

// Concise descriptive proper names for clean XAxis display
const SHORT_PROPER_NAMES: { [key: string]: string } = {
  "IC": "Steerable RSS",
  "ICRUISE": "Steerable RSS",
  "IC BM": "Bus Master (IC BM)",
  "IC(BM)": "Bus Master (IC BM)",
  "ICBM": "Bus Master (IC BM)",
  "IC-BM": "Bus Master (IC BM)",
  "BST": "Acoustic Sonic",
  "BASE": "Acoustic Sonic",
  "BASESTAR": "Acoustic Sonic",
  "RST": "Resistivity",
  "RESISTAR": "Resistivity",
  "RS": "Resistivity",
  "STS": "Mechanics Sub",
  "LST": "Density / Litho",
  "LITHO": "Density / Litho",
  "LITHOSTAR": "Density / Litho",
  "BATT": "Lithium Battery",
  "BATTERY": "Lithium Battery",
  "BATT1": "Lithium Battery",
  "BATT2": "Lithium Battery",
  "ISTARBATT": "Lithium Battery",
  "ISTARBATT1": "Lithium Battery",
  "ISTARBATT2": "Lithium Battery",
  "EM BATT": "EM Battery",
  "EMBATT": "EM Battery",
  "EM BATTERY": "EM Battery",
  "EMBATTERY": "EM Battery",
  "EM BATT1": "EM Battery",
  "EM BATT2": "EM Battery",
  "EM BATT 1": "EM Battery",
  "EM BATT 2": "EM Battery",
  "EM-BATT": "EM Battery",
  "EM-BATTERY": "EM Battery",
  "ILS": "Logging System",
  "HAL50ILS": "Logging System",
  "XBAT": "Sonic Acoustic",
  "AFR": "Focus Resistivity",
  "LC": "Litho Density",
  "LCS": "Litho Density",
  "UC": "Ultrasonic Caliper",
  "UCS": "Ultrasonic Caliper",
  "USC": "Ultrasonic Caliper",
  "GEOTAP": "Formation Pressure",
  "GT": "Formation Pressure",
  "ALD": "Azimuthal Density",
  "GWD": "Gyro Survey",
  "DGR": "Dual Gamma Ray",
  "PWD": "Annular Pressure",
  "DUMMYFLEX": "Dummy Flex Spacer",
  "DF": "Dummy Flex Spacer",
  "CALIPER": "BST Caliper",
  "BST CALIPER": "BST Caliper",
  "BSTCALIPER": "BST Caliper",
  "ADR": "Deep Resistivity",
  "PULSESTAR": "Pulse Telemetry",
  "PST": "Pulse Telemetry",
  "PULSAR": "Pulsar Telemetry",
  "FST": "Flow Sensor",
  "FLOWSTAR": "Flow Sensor",
  "FLOW STAR": "Flow Sensor",
  "FLOW-STAR": "Flow Sensor",
  "GYROSTAR": "Gyro Directional",
  "EARTHSTAR": "EM Telemetry",
  "ES": "EM Telemetry",
  "ES TX": "EM Telemetry",
  "ESTX": "EM Telemetry",
  "ESTX1": "EM Telemetry",
  "ESTX2": "EM Telemetry",
  "ES RX": "EM Telemetry",
  "ESRX": "EM Telemetry",
  "ESRX1": "EM Telemetry",
  "ESRX2": "EM Telemetry",
  "FDC": "Density Caliper",
  "HCIM": "Hydrocarbon Zone",
  "PCDC": "Pulse & Gamma",
  "PCGK": "Pulse & Gamma",
  "P4M": "Prop Resistivity",
  "DHS": "Downhole Sensor"
};

// Custom interactive XAxis Tick Renderer presenting a neat single-line label setup (Abbreviation + Proper Name)
const CustomXAxisTick = (props: any) => {
  const { x, y, payload, isLightMode } = props;
  const abb = payload.value;
  let displayAbb = abb && abb.toUpperCase() === "CALIPER" ? "BST Caliper" : abb;

  if (displayAbb) {
    const upper = displayAbb.toUpperCase();
    if (upper === "RESISTAR" || upper === "RST") {
      displayAbb = "RST";
    } else if (upper === "PULSESTAR" || upper === "PST") {
      displayAbb = "PST";
    } else if (upper === "GT" || upper === "GEOTAP") {
      displayAbb = "GeoTap";
    } else if (upper.startsWith("ES") || upper.startsWith("EXRX") || upper.startsWith("EARTHSTAR")) {
      displayAbb = "EarthStar";
    } else if (upper.startsWith("EMBATT") || upper.startsWith("EM BATT") || upper === "EM BATTERY") {
      displayAbb = "EM Battery";
    } else if (upper === "FLOWSTAR" || upper === "FLOW STAR" || upper === "FST") {
      displayAbb = "FlowStar";
    }
  }

  return (
    <g transform={`translate(${x},${y + 8})`}>
      <g transform="rotate(-30)">
        <text
          x={0}
          y={0}
          textAnchor="end"
          fill={isLightMode ? "#0F172A" : "#F5F7FA"}
          fontSize={9}
          fontWeight={600}
          fontFamily="Inter, sans-serif"
        >
          {displayAbb}
        </text>
      </g>
    </g>
  );
};

const SECTION_COLORS: { [key: string]: string } = {
  "4.75in": "#3B82F6",   // Blue (Electric Blue)
  "6.75in": "#10B981",   // Green (Emerald)
  "8in": "#E024A5",      // Magenta
  "9.5in": "#F59E0B",    // Yellow (Amber)
};

const SECTION_COLORS_LIGHT: { [key: string]: string } = {
  "4.75in": "#2563EB",   // Deep Sapphire Blue
  "6.75in": "#059669",   // Deep Emerald Green
  "8in": "#C026D3",      // Deep Magenta
  "9.5in": "#D97706",    // Rich Amber
};

const OTHER_COLORS = [
  "#06B6D4", // Cyan
  "#EC4899", // Pink
  "#14B8A6", // Teal
  "#F97316", // Orange
  "#64748B", // Slate
];

const OTHER_COLORS_LIGHT = [
  "#0891B2", // Rich Cyan
  "#DB2777", // Rich Pink
  "#0D9488", // Rich Teal
  "#EA580C", // Rich Orange
  "#475569", // Slate
];

const getSectionColor = (sec: string, index: number, isLightMode?: boolean) => {
  if (isLightMode) {
    if (SECTION_COLORS_LIGHT[sec]) {
      return SECTION_COLORS_LIGHT[sec];
    }
    return OTHER_COLORS_LIGHT[index % OTHER_COLORS_LIGHT.length];
  }
  if (SECTION_COLORS[sec]) {
    return SECTION_COLORS[sec];
  }
  return OTHER_COLORS[index % OTHER_COLORS.length];
};

const getNormalizedSection = (secRaw: string): string => {
  return getSectionFromHoleSize(secRaw);
};

const getSafeSectionKey = (sec: string): string => {
  return sec
    .replace(/\./g, "_")
    .replace(/"/g, "")
    .replace(/'/g, "")
    .replace(/\s+/g, "_")
    .replace(/[^a-zA-Z0-9_]/g, "");
};

const CustomTooltip = ({ active, payload, isLightMode, profileSortBy }: any) => {
  if (active && payload && payload.length) {
    const data = payload[0].payload;
    let properName = data.abbreviation && data.abbreviation.toUpperCase() === "CALIPER" ? "BST Caliper" : data.abbreviation;
    if (properName) {
      const upper = properName.toUpperCase();
      if (upper === "RESISTAR" || upper === "RST") {
        properName = "RST";
      } else if (upper === "PULSESTAR" || upper === "PST") {
        properName = "PST";
      } else if (upper === "GT" || upper === "GEOTAP") {
        properName = "GeoTap";
      } else if (upper.startsWith("ES") || upper.startsWith("EXRX") || upper.startsWith("EARTHSTAR")) {
        properName = "EarthStar";
      } else if (upper.startsWith("EMBATT") || upper.startsWith("EM BATT") || upper === "EM BATTERY") {
        properName = "EM Battery";
      } else if (upper === "FLOWSTAR" || upper === "FLOW STAR" || upper === "FST") {
        properName = "FlowStar";
      }
    }
    
    const bgClass = isLightMode ? "bg-white border-gray-200 text-gray-900" : "bg-[#1A1D21] border-white/10 text-[#F5F7FA]";
    const textMutedClass = isLightMode ? "text-gray-500" : "text-[#9AA0A6]";
    const borderClass = isLightMode ? "border-gray-100" : "border-white/5";
    const titleClass = isLightMode ? "text-gray-900 font-sans" : "text-[#F5F7FA] font-sans";
    
    return (
      <div className={`${bgClass} border p-3.5 rounded-lg shadow-xl text-left max-w-xs font-sans`}>
        <p className={`text-xs font-bold ${titleClass} mb-2.5 border-b ${borderClass} pb-2`}>
          {properName}
        </p>

        {/* Section breakdown */}
        {data.sections && Object.keys(data.sections).length > 0 && (
          <div className="space-y-1.5">
            <p className={`text-[10px] uppercase font-bold tracking-wider ${textMutedClass} mb-2`}>
              Tool Sizes Distribution
            </p>
            <div className="space-y-1.5">
              {Object.entries(data.sections)
                .sort((a, b) => {
                  const valA = (a[1] as any).count;
                  const valB = (b[1] as any).count;
                  return valB - valA;
                })
                .map(([sec, val]: any, idx) => {
                  const secColor = getSectionColor(sec, idx, isLightMode);
                  const isRigs = profileSortBy === "count";
                  const valueStr = isRigs 
                    ? `${val.count} rig${val.count > 1 ? 's' : ''}`
                    : `${val.footage.toLocaleString()} ft`;
                    
                  return (
                    <div key={sec} className="flex items-center justify-between text-[11px]">
                      <div className="flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 rounded-md shrink-0" style={{ backgroundColor: secColor }} />
                        <span className={isLightMode ? "text-slate-700 font-medium" : "text-slate-300 font-medium"}>{sec}</span>
                      </div>
                      <span className="font-mono font-bold text-blue-500">{valueStr}</span>
                    </div>
                  );
                })}
            </div>
          </div>
        )}
        <div className={`mt-3 pt-2 border-t ${borderClass} text-center text-[10px] text-emerald-500 font-bold tracking-tight flex items-center justify-center gap-1`}>
          <span>👆 Click bar to navigate to its Rig Breakdown page</span>
        </div>
      </div>
    );
  }
  return null;
};


function StatsDashboard(props: StatsDashboardProps) {
  const { data, abbreviations, onAbbreviationClick, selectedTool, isLightMode, view = "profile" } = props;
  
  // Local filters & settings for the upgraded Tools Forecast panel
  const [profileSearch, setProfileSearch] = useState("");
  const [profileCategories, setProfileCategories] = useState<string[]>([]);
  const [isCategoryDropdownOpen, setIsCategoryDropdownOpen] = useState(false);
  const profileSortBy = "count";
  const [showAllTools, setShowAllTools] = useState(false);
  const [expandedRigDates, setExpandedRigDates] = useState<{ [rigName: string]: boolean }>({});
  
  const [localSelectedTool, setLocalSelectedTool] = useState<string | null>(null);
  const selectedToolForRigPanel = props.selectedToolForRigPanel !== undefined ? props.selectedToolForRigPanel : localSelectedTool;
  const setSelectedToolForRigPanel = (val: string | null) => {
    if (props.onSelectToolForRigPanel) {
      props.onSelectToolForRigPanel(val);
    } else {
      setLocalSelectedTool(val);
    }
  };

  // Automatically collapse all date breakdowns whenever the selected tool changes
  useEffect(() => {
    setExpandedRigDates({});
  }, [selectedToolForRigPanel, selectedTool]);

  // 1. Core aggregates
  const stats = useMemo(() => {
    let totalFootage = 0;
    let totalQty = 0;
    const uniqueRigs = new Set<string>();
    const uniqueWells = new Set<string>();
    const clients = new Set<string>();

    data.forEach((row) => {
      if (row.footage) totalFootage += row.footage;
      totalQty += row.qty || 1;
      if (row.rig) {
        const rigTrim = row.rig.trim();
        const rigUpper = rigTrim.toUpperCase();
        if (rigUpper !== "UNKNOWN" && rigUpper !== "N/A" && rigTrim !== "") {
          uniqueRigs.add(rigTrim);
        }
      }
      if (row.well && row.well !== "N/A") uniqueWells.add(row.well);
      if (row.client) clients.add(row.client);
    });

    return {
      totalFootage,
      totalQty,
      rigCount: uniqueRigs.size,
      wellCount: uniqueWells.size,
      clientCount: clients.size
    };
  }, [data]);

  // 2. Calculate unique rigs operated grouped by full tool name with exact matches and footage accumulation
  const toolCounts = useMemo(() => {
    const counts: { 
      [fullName: string]: { 
        rigs: Set<string>; 
        footage: number;
        primaryAbb: string; 
        category: string;
        description: string;
        allAbbs: string[];
        sections: {
          [secName: string]: {
            rigs: Set<string>;
            footage: number;
          }
        };
      } 
    } = {};

    // Initialize list of abbreviations for each fullName
    abbreviations.forEach((abb) => {
      let name = abb.fullName;
      const upperAbb = (abb.abbreviation || "").trim().toUpperCase();
      const upperName = (name || "").trim().toUpperCase();

      if (upperAbb.startsWith("ES") || upperAbb.startsWith("EARTHSTAR") || upperName.startsWith("ES") || upperName.startsWith("EARTHSTAR")) {
        name = "EarthStar";
      } else if (upperAbb.startsWith("EMBATT") || upperAbb.startsWith("EM BATT") || upperName === "EM BATTERY" || upperName === "EM BATT") {
        name = "EM Battery";
      } else if (upperName === "IC BM" || upperAbb === "IC BM" || upperAbb === "IC(BM)" || upperAbb === "ICBM" || upperAbb === "IC-BM") {
        name = "IC BM";
      }

      let targetAbb = abb.abbreviation;
      if (name === "ResiStar") {
        targetAbb = "RST";
      } else if (name === "PulseStar") {
        targetAbb = "PST";
      } else if (name === "GeoTap") {
        targetAbb = "GeoTap";
      } else if (name === "EarthStar") {
        targetAbb = "EarthStar";
      } else if (name === "EM Battery") {
        targetAbb = "EM BATT";
      } else if (name === "FlowStar") {
        targetAbb = "FlowStar";
      } else if (name === "IC BM") {
        targetAbb = "IC BM";
      }

      if (!counts[name]) {
        counts[name] = {
          rigs: new Set<string>(),
          footage: 0,
          primaryAbb: targetAbb,
          category: abb.category,
          description: abb.description,
          allAbbs: [abb.abbreviation],
          sections: {}
        };
      } else {
        if (!counts[name].allAbbs.includes(abb.abbreviation)) {
          counts[name].allAbbs.push(abb.abbreviation);
        }
        if (targetAbb === "RST" || targetAbb === "PST" || targetAbb === "GeoTap" || targetAbb === "EarthStar" || targetAbb === "EM BATT" || targetAbb === "FlowStar" || targetAbb === "IC BM") {
          counts[name].primaryAbb = targetAbb;
        }
      }
    });

    data.forEach((row) => {
      const config = row.configuration || "";
      const comments = row.comments || "";
      const normalizedConfig = config
        .toUpperCase()
        .replace(/EM[\s\-_]*BATT(ERY)?(\d+)?/gi, (match, p1, p2) => `EMBATT${p2 || ''}`)
        .replace(/IC[\s\-_]*\(?BM\)?\b|ICBM\b/gi, " ICBM ");
      const rowFootage = row.footage || 0;
      const rowQty = row.qty || 1;
      const rig = row.rig ? row.rig.trim().toUpperCase() : "UNKNOWN";
      
      const toolsProcessedInRow = new Set<string>();

      const addToolCount = (fullName: string, isESToken = false, isEMBattToken = false, matchObj?: Abbreviation, isICBMToken = false) => {
        if (!fullName) return;
        if (!counts[fullName]) {
          counts[fullName] = {
            rigs: new Set<string>(),
            footage: 0,
            primaryAbb: isESToken ? "EarthStar" : isEMBattToken ? "EM BATT" : isICBMToken ? "IC BM" : fullName,
            category: isESToken ? "Telemetry / MWD" : isEMBattToken ? "Power" : isICBMToken ? "Directional Drilling" : (matchObj ? matchObj.category : "Tools"),
            description: isESToken 
              ? "EarthStar electromagnetic telemetry sub system." 
              : isEMBattToken 
              ? "Electromagnetic downhole lithium battery pack supplying high-capacity power." 
              : isICBMToken
              ? "iCruise Bus Master (IC BM) interface and telemetry tool module."
              : (matchObj ? matchObj.description : ""),
            allAbbs: matchObj ? [matchObj.abbreviation] : [fullName],
            sections: {}
          };
        }

        counts[fullName].rigs.add(rig);
        counts[fullName].footage += rowFootage * rowQty;

        const rawSec = row.section || "Other";
        const secNormalized = getNormalizedSection(rawSec);

        if (!counts[fullName].sections[secNormalized]) {
          counts[fullName].sections[secNormalized] = { rigs: new Set<string>(), footage: 0 };
        }
        counts[fullName].sections[secNormalized].rigs.add(rig);
        counts[fullName].sections[secNormalized].footage += rowFootage * rowQty;
        toolsProcessedInRow.add(fullName);
      };

      // Split the configuration string into uppercase tokens by removing brackets, parentheses, hyphens, and slashes
      const tokens = normalizedConfig
        .replace(/[()\[\]\-\/\\,;]/g, " ")
        .split(/\s+/)
        .filter(Boolean);

      tokens.forEach((token) => {
        const isESToken = token.startsWith("ES") || token.startsWith("EXRX") || token.startsWith("EARTHSTAR");
        const isEMBattToken = token.startsWith("EMBATT") || token.startsWith("EM_BATT");
        const isICBMToken = token === "ICBM";

        // Find abbreviation that matches this token exactly or matches fullName
        let match = abbreviations.find(
          (abb) => abb.abbreviation.toUpperCase() === token || abb.fullName.toUpperCase().replace(/[\s\-_]/g, "") === token.replace(/[\s\-_]/g, "")
        );

        let fullName = match ? match.fullName : "";
        if (isESToken) {
          fullName = "EarthStar";
        } else if (isEMBattToken) {
          fullName = "EM Battery";
        } else if (isICBMToken) {
          fullName = "IC BM";
        }

        if (fullName) {
          addToolCount(fullName, isESToken, isEMBattToken, match, isICBMToken);
        }
      });

      // Also inspect comments for tool mentions (such as BST Caliper, StrataStar, IC BM, or others)
      if (comments) {
        const commentsUpper = comments.toUpperCase();

        // 0. IC BM check
        if (/IC[\s\-_]*\(?BM\)?\b|ICBM\b/i.test(commentsUpper) && !toolsProcessedInRow.has("IC BM")) {
          const icbmMatch = abbreviations.find(a => a.fullName === "IC BM");
          addToolCount("IC BM", false, false, icbmMatch, true);
        }

        // 1. BST Caliper check
        if (/BST[\s\-_]*CALIPER|CALIPER/i.test(commentsUpper) && !toolsProcessedInRow.has("BST Caliper")) {
          const calMatch = abbreviations.find(a => a.fullName === "BST Caliper");
          addToolCount("BST Caliper", false, false, calMatch);
        }

        // 2. StrataStar check
        if (/STRATASTAR|STRATA[\s\-_]*STAR|\bSS\d*\b/i.test(commentsUpper) && !toolsProcessedInRow.has("StrataStar")) {
          const ssMatch = abbreviations.find(a => a.fullName === "StrataStar");
          addToolCount("StrataStar", false, false, ssMatch);
        }

        // 3. General tool abbreviation check in comments
        const commentsNormalized = commentsUpper
          .replace(/EM[\s\-_]*BATT(ERY)?(\d+)?/gi, (match, p1, p2) => `EMBATT${p2 || ''}`)
          .replace(/IC[\s\-_]*\(?BM\)?\b|ICBM\b/gi, " ICBM ");
        const cTokens = commentsNormalized.replace(/[()\[\]\-\/\\,;]/g, " ").split(/\s+/).filter(Boolean);
        for (const cToken of cTokens) {
          if (cToken.startsWith("ES") || cToken.startsWith("EARTHSTAR")) {
            if (!toolsProcessedInRow.has("EarthStar")) {
              addToolCount("EarthStar", true, false);
            }
          } else if (cToken.startsWith("EMBATT")) {
            if (!toolsProcessedInRow.has("EM Battery")) {
              addToolCount("EM Battery", false, true);
            }
          } else if (cToken === "ICBM") {
            if (!toolsProcessedInRow.has("IC BM")) {
              const icbmMatch = abbreviations.find(a => a.fullName === "IC BM");
              addToolCount("IC BM", false, false, icbmMatch, true);
            }
          } else {
            const match = abbreviations.find(a => a.abbreviation.toUpperCase() === cToken);
            if (match && match.fullName && !toolsProcessedInRow.has(match.fullName)) {
              addToolCount(match.fullName, false, false, match);
            }
          }
        }
      }
    });

    return Object.entries(counts)
      .map(([fullName, info]) => {
        const sectionProps: { [key: string]: number } = {};
        
        // Convert section rigs sets to sizes representing unique rig count
        const sectionsData: { [secName: string]: { count: number; footage: number } } = {};
        Object.entries(info.sections).forEach(([sec, val]) => {
          sectionsData[sec] = {
            count: val.rigs.size,
            footage: val.footage
          };
        });

        Object.entries(sectionsData).forEach(([sec, val]) => {
          const safeSec = getSafeSectionKey(sec);
          sectionProps[`sec_count_${safeSec}`] = val.count;
          sectionProps[`sec_footage_${safeSec}`] = val.footage;
        });

        return {
          abbreviation: info.primaryAbb,
          fullName,
          category: info.category,
          description: info.description,
          count: info.rigs.size,
          footage: info.footage,
          allAbbs: info.allAbbs,
          sections: sectionsData,
          ...sectionProps
        };
      })
      .filter((item) => item.count > 0);
  }, [data, abbreviations]);

  // Unique sections in the active dataset
  const uniqueSections = useMemo(() => {
    const sections = new Set<string>();
    toolCounts.forEach((t) => {
      Object.keys(t.sections).forEach((sec) => {
        sections.add(sec);
      });
    });
    const ordered = Array.from(sections);
    const priority = ["4.75in", "6.75in", "8in", "9.5in"];
    ordered.sort((a, b) => {
      const indexA = priority.indexOf(a);
      const indexB = priority.indexOf(b);
      if (indexA !== -1 && indexB !== -1) return indexA - indexB;
      if (indexA !== -1) return -1;
      if (indexB !== -1) return 1;
      return a.localeCompare(b);
    });
    return ordered;
  }, [toolCounts]);

  // Active categories in the active dataset
  const activeCategories = useMemo(() => {
    const cats = new Set<string>();
    toolCounts.forEach((t) => {
      if (t.category) cats.add(t.category);
    });
    return Array.from(cats).sort();
  }, [toolCounts]);

  // Filter and sort tool statistics dynamically
  const processedToolCounts = useMemo(() => {
    let result = [...toolCounts];

    // Apply selectedTool filter if active (user has selected specific tools in the global filter)
    if (selectedTool && selectedTool !== "All") {
      const selectedTerms = selectedTool
        .split(",")
        .map((s) => s.trim().toLowerCase())
        .filter(Boolean);

      if (selectedTerms.length > 0) {
        result = result.filter((t) => {
          return selectedTerms.some((term) => {
            return (
              t.abbreviation.toLowerCase() === term ||
              t.fullName.toLowerCase() === term ||
              t.allAbbs.some((abb) => abb.toLowerCase() === term)
            );
          });
        });
      }
    }

    // Apply search filter
    if (profileSearch.trim()) {
      const term = profileSearch.toLowerCase();
      result = result.filter(
        (t) =>
          t.fullName.toLowerCase().includes(term) ||
          t.abbreviation.toLowerCase().includes(term) ||
          t.allAbbs.some((a) => a.toLowerCase().includes(term)) ||
          t.category.toLowerCase().includes(term)
      );
    }

    // Apply category filter
    if (profileCategories.length > 0) {
      result = result.filter((t) => profileCategories.includes(t.category));
    }

    // Apply selected sort
    if (profileSortBy === "count") {
      result.sort((a, b) => b.count - a.count || b.footage - a.footage);
    } else {
      result.sort((a, b) => b.footage - a.footage || b.count - a.count);
    }

    return result;
  }, [toolCounts, profileSearch, profileCategories, profileSortBy, selectedTool]);

  // 3. Section Size Distribution
  const sectionCounts = useMemo(() => {
    const counts: { [key: string]: number } = {};
    data.forEach((row) => {
      const rawSec = row.section || row.holeSize || 'Unknown';
      const sec = rawSec !== 'Unknown' ? getSectionFromHoleSize(rawSec) : 'Unknown';
      counts[sec] = (counts[sec] || 0) + 1;
    });
    return Object.entries(counts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
  }, [data]);

  // 4. Client Rig Counts (By unique rigs operated under each client)
  const clientCounts = useMemo(() => {
    const counts: { [key: string]: Set<string> } = {};
    data.forEach((row) => {
      const cli = row.client ? row.client.trim() : 'Other';
      const cliUpper = cli.toUpperCase();
      if (cliUpper !== "UEPL" && cliUpper !== "UNKNOWN CLIENT" && cliUpper !== "UNKNOWN" && cli !== "") {
        const rig = row.rig ? row.rig.trim() : '';
        const rigUpper = rig.toUpperCase();
        if (rigUpper !== "UNKNOWN" && rigUpper !== "N/A" && rig !== "") {
          if (!counts[cli]) {
            counts[cli] = new Set<string>();
          }
          counts[cli].add(rigUpper);
        }
      }
    });
    return Object.entries(counts)
      .map(([name, rigSet]) => ({ name, value: rigSet.size }))
      .filter((item) => item.value > 0)
      .sort((a, b) => b.value - a.value);
  }, [data]);

  // 5. Rig Share (By qty of configurations requested)
  const rigCounts = useMemo(() => {
    const counts: { [key: string]: number } = {};
    data.forEach((row) => {
      const rig = row.rig ? row.rig.trim() : 'Other';
      const rigUpper = rig.toUpperCase();
      if (rigUpper !== "UNKNOWN" && rigUpper !== "N/A" && rig !== "") {
        counts[rig] = (counts[rig] || 0) + row.qty;
      }
    });
    return Object.entries(counts)
      .map(([name, value]) => ({ name, value }))
      .filter((item) => item.value > 0)
      .sort((a, b) => b.value - a.value);
  }, [data]);

  // 6. Detailed rig-by-rig tool counts breakdown for a selected tool
  const selectedToolRigBreakdown = useMemo(() => {
    if (!selectedToolForRigPanel) return [];

    const counts: { 
      [rigName: string]: { 
        sectionWells: { [size: string]: Set<string> };
        dates: string[];
        dateDetails: { date: string; well: string; qty: number; section: string }[];
      } 
    } = {};
    const matchedTool = toolCounts.find(t => t.fullName === selectedToolForRigPanel);
    const validAbbreviations = matchedTool ? matchedTool.allAbbs.map(a => a.toUpperCase()) : [];
    const panelUpper = selectedToolForRigPanel.toUpperCase();
    const isEarthStar = panelUpper === "EARTHSTAR" || panelUpper === "EARTH STAR" || panelUpper.startsWith("ES");
    const isEMBattery = panelUpper === "EM BATTERY" || panelUpper === "EM BATT" || panelUpper.startsWith("EMBATT");
    const isICBM = panelUpper === "IC BM" || panelUpper === "IC(BM)" || panelUpper === "ICBM" || panelUpper === "IC-BM";
    const isICruise = panelUpper === "ICRUISE" || panelUpper === "IC";

    data.forEach((row, rowIndex) => {
      const config = row.configuration || "";
      const comments = row.comments || "";
      const normalizedConfig = config
        .toUpperCase()
        .replace(/EM[\s\-_]*BATT(ERY)?(\d+)?/gi, (match, p1, p2) => `EMBATT${p2 || ''}`)
        .replace(/IC[\s\-_]*\(?BM\)?\b|ICBM\b/gi, " ICBM ");
      const rowQty = row.qty || 1;
      const rig = row.rig ? row.rig.trim() : "Other";
      const rigUpper = rig.toUpperCase();
      
      if (!rig || rigUpper === "UNKNOWN" || rigUpper === "N/A" || rig === "") {
        return;
      }

      // Tokenize configuration
      const tokens = normalizedConfig
        .replace(/[()\[\]\-\/\\,;]/g, " ")
        .split(/\s+/)
        .filter(Boolean);

      // Count how many tool instances match this selected tool in this configuration row
      let toolInstancesInRow = 0;
      for (const token of tokens) {
        if (isEarthStar) {
          if (token.startsWith("ES") || token.startsWith("EXRX") || token.startsWith("EARTHSTAR")) {
            toolInstancesInRow++;
          }
        } else if (isEMBattery) {
          if (token.startsWith("EMBATT") || token.startsWith("EM_BATT")) {
            toolInstancesInRow++;
          }
        } else if (isICBM) {
          if (token === "ICBM") {
            toolInstancesInRow++;
          }
        } else if (isICruise) {
          if (token === "IC" || token === "ICRUISE") {
            toolInstancesInRow++;
          }
        } else {
          if (validAbbreviations.includes(token) || panelUpper === token) {
            toolInstancesInRow++;
          }
        }
      }

      // If tool instance was not found in configuration tokens, check comments
      if (comments && toolInstancesInRow === 0) {
        const commentsUpper = comments.toUpperCase();
        const isBSTCaliper = panelUpper === "BST CALIPER" || panelUpper === "CALIPER" || panelUpper.includes("CALIPER");
        const isStrataStar = panelUpper === "STRATASTAR" || panelUpper === "STRATA STAR" || panelUpper === "SS";

        if (isBSTCaliper) {
          if (/BST[\s\-_]*CALIPER|CALIPER/i.test(commentsUpper)) {
            toolInstancesInRow = 1;
          }
        } else if (isStrataStar) {
          if (/STRATASTAR|STRATA[\s\-_]*STAR|\bSS\d*\b/i.test(commentsUpper)) {
            toolInstancesInRow = 1;
          }
        } else if (isEarthStar) {
          if (/EARTHSTAR|EARTH\s*STAR|\bES\b|\bESTX\b|\bESRX\b/i.test(commentsUpper)) {
            toolInstancesInRow = 1;
          }
        } else if (isEMBattery) {
          if (/EM[\s\-_]*(?:BATTERIES|BATTERY|BATT\d*|BATTPACK|BATT)|EMBATT/i.test(commentsUpper)) {
            toolInstancesInRow = 1;
          }
        } else if (isICBM) {
          if (/IC[\s\-_]*\(?BM\)?\b|ICBM\b/i.test(commentsUpper)) {
            toolInstancesInRow = 1;
          }
        } else if (isICruise) {
          const commentsWithoutICBM = commentsUpper.replace(/IC[\s\-_]*\(?BM\)?\b|ICBM\b/gi, " ");
          if (/ICRUISE|\bIC\b/i.test(commentsWithoutICBM)) {
            toolInstancesInRow = 1;
          }
        } else {
          const cTokens = commentsUpper.replace(/[()\[\]\-\/\\,;]/g, " ").split(/\s+/).filter(Boolean);
          for (const cToken of cTokens) {
            if (validAbbreviations.includes(cToken) || panelUpper === cToken) {
              toolInstancesInRow = 1;
              break;
            }
          }
        }
      }

      if (toolInstancesInRow > 0) {
        const rawSec = row.section || "Other";
        const secNormalized = getNormalizedSection(rawSec);
        if (!counts[rig]) {
          counts[rig] = { sectionWells: {}, dates: [], dateDetails: [] };
        }
        if (!counts[rig].sectionWells[secNormalized]) {
          counts[rig].sectionWells[secNormalized] = new Set<string>();
        }

        const wellClean = (row.well || "").trim();
        const isWellValid = wellClean && wellClean.toUpperCase() !== "N/A" && wellClean.toUpperCase() !== "UNKNOWN";
        const wellIdentifier = isWellValid ? wellClean : `_unspecified_row_${rowIndex}`;
        counts[rig].sectionWells[secNormalized].add(wellIdentifier);

        const addedQty = toolInstancesInRow * rowQty;
        const rowDate = (row.date || row.month || "").trim();
        if (rowDate && !counts[rig].dates.includes(rowDate)) {
          counts[rig].dates.push(rowDate);
        }

        counts[rig].dateDetails.push({
          date: rowDate || "Unspecified",
          well: (row.well || "").trim(),
          qty: addedQty,
          section: secNormalized
        });
      }
    });

    return Object.entries(counts)
      .map(([rigName, info]) => {
        const sizes: { [size: string]: number } = {};
        let totalWells = 0;
        for (const [sec, wellSet] of Object.entries(info.sectionWells)) {
          const wCount = wellSet.size;
          sizes[sec] = wCount;
          totalWells += wCount;
        }

        return {
          rigName,
          count: totalWells,
          sizes,
          dates: info.dates,
          dateDetails: info.dateDetails
        };
      })
      .sort((a, b) => b.count - a.count);
  }, [selectedToolForRigPanel, data, toolCounts]);

  if (view === "rig-breakdown") {
    return (
      <div className="space-y-6" id="stats-dashboard-rig-breakdown">
        {/* A beautiful control panel to select LWD/MWD tool */}
        <BorderGlow
          borderRadius={16}
          backgroundColor={isLightMode ? "#FFFFFF" : "#0F1420"}
          glowColor={isLightMode ? "220 80 50" : "210 100 70"}
          glowIntensity={isLightMode ? 0.6 : 1.2}
          edgeSensitivity={25}
          className="w-full"
        >
          <div className="p-5 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-emerald-500/10 text-emerald-500 flex items-center justify-center font-bold text-lg shrink-0">
                <Layers className="w-5 h-5 text-emerald-500" />
              </div>
              <div>
                <h3 className={`text-sm font-semibold ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"} font-sans`}>
                  Rig Breakdown Selector
                </h3>
                <p className="text-[11px] text-[#9AA0A6] mt-0.5 font-sans">
                  Analyze the rig-by-rig deployment and demand distribution for any active tool.
                </p>
              </div>
            </div>

            {/* Tool Selector dropdown inside Rig Breakdown tab */}
            <div className="relative w-full sm:w-72">
              <label className={`block text-[10px] font-semibold uppercase tracking-wider mb-1 ${isLightMode ? "text-slate-500" : "text-slate-400"}`}>
                Selected Tool
              </label>
              <select
                value={selectedToolForRigPanel || ""}
                onChange={(e) => {
                  const val = e.target.value;
                  setSelectedToolForRigPanel(val || null);
                }}
                className={`w-full h-10 px-3 text-xs border rounded-xl focus:outline-none font-medium transition-all cursor-pointer ${
                  isLightMode
                    ? "bg-white border-slate-200 text-slate-800 focus:border-emerald-500 hover:bg-slate-50"
                    : "bg-[#1A1D21] border-white/10 text-[#E0E0E0] focus:border-emerald-500 hover:bg-[#1f2329]"
                }`}
              >
                <option value="">-- Choose LWD/MWD Tool --</option>
                {toolCounts.map((tc) => (
                  <option key={tc.fullName} value={tc.fullName}>
                    {tc.fullName} ({tc.abbreviation})
                  </option>
                ))}
              </select>
            </div>
          </div>
        </BorderGlow>

        {/* Render the actual Rig breakdown if a tool is selected */}
        {selectedToolForRigPanel ? (
          <motion.div
            id="tool-rig-breakdown-panel"
            initial={{ opacity: 0, y: 15 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -15 }}
            transition={{ duration: 0.3 }}
          >
            <BorderGlow
              borderRadius={16}
              backgroundColor={isLightMode ? "#FFFFFF" : "#0F1420"}
              glowColor="16 185 129"
              glowIntensity={isLightMode ? 0.7 : 1.3}
              edgeSensitivity={25}
              className="w-full"
            >
              <div className="p-6 flex flex-col justify-between w-full h-full">
                <div>
                  <div className="flex items-center justify-between border-b pb-4 mb-5 border-dashed border-emerald-500/20">
                    <div className="flex items-center gap-2.5">
                      <span className="w-2.5 h-6 bg-emerald-500 rounded-full shrink-0" />
                      <div>
                        <h3 className={`text-sm font-semibold ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"} font-sans`}>
                          Rig Usage Breakdown: {selectedToolForRigPanel}
                        </h3>
                        <p className="text-[10px] text-[#9AA0A6] mt-0.5 font-sans">
                          Dynamic allocation of requested tool quantities across unique operating rigs.
                        </p>
                        <div className="flex flex-wrap gap-x-3.5 gap-y-1 mt-2.5">
                          {["4.75in", "6.75in", "8in", "9.5in"].map((size) => (
                            <div key={size} className="flex items-center gap-1.5 text-[9px] font-semibold">
                              <span className="w-2 h-2 rounded-full shrink-0" style={{ backgroundColor: SECTION_COLORS[size] }} />
                              <span className={isLightMode ? "text-slate-600 font-sans" : "text-slate-400 font-sans"}>{size}</span>
                            </div>
                          ))}
                          <div className="flex items-center gap-1.5 text-[9px] font-semibold">
                            <span className="w-2 h-2 rounded-full shrink-0 bg-cyan-500" />
                            <span className={isLightMode ? "text-slate-600 font-sans" : "text-slate-400 font-sans"}>Other</span>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className={`${isLightMode ? "bg-slate-50 border-slate-200 shadow-xs" : "bg-[#121620] border-slate-700/60 shadow-inner"} p-4 sm:p-5 rounded-2xl border`}>
                    {/* Progress List */}
                    <div className="space-y-4 font-sans max-h-[420px] overflow-y-auto pr-2 custom-scrollbar">
                      {selectedToolRigBreakdown.map((item) => {
                        const maxVal = Math.max(...selectedToolRigBreakdown.map(x => x.count), 1);
                        const fillPercent = Math.round((item.count / maxVal) * 100);
                        const totalQty = selectedToolRigBreakdown.reduce((acc, curr) => acc + curr.count, 0) || 1;
                        const sharePercent = Math.round((item.count / totalQty) * 100);

                        // Sort sizes in predictable order
                        const sortedSizeEntries = Object.entries(item.sizes || {}).sort((a, b) => {
                          const priority = ["4.75in", "6.75in", "8in", "9.5in"];
                          const indexA = priority.indexOf(a[0]);
                          const indexB = priority.indexOf(b[0]);
                          if (indexA !== -1 && indexB !== -1) return indexA - indexB;
                          if (indexA !== -1) return -1;
                          if (indexB !== -1) return 1;
                          return a[0].localeCompare(b[0]);
                        });

                        const isExpanded = !!expandedRigDates[item.rigName];

                        return (
                          <div key={item.rigName} className="space-y-1.5">
                            <div className="flex items-center gap-3">
                              <div className={`w-28 text-xs font-semibold ${isLightMode ? "text-slate-800" : "text-slate-200"} truncate font-sans shrink-0`}>
                                <span className="truncate block" title={item.rigName}>{item.rigName}</span>
                              </div>

                              <div className={`flex-1 h-2.5 ${isLightMode ? "bg-slate-200/80 border-slate-300/60" : "bg-[#0B0D14] border-slate-700/60"} rounded-full overflow-hidden relative border`}>
                                <motion.div
                                  initial={{ width: 0 }}
                                  animate={{ width: `${fillPercent}%` }}
                                  transition={{ duration: 0.6, ease: "easeOut" }}
                                  className="h-full flex rounded-full overflow-hidden"
                                >
                                  {sortedSizeEntries.map(([sizeName, wellCount], idx) => {
                                    const segmentWidthPercent = (wellCount / item.count) * 100;
                                    const segmentColor = getSectionColor(sizeName, idx, isLightMode);
                                    return (
                                      <div
                                        key={sizeName}
                                        style={{ 
                                          width: `${segmentWidthPercent}%`, 
                                          backgroundColor: segmentColor 
                                        }}
                                        className="h-full hover:brightness-110 transition-all duration-150"
                                        title={`${sizeName}: ${wellCount} well${wellCount > 1 ? "s" : ""}`}
                                      />
                                    );
                                  })}
                                </motion.div>
                              </div>

                              <div className="shrink-0 flex items-center justify-end gap-2 font-mono text-xs">
                                <button
                                  onClick={() => setExpandedRigDates(prev => ({ ...prev, [item.rigName]: !prev[item.rigName] }))}
                                  className={`px-2 py-0.5 rounded text-[10px] font-sans font-semibold flex items-center gap-1 transition-all cursor-pointer ${
                                    isExpanded
                                      ? (isLightMode ? "bg-emerald-50 text-emerald-700 border border-emerald-300 shadow-2xs" : "bg-emerald-500/25 text-emerald-300 border border-emerald-500/50 shadow-2xs")
                                      : (isLightMode ? "bg-white text-slate-700 hover:bg-slate-100 hover:text-slate-900 border border-slate-200 shadow-2xs" : "bg-[#1F2634] text-slate-200 hover:bg-[#283244] hover:text-white border border-slate-700/80 shadow-2xs")
                                  }`}
                                  title="Toggle dates breakdown for this rig"
                                >
                                  <Calendar className={`w-2.5 h-2.5 ${isLightMode ? "text-emerald-600" : "text-emerald-400"} shrink-0`} />
                                  <span>Show Dates</span>
                                  <ChevronDown className={`w-2.5 h-2.5 transition-transform duration-200 ${isExpanded ? "rotate-180" : ""}`} />
                                </button>
                              </div>
                            </div>

                            {/* Expandable Dates per Size */}
                            {isExpanded && (
                              <motion.div
                                initial={{ opacity: 0, height: 0 }}
                                animate={{ opacity: 1, height: "auto" }}
                                exit={{ opacity: 0, height: 0 }}
                                className={`ml-28 pl-3.5 py-2.5 pr-3 border-l-2 border-emerald-500 rounded-r-xl space-y-2 text-xs font-sans ${
                                  isLightMode
                                    ? "bg-white border-r border-t border-b border-slate-200"
                                    : "bg-[#0B0E16] border-r border-t border-b border-slate-800/80"
                                }`}
                              >
                                <div className={`flex items-center justify-between border-b ${isLightMode ? "border-slate-100" : "border-white/5"} pb-1`}>
                                  <span className={`text-[10px] font-bold uppercase tracking-wider ${isLightMode ? "text-emerald-600" : "text-emerald-400"} font-sans flex items-center gap-1.5`}>
                                    <Calendar className={`w-3.5 h-3.5 ${isLightMode ? "text-emerald-600" : "text-emerald-400"}`} />
                                    Mobilization Dates per Size:
                                  </span>
                                  <span className={`text-[10px] ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"} font-mono`}>
                                    {item.dates.length} active date{item.dates.length > 1 ? "s" : ""}
                                  </span>
                                </div>

                                {(() => {
                                  const datesBySection: { [section: string]: { dates: string[]; wells: string[] } } = {};
                                  item.dateDetails.forEach(dd => {
                                    const sec = dd.section || "Other";
                                    if (!datesBySection[sec]) datesBySection[sec] = { dates: [], wells: [] };
                                    if (dd.date && !datesBySection[sec].dates.includes(dd.date)) {
                                      datesBySection[sec].dates.push(dd.date);
                                    }
                                    if (dd.well && !datesBySection[sec].wells.includes(dd.well)) {
                                      datesBySection[sec].wells.push(dd.well);
                                    }
                                  });

                                  const sectionKeys = Object.keys(datesBySection);
                                  if (sectionKeys.length === 0) {
                                    return <p className={`text-[10px] ${isLightMode ? "text-slate-400" : "text-[#9AA0A6]"} italic py-1`}>No dates recorded for this rig</p>;
                                  }

                                  return (
                                    <div className="space-y-1.5 pt-0.5">
                                      {sectionKeys.map((secName, sIdx) => {
                                        const secInfo = datesBySection[secName];
                                        const color = getSectionColor(secName, sIdx, isLightMode);
                                        return (
                                          <div key={secName} className="flex flex-wrap items-center gap-2 text-[11px]">
                                            <span className={`inline-flex items-center gap-1.5 font-bold text-[10px] px-2 py-0.5 rounded ${
                                              isLightMode 
                                                ? "bg-slate-100 text-slate-700 border border-slate-200" 
                                                : "bg-[#1F2634] text-slate-100 border border-slate-700/80"
                                            } shrink-0`}>
                                              <span className="w-2 h-2 rounded-full" style={{ backgroundColor: color }} />
                                              {secName}
                                            </span>
                                            <div className="flex flex-wrap items-center gap-1.5">
                                              {secInfo.dates.length > 0 ? (
                                                secInfo.dates.map((d, dIdx) => (
                                                  <span
                                                    key={dIdx}
                                                    className={`px-2 py-0.5 rounded text-[10px] font-mono font-semibold border transition-all ${
                                                      isLightMode
                                                        ? "bg-emerald-50 text-emerald-700 border-emerald-200"
                                                        : "bg-emerald-950/80 text-emerald-300 border-emerald-800/80"
                                                    }`}
                                                  >
                                                    {d}
                                                  </span>
                                                ))
                                              ) : (
                                                <span className={`text-[10px] ${isLightMode ? "text-slate-400" : "text-[#9AA0A6]"} italic`}>No dates</span>
                                              )}
                                            </div>
                                            {secInfo.wells.length > 0 && (
                                              <span className={`text-[10px] ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"} font-mono`}>
                                                Well: {secInfo.wells.join(", ")}
                                              </span>
                                            )}
                                          </div>
                                        );
                                      })}
                                    </div>
                                  );
                                })()}
                              </motion.div>
                            )}
                          </div>
                        );
                      })}
                      {selectedToolRigBreakdown.length === 0 && (
                        <div className={`text-center py-12 text-xs ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"} font-sans`}>
                          No active rig demands found for this tool.
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            </BorderGlow>
          </motion.div>
        ) : (
          <BorderGlow
            borderRadius={12}
            backgroundColor={isLightMode ? "#FFFFFF" : "#0F1420"}
            glowColor="220 80 50"
            glowIntensity={isLightMode ? 0.4 : 0.8}
            edgeSensitivity={20}
            className="w-full"
          >
            <div className="p-6 text-center flex flex-col items-center justify-center">
              <EmptyState
                title="No LWD/MWD Tool Selected"
                description="Select a tool from the dropdown above or click any tool bar in the chart to inspect its deployment distribution."
              />
            </div>
          </BorderGlow>
        )}
      </div>
    );
  }

  return (
    <div className="space-y-6" id="stats-dashboard">
      {/* Operational KPI Card */}
      {(view === "all" || view === "profile") && (
        <BorderGlow
          borderRadius={12}
          backgroundColor={isLightMode ? "#FFFFFF" : "#0F1420"}
          glowColor={isLightMode ? "220 80 50" : "210 100 70"}
          glowIntensity={isLightMode ? 0.6 : 1.2}
          edgeSensitivity={25}
          className="w-full"
        >
          <div className="p-6 flex flex-col justify-between h-full" id="kpi-total-rigs">
            <div>
              <p className={`text-xs font-semibold uppercase tracking-wider ${isLightMode ? "text-slate-600" : "text-[#9AA0A6]"}`}>TOTAL RIGS</p>
              <div className="flex items-baseline gap-2 mt-2">
                <p className={`text-[44px] font-bold tracking-tight leading-none font-sans ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"}`}>
                  {stats.rigCount}
                </p>
                <span className={`text-xs font-medium font-sans ${isLightMode ? "text-slate-600" : "text-[#9AA0A6]"}`}>rigs</span>
              </div>
            </div>
          </div>
        </BorderGlow>
      )}

      {/* Grid: Tool Deployment Counts vs Client Share */}
      {(view === "all" || view === "profile") && (
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
          {/* Tool deployment count chart */}
          <BorderGlow
            borderRadius={16}
            backgroundColor={isLightMode ? "#FFFFFF" : "#0F1420"}
            glowColor={isLightMode ? "220 80 50" : "210 100 70"}
            glowIntensity={isLightMode ? 0.6 : 1.2}
            edgeSensitivity={25}
            className="lg:col-span-12 flex flex-col h-full"
          >
          <div className="p-5 flex flex-col justify-between min-h-[480px] w-full h-full">
            <div>
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4">
                <div>
                  <h3 className={`text-sm font-semibold ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"} font-sans`}>Tools Forecast</h3>
                </div>
                <span className={`text-[10px] ${isLightMode ? "bg-slate-100 text-slate-700 border-slate-200" : "bg-[#1A1D21] text-[#E0E0E0] border-white/5"} border px-2.5 py-1 rounded-lg font-mono font-bold self-start sm:self-auto`}>
                  {processedToolCounts.length} active classes
                </span>
              </div>


            </div>

            {/* Recharts Bar Chart representing active tool counts or footage */}
            <div className="w-full h-[420px] flex-1 mt-4" id="tool-bar-chart-container">
              {processedToolCounts.length === 0 ? (
                <div className="h-full w-full flex flex-col items-center justify-center gap-2 text-[#9AA0A6] text-xs font-sans">
                  <Info className="w-6 h-6 text-[#9AA0A6]/40" />
                  No matching tools or abbreviations found inside the active filters.
                </div>
              ) : (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart
                    data={processedToolCounts}
                    margin={{ top: 15, right: 10, left: -20, bottom: 50 }}
                  >
                    <defs>
                      <linearGradient id="barGradient" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#3B82F6" stopOpacity={0.9} />
                        <stop offset="100%" stopColor="#1D4ED8" stopOpacity={0.25} />
                      </linearGradient>
                      <linearGradient id="selectedBarGradient" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#10B981" stopOpacity={0.9} />
                        <stop offset="100%" stopColor="#047857" stopOpacity={0.25} />
                      </linearGradient>
                      <linearGradient id="barGradientLight" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#3B82F6" stopOpacity={0.9} />
                        <stop offset="100%" stopColor="#60A5FA" stopOpacity={0.4} />
                      </linearGradient>
                      <linearGradient id="selectedBarGradientLight" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#10B981" stopOpacity={0.9} />
                        <stop offset="100%" stopColor="#34D399" stopOpacity={0.4} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={isLightMode ? "rgba(0, 0, 0, 0.05)" : "rgba(255, 255, 255, 0.05)"} vertical={false} />
                    <XAxis
                      dataKey="abbreviation"
                      stroke="#9AA0A6"
                      fontSize={9}
                      fontFamily="Inter, sans-serif"
                      tickLine={false}
                      axisLine={{ stroke: isLightMode ? "rgba(0, 0, 0, 0.1)" : "rgba(255, 255, 255, 0.1)" }}
                      interval={0}
                      tick={<CustomXAxisTick isLightMode={isLightMode} />}
                      height={85}
                    />
                    <YAxis
                      stroke="#9AA0A6"
                      fontSize={10}
                      fontFamily="JetBrains Mono, monospace"
                      tickLine={false}
                      axisLine={false}
                      tickFormatter={(value) => {
                        if (value >= 1000) {
                          return `${(value / 1000).toFixed(0)}k`;
                        }
                        return value.toString();
                      }}
                      allowDecimals={false}
                    />
                    <Tooltip 
                      content={<CustomTooltip isLightMode={isLightMode} profileSortBy={profileSortBy} />} 
                      cursor={{ fill: isLightMode ? "rgba(0, 0, 0, 0.02)" : "rgba(255, 255, 255, 0.02)" }} 
                    />
                    {uniqueSections.map((sec, idx) => {
                      const dataKey = `sec_${profileSortBy}_${getSafeSectionKey(sec)}`;
                      const color = getSectionColor(sec, idx, isLightMode);
                      return (
                        <Bar
                          key={sec}
                          dataKey={dataKey}
                          stackId="a"
                          fill={color}
                          stroke={isLightMode ? "#FFFFFF" : "#0F1420"}
                          strokeWidth={isLightMode ? 1 : 0}
                          maxBarSize={32}
                          className="hover:opacity-85 transition-opacity duration-200 cursor-pointer"
                          onClick={(data) => {
                            const item = data?.payload || data;
                            if (item && item.fullName) {
                              setSelectedToolForRigPanel(item.fullName);
                              if (view === "all") {
                                setTimeout(() => {
                                  document.getElementById("tool-rig-breakdown-panel")?.scrollIntoView({ behavior: "smooth" });
                                }, 120);
                              }
                            }
                          }}
                        />
                      );
                    })}
                  </BarChart>
                </ResponsiveContainer>
              )}
            </div>

            {/* Interactive Legend for Tool Sizes */}
            <div className={`mt-4 flex flex-wrap items-center justify-center gap-x-4 gap-y-2 pt-3 border-t ${isLightMode ? "border-slate-100" : "border-white/5"}`}>
              {uniqueSections.map((sec, idx) => {
                const color = getSectionColor(sec, idx, isLightMode);
                return (
                  <div key={sec} className="flex items-center gap-1.5 font-sans">
                    <span className="w-2.5 h-2.5 rounded-full shrink-0 shadow-sm" style={{ backgroundColor: color }} />
                    <span className={`text-[11px] ${isLightMode ? "text-slate-600" : "text-[#9AA0A6]"} font-medium`}>
                      {sec === "Other" ? "Other Sections" : sec}
                    </span>
                  </div>
                );
              })}
            </div>
          </div>
        </BorderGlow>
        </div>
      )}

      {/* Tool Rig Usage Breakdown Panel */}
      {view === "all" && selectedToolForRigPanel && (
        <motion.div
          id="tool-rig-breakdown-panel"
          initial={{ opacity: 0, y: 15 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -15 }}
          transition={{ duration: 0.3 }}
        >
          <BorderGlow
            borderRadius={16}
            backgroundColor={isLightMode ? "#FFFFFF" : "#0F1420"}
            glowColor="16 185 129"
            glowIntensity={isLightMode ? 0.7 : 1.3}
            edgeSensitivity={25}
            className="w-full"
          >
            <div className="p-6 flex flex-col justify-between w-full h-full">
              <div>
                <div className="flex items-center justify-between border-b pb-4 mb-5 border-dashed border-emerald-500/20">
                  <div className="flex items-center gap-2.5">
                    <span className="w-2.5 h-6 bg-emerald-500 rounded-full shrink-0" />
                    <div>
                      <h3 className={`text-sm font-semibold ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"} font-sans`}>
                        Rig Usage Breakdown: {selectedToolForRigPanel}
                      </h3>
                      <p className="text-[10px] text-[#9AA0A6] mt-0.5 font-sans">
                        Dynamic allocation of requested tool quantities across unique operating rigs.
                      </p>
                      <div className="flex flex-wrap gap-x-3.5 gap-y-1 mt-2.5">
                        {["4.75in", "6.75in", "8in", "9.5in"].map((size) => (
                          <div key={size} className="flex items-center gap-1.5 text-[9px] font-semibold">
                            <span className="w-2 h-2 rounded-full shrink-0" style={{ backgroundColor: SECTION_COLORS[size] }} />
                            <span className={isLightMode ? "text-slate-600 font-sans" : "text-slate-400 font-sans"}>{size}</span>
                          </div>
                        ))}
                        <div className="flex items-center gap-1.5 text-[9px] font-semibold">
                          <span className="w-2 h-2 rounded-full shrink-0 bg-cyan-500" />
                          <span className={isLightMode ? "text-slate-600 font-sans" : "text-slate-400 font-sans"}>Other</span>
                        </div>
                      </div>
                    </div>
                  </div>
                  <button
                    onClick={() => setSelectedToolForRigPanel(null)}
                    className={`p-1.5 rounded-lg transition-colors border ${
                      isLightMode 
                        ? "bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-500" 
                        : "bg-[#1A1D21] hover:bg-[#25282C] border-white/5 text-[#9AA0A6] hover:text-[#F5F7FA]"
                    }`}
                    title="Close Panel"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className={`${isLightMode ? "bg-slate-50 border-slate-200 shadow-xs" : "bg-[#121620] border-slate-700/60 shadow-inner"} p-4 sm:p-5 rounded-2xl border`}>
                    {/* Progress List */}
                    <div className="space-y-4 font-sans max-h-[380px] overflow-y-auto pr-2 custom-scrollbar">
                      {selectedToolRigBreakdown.map((item) => {
                        const maxVal = Math.max(...selectedToolRigBreakdown.map(x => x.count), 1);
                        const fillPercent = Math.round((item.count / maxVal) * 100);
                        const totalQty = selectedToolRigBreakdown.reduce((acc, curr) => acc + curr.count, 0) || 1;
                        const sharePercent = Math.round((item.count / totalQty) * 100);

                        // Sort sizes in predictable order
                        const sortedSizeEntries = Object.entries(item.sizes || {}).sort((a, b) => {
                          const priority = ["4.75in", "6.75in", "8in", "9.5in"];
                          const indexA = priority.indexOf(a[0]);
                          const indexB = priority.indexOf(b[0]);
                          if (indexA !== -1 && indexB !== -1) return indexA - indexB;
                          if (indexA !== -1) return -1;
                          if (indexB !== -1) return 1;
                          return a[0].localeCompare(b[0]);
                        });

                        const isExpanded = !!expandedRigDates[item.rigName];

                        return (
                          <div key={item.rigName} className="space-y-1.5">
                            <div className="flex items-center gap-3">
                              <div className={`w-28 text-xs font-semibold ${isLightMode ? "text-slate-800" : "text-slate-200"} truncate font-sans shrink-0`}>
                                <span className="truncate block" title={item.rigName}>{item.rigName}</span>
                              </div>

                              <div className={`flex-1 h-2.5 ${isLightMode ? "bg-slate-200/80 border-slate-300/60" : "bg-[#0B0D14] border-slate-700/60"} rounded-full overflow-hidden relative border`}>
                                <motion.div
                                  initial={{ width: 0 }}
                                  animate={{ width: `${fillPercent}%` }}
                                  transition={{ duration: 0.6, ease: "easeOut" }}
                                  className="h-full flex rounded-full overflow-hidden"
                                >
                                  {sortedSizeEntries.map(([sizeName, wellCount], idx) => {
                                    const segmentWidthPercent = (wellCount / item.count) * 100;
                                    const segmentColor = getSectionColor(sizeName, idx, isLightMode);
                                    return (
                                      <div
                                        key={sizeName}
                                        style={{ 
                                          width: `${segmentWidthPercent}%`, 
                                          backgroundColor: segmentColor 
                                        }}
                                        className="h-full hover:brightness-110 transition-all duration-150"
                                        title={`${sizeName}: ${wellCount} well${wellCount > 1 ? "s" : ""}`}
                                      />
                                    );
                                  })}
                                </motion.div>
                              </div>

                              <div className="shrink-0 flex items-center justify-end gap-2 font-mono text-xs">
                                <button
                                  onClick={() => setExpandedRigDates(prev => ({ ...prev, [item.rigName]: !prev[item.rigName] }))}
                                  className={`px-2 py-0.5 rounded text-[10px] font-sans font-semibold flex items-center gap-1 transition-all cursor-pointer ${
                                    isExpanded
                                      ? (isLightMode ? "bg-emerald-50 text-emerald-700 border border-emerald-300 shadow-2xs" : "bg-emerald-500/25 text-emerald-300 border border-emerald-500/50 shadow-2xs")
                                      : (isLightMode ? "bg-white text-slate-700 hover:bg-slate-100 hover:text-slate-900 border border-slate-200 shadow-2xs" : "bg-[#1F2634] text-slate-200 hover:bg-[#283244] hover:text-white border border-slate-700/80 shadow-2xs")
                                  }`}
                                  title="Toggle dates breakdown for this rig"
                                >
                                  <Calendar className={`w-2.5 h-2.5 ${isLightMode ? "text-emerald-600" : "text-emerald-400"} shrink-0`} />
                                  <span>Show Dates</span>
                                  <ChevronDown className={`w-2.5 h-2.5 transition-transform duration-200 ${isExpanded ? "rotate-180" : ""}`} />
                                </button>
                              </div>
                            </div>

                            {/* Expandable Dates per Size */}
                            {isExpanded && (
                              <motion.div
                                initial={{ opacity: 0, height: 0 }}
                                animate={{ opacity: 1, height: "auto" }}
                                exit={{ opacity: 0, height: 0 }}
                                className={`ml-28 pl-3.5 py-2.5 pr-3 border-l-2 border-emerald-500 rounded-r-xl space-y-2 text-xs font-sans ${
                                  isLightMode
                                    ? "bg-white border-r border-t border-b border-slate-200"
                                    : "bg-[#0B0E16] border-r border-t border-b border-slate-800/80"
                                }`}
                              >
                                <div className={`flex items-center justify-between border-b ${isLightMode ? "border-slate-100" : "border-white/5"} pb-1`}>
                                  <span className={`text-[10px] font-bold uppercase tracking-wider ${isLightMode ? "text-emerald-600" : "text-emerald-400"} font-sans flex items-center gap-1.5`}>
                                    <Calendar className={`w-3.5 h-3.5 ${isLightMode ? "text-emerald-600" : "text-emerald-400"}`} />
                                    Mobilization Dates per Size:
                                  </span>
                                  <span className={`text-[10px] ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"} font-mono`}>
                                    {item.dates.length} active date{item.dates.length > 1 ? "s" : ""}
                                  </span>
                                </div>

                                {(() => {
                                  const datesBySection: { [section: string]: { dates: string[]; wells: string[] } } = {};
                                  item.dateDetails.forEach(dd => {
                                    const sec = dd.section || "Other";
                                    if (!datesBySection[sec]) datesBySection[sec] = { dates: [], wells: [] };
                                    if (dd.date && !datesBySection[sec].dates.includes(dd.date)) {
                                      datesBySection[sec].dates.push(dd.date);
                                    }
                                    if (dd.well && !datesBySection[sec].wells.includes(dd.well)) {
                                      datesBySection[sec].wells.push(dd.well);
                                    }
                                  });

                                  const sectionKeys = Object.keys(datesBySection);
                                  if (sectionKeys.length === 0) {
                                    return <p className={`text-[10px] ${isLightMode ? "text-slate-400" : "text-[#9AA0A6]"} italic py-1`}>No dates recorded for this rig</p>;
                                  }

                                  return (
                                    <div className="space-y-1.5 pt-0.5">
                                      {sectionKeys.map((secName, sIdx) => {
                                        const secInfo = datesBySection[secName];
                                        const color = getSectionColor(secName, sIdx, isLightMode);
                                        return (
                                          <div key={secName} className="flex flex-wrap items-center gap-2 text-[11px]">
                                            <span className={`inline-flex items-center gap-1.5 font-bold text-[10px] px-2 py-0.5 rounded ${
                                              isLightMode 
                                                ? "bg-slate-100 text-slate-700 border border-slate-200" 
                                                : "bg-[#1F2634] text-slate-100 border border-slate-700/80"
                                            } shrink-0`}>
                                              <span className="w-2 h-2 rounded-full" style={{ backgroundColor: color }} />
                                              {secName}
                                            </span>
                                            <div className="flex flex-wrap items-center gap-1.5">
                                              {secInfo.dates.length > 0 ? (
                                                secInfo.dates.map((d, dIdx) => (
                                                  <span
                                                    key={dIdx}
                                                    className={`px-2 py-0.5 rounded text-[10px] font-mono font-semibold border transition-all ${
                                                      isLightMode
                                                        ? "bg-emerald-50 text-emerald-700 border-emerald-200"
                                                        : "bg-emerald-950/80 text-emerald-300 border-emerald-800/80"
                                                    }`}
                                                  >
                                                    {d}
                                                  </span>
                                                ))
                                              ) : (
                                                <span className={`text-[10px] ${isLightMode ? "text-slate-400" : "text-[#9AA0A6]"} italic`}>No dates</span>
                                              )}
                                            </div>
                                            {secInfo.wells.length > 0 && (
                                              <span className={`text-[10px] ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"} font-mono`}>
                                                Well: {secInfo.wells.join(", ")}
                                              </span>
                                            )}
                                          </div>
                                        );
                                      })}
                                    </div>
                                  );
                                })()}
                              </motion.div>
                            )}
                          </div>
                        );
                      })}
                      {selectedToolRigBreakdown.length === 0 && (
                        <div className={`text-center py-12 text-xs ${isLightMode ? "text-slate-500" : "text-[#9AA0A6]"} font-sans`}>
                          No active rig demands found for this tool.
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Summary Callout Widget */}
                  <div className={`p-4 rounded-xl border flex flex-col justify-center h-full ${
                    isLightMode 
                      ? "bg-slate-50/50 border-slate-100" 
                      : "bg-[#141822] border-white/5"
                  }`}>
                    <div className="flex items-start gap-3">
                      <Info className="w-4 h-4 text-emerald-500 mt-0.5 shrink-0" />
                      <div className="space-y-3 font-sans w-full">
                        <p className={`text-xs leading-relaxed ${isLightMode ? "text-slate-600" : "text-[#9AA0A6]"}`}>
                          The tool <span className={`font-semibold ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"}`}>{selectedToolForRigPanel}</span> is deployed across <span className="text-emerald-500 font-bold">{selectedToolRigBreakdown.length} unique rigs</span> in your current active selection.
                        </p>
                        <div className="pt-2 border-t border-dashed border-white/5 grid grid-cols-3 gap-2">
                          <div>
                            <p className="text-[9px] text-[#9AA0A6] uppercase tracking-wider font-semibold">Total Wells</p>
                            <p className={`text-base font-bold font-mono ${isLightMode ? "text-slate-900" : "text-emerald-400"}`}>
                              {selectedToolRigBreakdown.reduce((acc, curr) => acc + curr.count, 0)}
                            </p>
                          </div>
                          <div>
                            <p className="text-[9px] text-[#9AA0A6] uppercase tracking-wider font-semibold">Max Wells / Rig</p>
                            <p className={`text-base font-bold font-mono ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"}`}>
                              {selectedToolRigBreakdown.length > 0 ? Math.max(...selectedToolRigBreakdown.map(x => x.count)) : 0}
                            </p>
                          </div>
                          <div>
                            <p className="text-[9px] text-[#9AA0A6] uppercase tracking-wider font-semibold">Active Dates</p>
                            <p className={`text-base font-bold font-mono ${isLightMode ? "text-slate-900" : "text-emerald-300"}`}>
                              {Array.from(new Set(selectedToolRigBreakdown.flatMap(x => x.dates))).length}
                            </p>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </BorderGlow>
        </motion.div>
      )}

      {/* Client Share Breakdown */}
      {(view === "all" || view === "allocations") && (
        <div className="grid grid-cols-1 gap-5 mt-2">
          {/* Client share breakdown */}
          <div className="w-full flex flex-col h-full">
            {/* Client share card */}
            <BorderGlow
              borderRadius={16}
              backgroundColor={isLightMode ? "#FFFFFF" : "#0F1420"}
              glowColor={isLightMode ? "220 80 50" : "210 100 70"}
              glowIntensity={isLightMode ? 0.6 : 1.2}
              edgeSensitivity={25}
              className="w-full h-full min-h-[380px]"
            >
              <div className="p-5 w-full h-full flex flex-col justify-between">
                <div>
                  <h3 className={`text-sm font-semibold ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"} font-sans mb-1`}>Client Rig Allocations</h3>
                  <p className="text-[10px] text-[#9AA0A6] mb-5 font-sans">
                    Number of active unique rigs operated by each client.
                  </p>
                  <div className="space-y-4 font-sans">
                    {clientCounts.slice(0, 7).map((cli) => {
                      const totalSet = clientCounts.reduce((acc, curr) => acc + curr.value, 0) || 1;
                      const sharePercent = Math.round((cli.value / totalSet) * 100);

                      return (
                        <div key={cli.name} className="flex items-center gap-3">
                          <div className={`w-24 text-xs font-semibold ${isLightMode ? "text-slate-700" : "text-[#9AA0A6]"} truncate font-sans`}>{cli.name}</div>
                          <div className={`flex-1 h-2.5 ${isLightMode ? "bg-slate-100 border-slate-200" : "bg-[#1A1D21] border-white/5"} rounded-full overflow-hidden relative border`}>
                            <motion.div
                              initial={{ width: 0 }}
                              animate={{ width: `${sharePercent}%` }}
                              transition={{ duration: 0.8, ease: "easeOut" }}
                              className="h-full bg-gradient-to-r from-blue-500 to-indigo-600 rounded-full"
                            />
                          </div>
                          <div className={`w-12 text-right font-mono text-xs ${isLightMode ? "text-slate-900" : "text-[#F5F7FA]"} font-bold shrink-0`}>{cli.value}</div>
                        </div>
                      );
                    })}
                    {clientCounts.length === 0 && (
                      <div className="text-center py-12 text-xs text-[#9AA0A6] font-sans">
                        No client rig allocations found.
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </BorderGlow>
          </div>
        </div>
      )}
    </div>
  );
}

export default React.memo(StatsDashboard);
