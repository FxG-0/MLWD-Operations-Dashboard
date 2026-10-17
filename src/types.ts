export interface RowData {
  id: string;
  client: string;
  rig: string;
  well: string;
  section: string;
  holeSize: string;
  qty: number;
  configuration: string;
  footage: number | null;
  date: string; // e.g., '22-Jun-26'
  month: string; // e.g., 'Jun'
  shopDate: string; // e.g., 'Delivered' or empty
  comments: string;
  validationErrors?: string[]; // Schema validation flags and error messages
}

/**
 * Linking logic between Tool Size (Section) and Hole Size:
 *  - 4.75in Tool Size -> Hole Sizes: 5 7/8in, and all 6in (6", 6.125", 6 1/8", 6 1/2")
 *  - 6.75in Tool Size -> Hole Size: 8.5in (8.5", 8 1/2")
 *  - 8in Tool Size    -> Hole Size: 12.25in (12.25", 12 1/4")
 *  - 9.5in Tool Size  -> Hole Size: 16in and bigger (16", 17.5", 17 1/2", >=15")
 */
function parseNumericSize(str: string): number {
  if (!str) return 0;
  let s = str.replace(/["'”’in\-]/gi, " ").replace(/\s+/g, " ").trim();
  if (!s) return 0;
  if (/^\d+(\.\d+)?$/.test(s)) return parseFloat(s);
  const parts = s.split(" ");
  if (parts.length === 2) {
    const whole = parseFloat(parts[0]);
    const frac = parts[1];
    if (frac.includes("/")) {
      const fp = frac.split("/");
      if (fp.length === 2) {
        const num = parseFloat(fp[0]);
        const den = parseFloat(fp[1]);
        if (!isNaN(whole) && !isNaN(num) && !isNaN(den) && den !== 0) {
          return whole + (num / den);
        }
      }
    }
  } else if (parts.length === 1 && parts[0].includes("/")) {
    const fp = parts[0].split("/");
    if (fp.length === 2) {
      const num = parseFloat(fp[0]);
      const den = parseFloat(fp[1]);
      if (!isNaN(num) && !isNaN(den) && den !== 0) return num / den;
    }
  }
  const match = s.match(/^\d+(\.\d+)?/);
  if (match) return parseFloat(match[0]);
  return 0;
}

export function getSectionFromHoleSize(raw: string): string {
  if (!raw) return '8in';
  const cleaned = raw.trim().toLowerCase();
  if (!cleaned || cleaned === 'n/a') return '8in';

  // Direct check if it's already one of the 4 Tool Sizes
  if (cleaned === '4.75in' || cleaned === '4.75"' || cleaned === '4.75') return '4.75in';
  if (cleaned === '6.75in' || cleaned === '6.75"' || cleaned === '6.75') return '6.75in';
  if (cleaned === '8in' || cleaned === '8.25in' || cleaned === '8"' || cleaned === '8.25"' || cleaned === '8') return '8in';
  if (cleaned === '9.5in' || cleaned === '9.5"' || cleaned === '9.5') return '9.5in';

  // Explicit string pattern matches for Hole Sizes
  // 1. 4.75in tool size for hole sizes: 5 7/8in and all 6in
  if (cleaned.includes('5 7/8') || cleaned.includes('5.875') || cleaned.includes('6.125') || cleaned.includes('6 1/8') || cleaned.includes('6 1/4') || cleaned.includes('6 1/2') || cleaned === '6in' || cleaned === '6"' || cleaned === '6') {
    return '4.75in';
  }
  // 2. 6.75in tool size for hole size: 8.5in
  if (cleaned.includes('8.5') || cleaned.includes('8 1/2')) {
    return '6.75in';
  }
  // 3. 8in tool size for hole size: 12.25in
  if (cleaned.includes('12.25') || cleaned.includes('12 1/4') || cleaned.includes('12.5')) {
    return '8in';
  }
  // 4. 9.5in tool size for hole size: 16in and bigger
  if (cleaned.includes('16') || cleaned.includes('17.5') || cleaned.includes('17 1/2') || cleaned.includes('20')) {
    return '9.5in';
  }

  const val = parseNumericSize(cleaned);
  if (val <= 0) return raw.trim();

  // Numeric range logic based on Hole Size
  if (val <= 7.2) return '4.75in';      // Hole sizes <= 7.2" (5 7/8", 6", 6.125", etc.) -> 4.75in
  if (val <= 10.0) return '6.75in';     // Hole sizes 7.3" - 10.0" (8.5") -> 6.75in
  if (val < 15.0) return '8in';         // Hole sizes 10.1" - 14.9" (12.25") -> 8in
  return '9.5in';                       // Hole sizes >= 15.0" (16" and bigger) -> 9.5in
}

export function getHoleSizeFromSection(raw: string): string {
  if (!raw) return '12.25"';
  const cleaned = raw.trim().toLowerCase();
  if (!cleaned || cleaned === 'n/a') return '12.25"';

  // If already formatted like a Hole Size, preserve it
  if (cleaned.includes('5 7/8') || cleaned.includes('5.875')) return '5 7/8"';
  if (cleaned.includes('6.125') || cleaned.includes('6 1/8')) return '6.125"';
  if (cleaned.includes('6in') || cleaned.endsWith('6"') || cleaned === '6') return '6"';
  if (cleaned.includes('8.5') || cleaned.includes('8 1/2')) return '8.5"';
  if (cleaned.includes('12.25') || cleaned.includes('12 1/4')) return '12.25"';
  if (cleaned.includes('16')) return '16"';
  if (cleaned.includes('17.5') || cleaned.includes('17 1/2')) return '17.5"';

  const toolSize = getSectionFromHoleSize(raw);
  if (toolSize === '4.75in') return '6.125"';
  if (toolSize === '6.75in') return '8.5"';
  if (toolSize === '8in') return '12.25"';
  if (toolSize === '9.5in') return '16"';

  return '12.25"';
}

export const VALID_TOOL_CLASSES = new Set([
  "IC", "ICRUISE", "IC BM", "IC(BM)", "ICBM", "IC-BM", "BST", "BASE", "BASESTAR", "RST", "RESISTAR", "RS", "STS",
  "LST", "LITHO", "BATT", "BATTERY", "BATT1", "BATT2", "ISTARBATT", "ISTARBATT1",
  "ISTARBATT2", "EM BATT", "EMBATT", "EMBATTERY", "EM BATTERY", "EM BATTERIES", "EMBATTERIES", "EM-BATTERIES", "EM BATT1", "EM BATT2", "EM BATT 1", "EM BATT 2", "EM-BATT", "EM-BATTERY", "EM BATT PACK", "EMBATTPACK", "ILS", "HAL50ILS", "XBAT", "AFR", "LC", "LCS", "UC", "UCS", "USC",
  "GEOTAP", "GT", "ALD", "GWD", "DGR", "PWD", "DUMMYFLEX", "DF", "CALIPER", "BST CALIPER", "BSTCALIPER", "ADR",
  "PULSESTAR", "PST", "PULSAR", "FST", "FLOWSTAR", "FLOW STAR", "FLOW-STAR", "FLOW", "GYROSTAR", "EARTHSTAR", "ES", "ES TX", "ES RX", "ESTX", "ESTX1", "ESTX2", "ESRX", "ESRX1", "ESRX2", "ESRCVR", "ESRCVR1", "ESRCVR2", "FDC", "HCIM", "PCDC",
  "PCGK", "P4M", "DHS"
]);

export const validateRow = (row: RowData): string[] => {
  const errors: string[] = [];

  // 1. Client Validation
  const clientClean = (row.client || "").trim();
  if (!clientClean || clientClean.toUpperCase() === "N/A" || clientClean.toUpperCase() === "UNKNOWN") {
    errors.push("Missing or invalid Operator/Client name");
  }

  // 2. Rig Validation
  const rigClean = (row.rig || "").trim();
  if (!rigClean || rigClean.toUpperCase() === "N/A" || rigClean.toUpperCase() === "UNKNOWN") {
    errors.push("Missing or invalid Rig designation");
  }

  // 3. Well Validation
  const wellClean = (row.well || "").trim();
  if (!wellClean || wellClean.toUpperCase() === "N/A" || wellClean.toUpperCase() === "UNKNOWN") {
    errors.push("Missing or invalid Well name / Wellbore ID");
  }

  // 4. Section size Validation
  const sectionClean = (row.section || "").trim();
  if (!sectionClean || sectionClean.toUpperCase() === "N/A") {
    errors.push("Missing or unspecified Tool Size");
  } else {
    // Check if it contains a digit (e.g. 8.5", 12 1/4, 6 in)
    if (!/[0-9]/.test(sectionClean)) {
      errors.push(`Tool Size format warning: "${sectionClean}" is not a recognized dimension`);
    }
  }

  // 5. Quantity Validation
  if (row.qty === undefined || row.qty === null || isNaN(row.qty) || row.qty < 1) {
    errors.push(`Invalid tool quantity: ${row.qty} (must be at least 1)`);
  }

  // 6. Tool Configuration and Tool Classes Validation
  const configClean = (row.configuration || "").trim();
  if (!configClean || configClean.toUpperCase() === "NO CONFIG SPECIFIED" || configClean.toUpperCase() === "N/A") {
    errors.push("Missing or unspecified BHA Tool Configuration string");
  } else {
    // Extract tokens from configuration (split on punctuation, spaces, dashes)
    const tokens = configClean.toUpperCase().split(/[^A-Z0-9]+/);
    const hasValidTool = tokens.some(token => VALID_TOOL_CLASSES.has(token) || token.startsWith("ES") || token.startsWith("EARTHSTAR"));
    if (!hasValidTool) {
      errors.push(`Configuration warning: no registered downhole tool acronyms detected in "${configClean}"`);
    }
  }

  // 7. Footage Validation
  if (row.footage !== null && row.footage !== undefined) {
    if (isNaN(row.footage) || row.footage < 0) {
      errors.push(`Invalid footage value: "${row.footage}" (footage cannot be negative)`);
    }
  }

  // 8. Date and Month Validation
  const dateClean = (row.date || "").trim();
  if (dateClean && dateClean.toUpperCase() !== "N/A" && !dateClean.toLowerCase().includes("invalid")) {
    const parts = dateClean.split("-");
    if (parts.length !== 3) {
      errors.push(`Expected DD-MMM-YY date, received format: "${dateClean}"`);
    } else {
      const day = parseInt(parts[0], 10);
      if (isNaN(day) || day < 1 || day > 31) {
        errors.push(`Invalid day number in log date: "${parts[0]}"`);
      }
      const validMonths = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      if (!validMonths.includes(parts[1])) {
        errors.push(`Invalid month abbreviation: "${parts[1]}" (must be 3-letter, e.g. Jun)`);
      }
    }
  }

  return errors;
};

export const formatMonthName = (val: any): string => {
  if (val === null || val === undefined) return "";
  const str = String(val).trim();
  if (!str || str.toUpperCase() === "N/A") return "";

  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const monthsLower = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

  // 1. Is it already a month name like "Jan", "January", "Feb", "February"?
  const lower = str.toLowerCase();
  const mIdx = monthsLower.findIndex(m => lower === m || lower.startsWith(m));
  if (mIdx !== -1) {
    return months[mIdx];
  }

  // 2. Is it a number?
  const num = Number(str);
  if (!isNaN(num)) {
    // 1 to 12
    if (num >= 1 && num <= 12 && Number.isInteger(num)) {
      return months[num - 1];
    }
    // Excel serial date e.g. 46834 or similar
    if (num > 20000 && num < 60000) {
      const utcDays = num - 25569;
      const dObj = new Date(utcDays * 86400 * 1000);
      if (!isNaN(dObj.getTime())) {
        return months[dObj.getUTCMonth()];
      }
    }
  }

  // 3. Fallback: Parse as date string
  const parsed = parseAnyDate(str);
  if (parsed.month && parsed.month !== "N/A") {
    return parsed.month;
  }

  return str.toUpperCase() === "N/A" ? "" : str;
};

export const parseAnyDate = (
  val: any, 
  dateFormat: "DD/MM" | "MM/DD" = "DD/MM",
  expectedMonth?: string
): { date: string; month: string } => {
  if (val === null || val === undefined) return { date: "", month: "" };
  
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const monthsLower = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

  let dObj: Date | null = null;

  // 1. If it is already a Date object
  if (val instanceof Date && !isNaN(val.getTime())) {
    dObj = val;
  }
  // 2. If it is a number (Excel serial date)
  else if (typeof val === "number") {
    if (val > 20000 && val < 60000) {
      const utcDays = val - 25569;
      dObj = new Date(utcDays * 86400 * 1000);
    }
  }
  // 3. If it is a string
  else {
    const str = String(val).trim();
    if (!str || str.toLowerCase() === "n/a") return { date: "", month: "" };

    // Check if it's a numeric string representing Excel serial date
    if (/^\d+(\.\d+)?$/.test(str)) {
      const numVal = Number(str);
      if (numVal > 20000 && numVal < 60000) {
        const utcDays = numVal - 25569;
        dObj = new Date(utcDays * 86400 * 1000);
      }
    }
    // Check if it's an erroneous ISO string with a 5-digit year (e.g. "46028-01-01" or similar)
    else if (/^\d{5,}-\d{2}-\d{2}$/.test(str)) {
      const yearPart = str.split("-")[0];
      const numVal = Number(yearPart);
      if (numVal > 20000 && numVal < 60000) {
        const utcDays = numVal - 25569;
        dObj = new Date(utcDays * 86400 * 1000);
      }
    }
    else {
      // Clean separators to spaces
      const cleanStr = str.replace(/[\-\.\/]/g, " ").replace(/\s+/g, " ").trim();
      const parts = cleanStr.split(" ");

      if (parts.length === 3) {
        // Case A: YYYY MM DD
        if (parts[0].length === 4 && !isNaN(Number(parts[0]))) {
          const year = Number(parts[0]);
          const mPart = parts[1].toLowerCase();
          const dPart = Number(parts[2]);
          let mIdx = monthsLower.indexOf(mPart);
          if (mIdx === -1 && !isNaN(Number(mPart))) {
            mIdx = Number(mPart) - 1;
          }
          if (mIdx >= 0 && mIdx <= 11 && !isNaN(dPart)) {
            dObj = new Date(year, mIdx, dPart);
          }
        }
        // Case B: DD MMM YY or DD MMM YYYY / MMM DD YYYY
        else {
          const p0 = parts[0].toLowerCase();
          const p1 = parts[1].toLowerCase();
          const p2 = parts[2];

          let mIdx = -1;
          let dVal = -1;
          let yVal = -1;

          const p1MonthIdx = monthsLower.findIndex(m => p1.startsWith(m));
          const p0MonthIdx = monthsLower.findIndex(m => p0.startsWith(m));

          if (p1MonthIdx !== -1) {
            mIdx = p1MonthIdx;
            dVal = Number(parts[0]);
            yVal = Number(parts[2]);
          } else if (p0MonthIdx !== -1) {
            mIdx = p0MonthIdx;
            dVal = Number(parts[1]);
            yVal = Number(parts[2]);
          } else {
            const n1 = Number(parts[0]);
            const n2 = Number(parts[1]);
            const n3 = Number(parts[2]);

            if (!isNaN(n1) && !isNaN(n2) && !isNaN(n3)) {
              yVal = n3;
              if (dateFormat === "DD/MM") {
                dVal = n1;
                mIdx = n2 - 1;
              } else {
                dVal = n2;
                mIdx = n1 - 1;
              }
            }
          }

          if (mIdx >= 0 && mIdx <= 11 && dVal > 0 && dVal <= 31 && yVal >= 0) {
            const fullYear = yVal < 100 ? (yVal + 2000) : yVal;
            dObj = new Date(fullYear, mIdx, dVal);
          }
        }
      } else {
        // Fallback standard JS date parse
        const parsed = Date.parse(str);
        if (!isNaN(parsed)) {
          dObj = new Date(parsed);
        }
      }
    }
  }

  let result = { date: (val === null || val === undefined || String(val).toUpperCase() === "N/A") ? "" : String(val), month: "" };

  if (dObj && !isNaN(dObj.getTime())) {
    let day = dObj.getDate();
    let monthIdx = dObj.getMonth();
    let year = dObj.getFullYear();

    const isISO = typeof val === "string" && (val.includes("T") || val.includes("Z"));

    if (isISO || (dObj.getUTCHours() === 0 && dObj.getUTCMinutes() === 0)) {
      day = dObj.getUTCDate();
      monthIdx = dObj.getUTCMonth();
      year = dObj.getUTCFullYear();
    } else {
      day = dObj.getDate();
      monthIdx = dObj.getMonth();
      year = dObj.getFullYear();
    }

    const dayStr = String(day);
    const monthStr = months[monthIdx];
    const yearStr = String(year).slice(-2);

    result = {
      date: `${dayStr}-${monthStr}-${yearStr}`,
      month: monthStr
    };
  }

  // Self-correcting Swap Mechanism based on explicit month from spreadsheet column (e.g. "Jun", "Jul")
  if (expectedMonth && expectedMonth !== "N/A" && expectedMonth.trim() !== "") {
    const expLower = expectedMonth.trim().toLowerCase();
    const expIdx = monthsLower.findIndex(m => expLower.startsWith(m));
    
    if (expIdx !== -1) {
      const parts = result.date.split("-");
      if (parts.length === 3) {
        const parsedDay = Number(parts[0]);
        const parsedMonthStr = parts[1].toLowerCase();
        const parsedMonthIdx = monthsLower.findIndex(m => parsedMonthStr.startsWith(m));
        
        if (parsedMonthIdx !== -1 && parsedMonthIdx !== expIdx) {
          // Check if parsedDay exactly matches expected Month index + 1
          if (parsedDay === expIdx + 1 && parsedMonthIdx + 1 <= 31) {
            // Day and month are inverted! Swap them.
            const correctedDay = parsedMonthIdx + 1;
            const correctedMonth = months[expIdx];
            const yearStr = parts[2];
            result = {
              date: `${correctedDay}-${correctedMonth}-${yearStr}`,
              month: correctedMonth
            };
          }
        }
      }
    }
  }

  return result;
};

export interface Abbreviation {
  abbreviation: string;
  fullName: string;
  category: string;
  description: string;
}

export const DEFAULT_ABBREVIATIONS: Abbreviation[] = [
  // 1. iCruise Group
  { abbreviation: "IC", fullName: "iCruise", category: "Directional Drilling", description: "iCruise Rotary Steerable System (RSS) for high-precision directional and horizontal drilling." },
  { abbreviation: "ICRUISE", fullName: "iCruise", category: "Directional Drilling", description: "iCruise Rotary Steerable System (RSS) for high-precision directional and horizontal drilling." },

  // 1b. IC BM Group
  { abbreviation: "IC BM", fullName: "IC BM", category: "Directional Drilling", description: "iCruise Bus Master (IC BM) interface and telemetry tool module." },
  { abbreviation: "IC(BM)", fullName: "IC BM", category: "Directional Drilling", description: "iCruise Bus Master (IC BM) interface and telemetry tool module." },
  { abbreviation: "ICBM", fullName: "IC BM", category: "Directional Drilling", description: "iCruise Bus Master (IC BM) interface and telemetry tool module." },
  { abbreviation: "IC-BM", fullName: "IC BM", category: "Directional Drilling", description: "iCruise Bus Master (IC BM) interface and telemetry tool module." },

  // 2. BaseStar Group
  { abbreviation: "BST", fullName: "BaseStar", category: "Acoustic / Sonic", description: "BaseStar acoustic logging and caliper sensor providing stress analysis and cement evaluation." },
  { abbreviation: "BASE", fullName: "BaseStar", category: "Acoustic / Sonic", description: "BaseStar acoustic logging and caliper sensor providing stress analysis and cement evaluation." },
  { abbreviation: "BASESTAR", fullName: "BaseStar", category: "Acoustic / Sonic", description: "BaseStar acoustic logging and caliper sensor providing stress analysis and cement evaluation." },

  // 3. ResiStar Group
  { abbreviation: "RST", fullName: "ResiStar", category: "Resistivity", description: "ResiStar propagation resistivity sensor providing robust logs in all mud types." },
  { abbreviation: "RESISTAR", fullName: "ResiStar", category: "Resistivity", description: "ResiStar propagation resistivity sensor providing robust logs in all mud types." },
  { abbreviation: "RS", fullName: "ResiStar", category: "Resistivity", description: "ResiStar propagation resistivity sensor providing robust logs in all mud types." },

  // 4. StrataStar Group
  { abbreviation: "STS", fullName: "StrataStar", category: "Mechanics", description: "StrataStar slotted tension, torque, and downhole mechanical forces sensor sub." },

  // 5. LithoStar Group
  { abbreviation: "LST", fullName: "LithoStar", category: "Nuclear / Density", description: "LithoStar high-resolution formation density and lithology determination tool." },
  { abbreviation: "LITHO", fullName: "LithoStar", category: "Nuclear / Density", description: "LithoStar high-resolution formation density and lithology determination tool." },

  // 6. Battery Group
  { abbreviation: "BATT", fullName: "Battery", category: "Power", description: "Downhole lithium battery pack providing reliable power to the LWD string." },
  { abbreviation: "BATTERY", fullName: "Battery", category: "Power", description: "Downhole lithium battery pack providing reliable power to the LWD string." },
  { abbreviation: "BATT1", fullName: "Battery", category: "Power", description: "Downhole lithium battery pack providing reliable power to the LWD string." },
  { abbreviation: "BATT2", fullName: "Battery", category: "Power", description: "Downhole lithium battery pack providing reliable power to the LWD string." },
  { abbreviation: "ISTARBATT", fullName: "Battery", category: "Power", description: "Downhole lithium battery pack providing reliable power to the LWD string." },
  { abbreviation: "ISTARBATT1", fullName: "Battery", category: "Power", description: "Downhole lithium battery pack providing reliable power to the LWD string." },
  { abbreviation: "ISTARBATT2", fullName: "Battery", category: "Power", description: "Downhole lithium battery pack providing reliable power to the LWD string." },

  // 6b. EM Battery Group
  { abbreviation: "EM BATT", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EMBATT", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EM BATTERY", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EMBATTERY", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EM BATTERIES", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EMBATTERIES", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EM BATT1", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EM BATT2", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EM BATT 1", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EM BATT 2", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EM-BATT", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EM-BATTERY", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EM-BATTERIES", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },
  { abbreviation: "EM BATT PACK", fullName: "EM Battery", category: "Power", description: "Electromagnetic downhole lithium battery pack supplying high-capacity power." },

  // 7. ILS Group
  { abbreviation: "ILS", fullName: "ILS", category: "Telemetry / MWD", description: "ILS - Integrated Logging System / In-line Stabilizer downhole module." },
  { abbreviation: "HAL50ILS", fullName: "ILS", category: "Telemetry / MWD", description: "ILS - Integrated Logging System / In-line Stabilizer downhole module." },

  // 8. XBAT Group (Renamed to XBAT)
  { abbreviation: "XBAT", fullName: "XBAT", category: "Acoustic / Sonic", description: "Sonic LWD cross-dipole acoustic tool providing high-quality wave data." },

  // 9. AFR Group
  { abbreviation: "AFR", fullName: "AFR", category: "Resistivity", description: "AFR - Azimuthal Focus Resistivity geosteering and formation evaluation tool." },

  // 10. LCS Group
  { abbreviation: "LC", fullName: "LCS", category: "Nuclear / Density", description: "Lithology Density and Caliper tool (LCS)." },
  { abbreviation: "LCS", fullName: "LCS", category: "Nuclear / Density", description: "Lithology Density and Caliper tool (LCS)." },

  // 11. UCS Group
  { abbreviation: "UC", fullName: "UCS", category: "Borehole Caliper", description: "Ultrasonic Caliper sensor (UCS) measuring borehole diameter and shape." },
  { abbreviation: "UCS", fullName: "UCS", category: "Borehole Caliper", description: "Ultrasonic Caliper sensor (UCS) measuring borehole diameter and shape." },
  { abbreviation: "USC", fullName: "UCS", category: "Borehole Caliper", description: "Ultrasonic Caliper sensor (UCS) measuring borehole diameter and shape." },

  // 12. GeoTap Group
  { abbreviation: "GeoTap", fullName: "GeoTap", category: "Pressure", description: "GeoTap Formation Pressure-While-Drilling provides direct downhole formation pressure testing." },
  { abbreviation: "GT", fullName: "GeoTap", category: "Pressure", description: "GeoTap Formation Pressure-While-Drilling provides direct downhole formation pressure testing." },
  { abbreviation: "GEOTAP", fullName: "GeoTap", category: "Pressure", description: "GeoTap Formation Pressure-While-Drilling provides direct downhole formation pressure testing." },

  // 13. ALD Group
  { abbreviation: "ALD", fullName: "ALD", category: "Sensors", description: "Azimuthal Lithodensity tool providing real-time bulk density and caliper logs." },

  // 14. GWD Group
  { abbreviation: "GWD", fullName: "GWD", category: "Directional Drilling", description: "Gyro While Drilling tool providing real-time high-precision surveys." },

  // 15. DGR Group
  { abbreviation: "DGR", fullName: "DGR", category: "Sensors", description: "Dual Gamma Ray tool with two detectors for clay-volume calculation." },

  // 16. PWD Group
  { abbreviation: "PWD", fullName: "PWD", category: "Pressure", description: "Pressure While Drilling sensor measuring annular and drill pipe pressure." },

  // 17. Dummy Flex Group
  { abbreviation: "DUMMYFLEX", fullName: "Dummy Flex", category: "Mechanics", description: "Dummy flex sub or spacer sub used in LWD drill string configuration." },
  { abbreviation: "DF", fullName: "Dummy Flex", category: "Mechanics", description: "Dummy flex sub or spacer sub used in LWD drill string configuration." },

  // 18. Caliper Group
  { abbreviation: "BST CALIPER", fullName: "BST Caliper", category: "Borehole Caliper", description: "Caliper sensor logging borehole dimensions and volume calculations." },
  { abbreviation: "BSTCALIPER", fullName: "BST Caliper", category: "Borehole Caliper", description: "Caliper sensor logging borehole dimensions and volume calculations." },
  { abbreviation: "CALIPER", fullName: "BST Caliper", category: "Borehole Caliper", description: "Caliper sensor logging borehole dimensions and volume calculations." },

  // 19. ADR Group
  { abbreviation: "ADR", fullName: "ADR", category: "Resistivity", description: "Azimuthal Deep Resistivity mapping boundary distances up to 18 feet." },

  // 20. PulseStar Group
  { abbreviation: "PST", fullName: "PulseStar", category: "Telemetry / MWD", description: "High-resolution mud pulse telemetry transmitter." },
  { abbreviation: "PULSESTAR", fullName: "PulseStar", category: "Telemetry / MWD", description: "High-resolution mud pulse telemetry transmitter." },

  // 21. GyroStar Group
  { abbreviation: "GYROSTAR", fullName: "GyroStar", category: "Directional Drilling", description: "Precision gyroscopic directional logging sensor system." },

  // 22. EarthStar Group (linked with ES, ES TX, ES RX, ESTX, ESTX1, ESTX2, ESRX, ESRX1, ESRX2, ESRCVR, ESRCVR1, ESRCVR2)
  { abbreviation: "EarthStar", fullName: "EarthStar", category: "Telemetry / MWD", description: "EarthStar electromagnetic telemetry sub system." },
  { abbreviation: "ES", fullName: "EarthStar", category: "Telemetry / MWD", description: "EarthStar electromagnetic telemetry sub system." },
  { abbreviation: "ES TX", fullName: "EarthStar", category: "Telemetry / MWD", description: "Electromagnetic telemetry transmitter sub." },
  { abbreviation: "ESTX", fullName: "EarthStar", category: "Telemetry / MWD", description: "Electromagnetic telemetry transmitter sub." },
  { abbreviation: "ESTX1", fullName: "EarthStar", category: "Telemetry / MWD", description: "Electromagnetic telemetry transmitter sub 1." },
  { abbreviation: "ESTX2", fullName: "EarthStar", category: "Telemetry / MWD", description: "Electromagnetic telemetry transmitter sub 2." },
  { abbreviation: "ES RX", fullName: "EarthStar", category: "Telemetry / MWD", description: "Electromagnetic telemetry receiver sub." },
  { abbreviation: "ESRX", fullName: "EarthStar", category: "Telemetry / MWD", description: "Electromagnetic telemetry receiver sub." },
  { abbreviation: "ESRX1", fullName: "EarthStar", category: "Telemetry / MWD", description: "Electromagnetic telemetry receiver sub 1." },
  { abbreviation: "ESRX2", fullName: "EarthStar", category: "Telemetry / MWD", description: "Electromagnetic telemetry receiver sub 2." },
  { abbreviation: "ESRCVR", fullName: "EarthStar", category: "Telemetry / MWD", description: "Electromagnetic telemetry receiver sub." },
  { abbreviation: "ESRCVR1", fullName: "EarthStar", category: "Telemetry / MWD", description: "Electromagnetic telemetry receiver sub 1." },
  { abbreviation: "ESRCVR2", fullName: "EarthStar", category: "Telemetry / MWD", description: "Electromagnetic telemetry receiver sub 2." },

  // 23. FDC Group
  { abbreviation: "FDC", fullName: "FDC", category: "Nuclear / Density", description: "Formation Density Caliper tool for porosity and borehole diameter measurement." },

  // 24. HCIM Group
  { abbreviation: "HCIM", fullName: "HCIM", category: "Evaluation", description: "Sensor module designed for early identification and evaluation of hydrocarbon-bearing zones." },

  // 25. PCDC Group
  { abbreviation: "PCDC", fullName: "PCDC", category: "Telemetry / MWD", description: "Controls mud pulse telemetry and integrates essential directional surveying and gamma sensors." },

  // 26. PCGK Group
  { abbreviation: "PCGK", fullName: "PCGK", category: "Sensors", description: "Power-Pulse Control and Gamma Ray tool for directional/gamma measurements." },

  // 27. P4M Group
  { abbreviation: "P4M", fullName: "P4M", category: "Resistivity", description: "Multi-frequency, multi-spacing propagation resistivity tool with directional options." },

  // 28. DHS Group
  { abbreviation: "DHS", fullName: "DHS", category: "Sensors", description: "A modular housing containing critical pressure and temperature sensors near the bit." },

  // 29. Pulsar Group
  { abbreviation: "PULSAR", fullName: "Pulsar", category: "Telemetry / MWD", description: "High-resolution pulsar telemetry and logging tool." },

  // 30. FlowStar Group
  { abbreviation: "FST", fullName: "FlowStar", category: "Fluid Mechanics", description: "FlowStar high-accuracy flow and fluid dynamics measurement sensor." },
  { abbreviation: "FLOWSTAR", fullName: "FlowStar", category: "Fluid Mechanics", description: "FlowStar high-accuracy flow and fluid dynamics measurement sensor." },
  { abbreviation: "FLOW STAR", fullName: "FlowStar", category: "Fluid Mechanics", description: "FlowStar high-accuracy flow and fluid dynamics measurement sensor." },
  { abbreviation: "FLOW-STAR", fullName: "FlowStar", category: "Fluid Mechanics", description: "FlowStar high-accuracy flow and fluid dynamics measurement sensor." },
  { abbreviation: "FLOW", fullName: "FlowStar", category: "Fluid Mechanics", description: "FlowStar high-accuracy flow and fluid dynamics measurement sensor." }
];

export type UserRole = "admin" | "editor" | "viewer";
export type UserStatus = "approved" | "pending" | "rejected";
export type UserPermissions = "read" | "write" | "both";

export interface UserAccount {
  id?: string;
  email: string;
  username?: string;
  password?: string;
  name: string;
  role: UserRole;
  permissions: UserPermissions; // 'read', 'write', 'both'
  status: UserStatus; // 'approved', 'pending', 'rejected'
  title?: string;
  requestedRole?: string;
  requestedAt?: string;
  createdAt?: string;
}

export const INITIAL_DEFAULT_USERS: UserAccount[] = [
  {
    id: "usr_admin",
    email: "admin@drillops.com",
    password: "admin123",
    name: "Admin",
    role: "admin",
    permissions: "both",
    status: "approved",
    title: "Administrator (Full Edit Rights)"
  },
  {
    id: "usr_user",
    email: "user@drillops.com",
    password: "user123",
    name: "User",
    role: "viewer",
    permissions: "read",
    status: "approved",
    title: "Observer (Read-Only Access)"
  }
];


