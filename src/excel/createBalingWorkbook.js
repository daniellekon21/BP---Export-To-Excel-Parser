import { dateSortKey, dateToStr, monthLabel } from "../helpers.js";
import {
  BALING_SHEET_NAMES,
  BALING_FAILED_HEADERS,
  BALING_SCRAP_HEADERS,
  BALING_CRCA_HEADERS,
} from "../config/balingSchemas.js";
import { baseStyles, styleHeaderRow, styleBodyRows, applyColumnWidths } from "./excelCommon.js";

const EXAMPLE_PROD_HEADER_ROW2 = [
  "",
  "Date",
  "Baling Machine \nNumber",
  "Bale Type",
  "Bale Number",
  "Bale Series\nM/Y",
  "Baler Operator",
  "Assistant Baler Operator",
  "Start Time",
  "Finish Time",
  "Total Time to Bale",
  "Passenger",
  "4 X 4",
  "Motorcylce",
  "Light Commercial",
  "Light Commercial T",
  "Light Commercial SW",
  "Light Commercial (Whole)",
  "Heavy Commercial T",
  "Heavy Commercial SW",
  "Heavy Commercial (Whole)",
  "Agricultural T",
  "Agricultural SW",
  "Total Number of T",
  "Total Number of SW",
  "Nylon T",
  "Nylon SW",
  "Tubes",
  "Total Number of Tyres",
  "Bale Weight\nKG",
  "Bale Weight\nTONS",
  "Raw Text",
  "Body Date Text",
];

const EXAMPLE_PROD_HEADER_ROW3 = (() => {
  const row = new Array(33).fill("");
  row[11] = "PCR";     // L3
  row[15] = "RADIALS"; // P3
  row[25] = "NYLONS";  // Z3
  return row;
})();

function sortByDateAndTime(records) {
  const baleSortKey = (row) => {
    const candidates = [row?.baleNumber, row?.baleTestCode];
    for (const c of candidates) {
      const s = String(c || "").trim();
      if (!s) continue;
      const m = s.match(/\bB\s*[- ]?(\d+)\b/i);
      if (m) return parseInt(m[1], 10);
    }
    return Number.MAX_SAFE_INTEGER;
  };

  return [...records].sort((a, b) => {
    const d = dateSortKey(a.chatDateParsed || a.date) - dateSortKey(b.chatDateParsed || b.date);
    if (d !== 0) return d;

    const bk = baleSortKey(a) - baleSortKey(b);
    if (bk !== 0) return bk;

    const aTime = a.startTime ? (a.startTime.h * 60 + a.startTime.m) : 0;
    const bTime = b.startTime ? (b.startTime.h * 60 + b.startTime.m) : 0;
    if (aTime !== bTime) return aTime - bTime;

    return String(a.machine || "").localeCompare(String(b.machine || ""));
  });
}

function num(v) {
  return v === null || v === undefined ? "" : v;
}

function finiteOrNull(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

// Extract canonical bale-code prefix from bale number string.
function extractBalePrefix(baleNumber) {
  if (!baleNumber) return "";
  const m = String(baleNumber).match(/^(PShrB|ConV|CRC|CRS|PCR|PB|CA|TB|CN|SR|CR|B)/i);
  return m ? m[1] : "";
}

function displayBaleNumberOnly(baleCode) {
  const src = String(baleCode || "").trim();
  if (!src) return "";
  const m = src.match(/^(PShrB|ConV|CRC|CRS|PCR|PB|CA|TB|CN|SR|CR|B)\s*[- ]?(\d{1,4})/i);
  if (m) return m[2];
  return src;
}

function canonicalBaleTypePrefix(prefixRaw) {
  const prefix = String(prefixRaw || "").toUpperCase();
  if (prefix === "CA") return "CA";
  if (prefix === "PCR" || prefix === "PB" || prefix === "B") return "PCR";
  if (prefix === "CRC" || prefix === "CR") return "CRC";
  if (prefix === "CRS") return "CRS";
  if (prefix === "TB") return "TB";
  if (prefix === "CN") return "CN";
  if (prefix === "PSHRB") return "PShrB";
  if (prefix === "CONV") return "ConV";
  if (prefix === "SR") return "CRS";
  return "";
}

function formatTimeWithSeconds(t) {
  if (!t) return "";
  const hours = Number.isFinite(t.h) ? t.h : 0;
  const minutes = Number.isFinite(t.m) ? t.m : 0;
  const seconds = Number.isFinite(t.s) ? t.s : 0;
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
}

function toSeconds(timeObj) {
  if (!timeObj || !Number.isFinite(timeObj.h) || !Number.isFinite(timeObj.m)) return null;
  const s = Number.isFinite(timeObj.s) ? timeObj.s : 0;
  return (timeObj.h * 3600) + (timeObj.m * 60) + s;
}

function formatSecondsAsHHMMSS(totalSeconds) {
  if (!Number.isFinite(totalSeconds) || totalSeconds < 0) return "";
  const hrs = Math.floor(totalSeconds / 3600);
  const mins = Math.floor((totalSeconds % 3600) / 60);
  const secs = Math.floor(totalSeconds % 60);
  return `${String(hrs).padStart(2, "0")}:${String(mins).padStart(2, "0")}:${String(secs).padStart(2, "0")}`;
}

function durationSecondsFromRow(r) {
  const startSec = toSeconds(r?.startTime);
  const finishSec = toSeconds(r?.finishTime);
  if (Number.isFinite(startSec) && Number.isFinite(finishSec) && finishSec >= startSec) {
    return finishSec - startSec;
  }

  if (Number.isFinite(r?.durationMinutes) && r.durationMinutes >= 0) {
    return Math.round(r.durationMinutes * 60);
  }

  return null;
}

function formatDurationForRow(r) {
  const seconds = durationSecondsFromRow(r);
  return Number.isFinite(seconds) ? formatSecondsAsHHMMSS(seconds) : "";
}

function parseHHMMSSToSeconds(value) {
  const m = String(value || "").trim().match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
  if (!m) return null;
  const h = Number(m[1]);
  const min = Number(m[2]);
  const s = Number(m[3]);
  if (!Number.isFinite(h) || !Number.isFinite(min) || !Number.isFinite(s)) return null;
  return (h * 3600) + (min * 60) + s;
}

function applyDurationTrafficLight(ws, durationColumnIndex, fromRow = 2) {
  const fills = {
    green: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFC6EFCE" },
    },
    orange: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFE699" },
    },
    red: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF4CCCC" },
    },
  };

  for (let rowIdx = fromRow; rowIdx <= ws.rowCount; rowIdx += 1) {
    const cell = ws.getRow(rowIdx).getCell(durationColumnIndex);
    const seconds = parseHHMMSSToSeconds(cell.value);
    if (!Number.isFinite(seconds)) continue;
    if (seconds < 15 * 60) {
      cell.fill = fills.green;
    } else if (seconds < 20 * 60) {
      cell.fill = fills.orange;
    } else {
      cell.fill = fills.red;
    }
  }
}

function addProductionGroupedHeader(ws) {
  ws.addRow(new Array(EXAMPLE_PROD_HEADER_ROW2.length).fill(""));
  ws.addRow(EXAMPLE_PROD_HEADER_ROW2);
  ws.addRow(EXAMPLE_PROD_HEADER_ROW3);
  // Match template behavior: sublabels are centered across grouped blocks.
  ws.mergeCells(3, 12, 3, 15); // L3:O3 -> PCR
  ws.mergeCells(3, 16, 3, 25); // P3:Y3 -> RADIALS
  ws.mergeCells(3, 26, 3, 28); // Z3:AB3 -> NYLONS
}

function styleProductionHeader(ws, styles) {
  const thin = styles.thinBlack;
  const medium = styles.mediumBlack;
  const groupStarts = new Set([2, 12, 16, 26, 29]); // B, L, P, Z, AC
  const groupEnds = new Set([11, 15, 25, 28, 33]); // K, O, Y, AB, AG

  const fills = [
    { from: 2, to: 15, fgColor: { theme: 4, tint: 0.7999511703848384 } }, // B:O light blue
    { from: 16, to: 25, fgColor: { argb: "FFE97132" } }, // P:Y RADIALS
    { from: 26, to: 28, fgColor: { argb: "FF0D8ABA" } }, // Z:AB NYLONS
    { from: 29, to: 33, fgColor: { theme: 4, tint: 0.7999511703848384 } }, // AC:AG same as PCR
  ];

  const fillForCol = (col) => fills.find((f) => col >= f.from && col <= f.to)?.fgColor || { theme: 4, tint: 0.7999511703848384 };

  ws.getRow(2).height = 48;
  ws.getRow(3).height = 17;

  for (let row = 2; row <= 3; row += 1) {
    for (let col = 2; col <= 33; col += 1) {
      const cell = ws.getRow(row).getCell(col);
      const isRadials = col >= 16 && col <= 25;
      const isNylons = col >= 26 && col <= 28;
      cell.font = {
        bold: true,
        size: 11,
        color: (isRadials || isNylons) ? { argb: "FF000000" } : { theme: 1 },
        name: "Aptos Narrow",
        family: 2,
        scheme: "minor",
      };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: row === 2 };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: fillForCol(col), bgColor: { indexed: 64 } };
      cell.border = {
        left: groupStarts.has(col) ? medium : thin,
        right: groupEnds.has(col) ? medium : thin,
        top: row === 2 ? medium : thin,
        bottom: row === 3 ? medium : thin,
      };
    }
  }
}

function applyNumericFormatting(ws, numericColumns) {
  for (const col of numericColumns) {
    ws.getColumn(col).numFmt = "0";
  }
}

function componentQtySum(r) {
  const values = [
    r.passengerQty,
    r.fourx4Qty,
    r.lcQty,
    r.motorcycleQty,
    r.srQty,
    r.agriQty,
    r.treadQty,
    r.sideWallQty,
  ].map((v) => Number(v));

  const present = values.filter((v) => Number.isFinite(v));
  if (present.length === 0) return null;
  return present.reduce((a, b) => a + b, 0);
}

function componentTyresFromTreadAndSideWall(r) {
  const treads = Number(r?.treadQty);
  const sideWalls = Number(r?.sideWallQty);
  const hasTreads = Number.isFinite(treads) && treads > 0;
  const hasSideWalls = Number.isFinite(sideWalls) && sideWalls > 0;
  if (!hasTreads && !hasSideWalls) return 0;
  if (hasTreads && hasSideWalls) return Math.max(sideWalls / 2, treads);
  if (hasTreads) return treads;
  return sideWalls / 2;
}

function resolvedTotalQty(r) {
  // New-format rows carry a pre-computed total that already applies max(T, floor(SW/2)) per category.
  if (r?.newFormatTotalQty !== null && r?.newFormatTotalQty !== undefined) {
    return r.newFormatTotalQty;
  }

  const baseWholeTyres = [
    r?.passengerQty,
    r?.fourx4Qty,
    r?.motorcycleQty,
    lcWholeForDisplay(r),
  ]
    .map((v) => Number(v))
    .filter((v) => Number.isFinite(v))
    .reduce((a, b) => a + b, 0);

  // LC/HC use the same category rule: whole + max(SW/2, T).
  const lcWhole = finiteOrNull(lcWholeRadialForDisplay(r));
  const lcT = finiteOrNull(lcTreadForDisplay(r));
  const lcSW = finiteOrNull(lcSideWallForDisplay(r));
  const hcWhole = finiteOrNull(hcWholeForDisplay(r));
  const hcT = finiteOrNull(hcTForDisplay(r));
  const hcSW = finiteOrNull(hcSWForDisplay(r));
  const agriT = Number(agriTForDisplay(r));
  const agriSW = Number(agriSWForDisplay(r));

  const hasLcRadial = (Number.isFinite(lcWhole) && lcWhole > 0) || (Number.isFinite(lcT) && lcT > 0) || (Number.isFinite(lcSW) && lcSW > 0);
  const lcTotal = hasLcRadial
    ? (Number.isFinite(lcWhole) ? lcWhole : 0) + Math.max(Number.isFinite(lcSW) ? (lcSW / 2) : 0, Number.isFinite(lcT) ? lcT : 0)
    : 0;

  const hasHcRadial = (Number.isFinite(hcWhole) && hcWhole > 0) || (Number.isFinite(hcT) && hcT > 0) || (Number.isFinite(hcSW) && hcSW > 0);
  const hcTotal = hasHcRadial
    ? (Number.isFinite(hcWhole) ? hcWhole : 0) + Math.max(Number.isFinite(hcSW) ? (hcSW / 2) : 0, Number.isFinite(hcT) ? hcT : 0)
    : 0;

  if (hasLcRadial || hasHcRadial) {
    const radialTyres = lcTotal + hcTotal;
    const hasAgri = (Number.isFinite(agriT) && agriT > 0) || (Number.isFinite(agriSW) && agriSW > 0);
    if (hasAgri) {
      const agriTyres = Math.max(Number.isFinite(agriT) ? agriT : 0, Number.isFinite(agriSW) ? (agriSW / 2) : 0);
      return baseWholeTyres + agriTyres + radialTyres;
    }
    return baseWholeTyres + radialTyres;
  }

  const componentTyres = componentTyresFromTreadAndSideWall(r);
  const calculated = baseWholeTyres + componentTyres;
  if (calculated > 0) return calculated;

  const sum = componentQtySum(r);
  if (sum !== null) return sum;
  return r.totalQty ?? null;
}

function resolvedWeightTons(r) {
  const kg = Number(r?.weightKg);
  if (!Number.isFinite(kg)) return null;
  return Math.round((kg / 1000) * 1000) / 1000;
}

function dateToMonthYear(d) {
  if (!d || !Number.isInteger(d.month) || !Number.isInteger(d.year)) return "";
  return `${String(d.month).padStart(2, "0")}/${String(d.year).slice(-2)}`;
}

function baleSeriesForDisplay(r) {
  const explicit = String(r?.baleSeries || "").trim();
  if (explicit) return explicit;
  return dateToMonthYear(r?.chatDateParsed || r?.date);
}

function hasSideWalls(r) {
  const sw = Number(r?.sideWallQty);
  return Number.isFinite(sw) && sw > 0;
}

function hasLcMention(r) {
  const lc = Number(r?.lcQty);
  if (Number.isFinite(lc) && lc > 0) return true;
  return /\blc\b|light\s*commercial/i.test(String(r?.rawMessage || r?.normalizedMessage || ""));
}

function hasHcMention(r) {
  const hc = Number(r?.hcQty);
  if (Number.isFinite(hc) && hc > 0) return true;
  return /\bhc\b|heavy\s*commercial/i.test(String(r?.rawMessage || r?.normalizedMessage || ""));
}

function lcWholeForDisplay(r) {
  const type = canonicalBaleTypeForRow(r);
  const lc = finiteOrNull(r?.lcQty);
  if (!Number.isFinite(lc)) return null;
  if (type === "CN" || type === "CRC" || type === "CRS" || type === "CA") return null;
  return hasSideWalls(r) ? 0 : lc;
}

function lcWholeRadialForDisplay(r) {
  const type = canonicalBaleTypeForRow(r);
  if (type === "CN" || type === "CA") return null;
  const v = finiteOrNull(r?.lcWholeQty);
  return Number.isFinite(v) ? v : null;
}

function lcTreadForDisplay(r) {
  const type = canonicalBaleTypeForRow(r);
  // Bale-type gating must win over any optional override fields.
  if (type === "CN") return null;
  if (r?.lcTread !== undefined && r.lcTread !== null) return r.lcTread;
  const lc = finiteOrNull(r?.lcQty);
  if (type === "CRC" || type === "CRS") {
    if (!hasLcMention(r)) return null;
    if (Number.isFinite(lc)) return lc;
    // Old-format fallback: generic T line belongs to LC when HC is not mentioned.
    const t = finiteOrNull(r?.treadQty);
    if (Number.isFinite(t) && !hasHcMention(r)) return t;
    return null;
  }
  if (type === "CA") return null;
  if (!Number.isFinite(lc)) return null;
  return hasSideWalls(r) ? lc : null;
}

function lcSideWallForDisplay(r) {
  const type = canonicalBaleTypeForRow(r);
  // Bale-type gating must win over any optional override fields.
  if (type === "CN") return null;
  if (r?.lcSideWall !== undefined && r.lcSideWall !== null) return r.lcSideWall;
  const sw = finiteOrNull(r?.sideWallQty);
  const lc = finiteOrNull(r?.lcQty);
  if (type === "CRC" || type === "CRS") {
    if (!hasLcMention(r) || !Number.isFinite(sw)) return null;
    return sw;
  }
  if (type === "CA") return null;
  if (!Number.isFinite(sw) || !Number.isFinite(lc)) return null;
  return hasSideWalls(r) ? sw : null;
}

function canonicalBaleTypeForRow(r) {
  const code = String(r?.baleNumber || r?.baleTestCode || "");
  const prefix = extractBalePrefix(code);
  return canonicalBaleTypePrefix(prefix);
}

function isCnBale(r) {
  return canonicalBaleTypeForRow(r) === "CN";
}

function isCrcBale(r) {
  return canonicalBaleTypeForRow(r) === "CRC";
}

function isCnLcNylonMode(r) {
  const code = String(r?.baleNumber || r?.baleTestCode || "");
  const canonical = canonicalBaleTypePrefix(extractBalePrefix(code));
  if (canonical !== "CN") return false;
  const lcQty = Number(r?.lcQty);
  if (Number.isFinite(lcQty) && lcQty > 0) return true;
  return /\blc\b|light\s*commercial/i.test(String(r?.rawMessage || r?.normalizedMessage || ""));
}

function totalTForDisplay(r) {
  const lcT = finiteOrNull(lcTreadForDisplay(r));
  const hcT = finiteOrNull(hcTForDisplay(r));
  const agriT = finiteOrNull(agriTForDisplay(r));
  const values = [lcT, hcT, agriT].filter((v) => Number.isFinite(v));
  if (values.length === 0) return null;
  return values.reduce((a, b) => a + b, 0);
}

function totalSWForDisplay(r) {
  const lcSW = finiteOrNull(lcSideWallForDisplay(r));
  const hcSW = finiteOrNull(hcSWForDisplay(r));
  const agriSW = finiteOrNull(agriSWForDisplay(r));
  const values = [lcSW, hcSW, agriSW].filter((v) => Number.isFinite(v));
  if (values.length === 0) return null;
  return values.reduce((a, b) => a + b, 0);
}

function nylonTForDisplay(r) {
  if (isCnBale(r)) {
    const t = finiteOrNull(r?.treadQty);
    if (Number.isFinite(t) && t > 0) return t;
    const lcT = finiteOrNull(r?.lcTread);
    if (Number.isFinite(lcT) && lcT > 0) return lcT;
    const lc = finiteOrNull(r?.lcQty);
    if (Number.isFinite(lc) && lc > 0) return lc;
    const lcWhole = finiteOrNull(r?.lcWholeQty);
    return Number.isFinite(lcWhole) ? lcWhole : null;
  }
  if (isCrcBale(r)) return null;
  if (!isCnLcNylonMode(r)) return null;
  const v = finiteOrNull(r?.treadQty);
  return Number.isFinite(v) ? v : null;
}

function nylonSWForDisplay(r) {
  if (isCnBale(r)) {
    const sw = finiteOrNull(r?.sideWallQty);
    if (Number.isFinite(sw)) return sw;
    const lcSW = finiteOrNull(r?.lcSideWall);
    return Number.isFinite(lcSW) ? lcSW : null;
  }
  if (isCrcBale(r)) return null;
  if (!isCnLcNylonMode(r)) return null;
  const v = finiteOrNull(r?.sideWallQty);
  return Number.isFinite(v) ? v : null;
}

function agriTForDisplay(r) {
  if (r?.agriTread !== undefined && r.agriTread !== null) return r.agriTread;
  const type = canonicalBaleTypeForRow(r);
  if (type === "CA") {
    const t = finiteOrNull(r?.treadQty);
    return Number.isFinite(t) ? t : null;
  }
  const v = finiteOrNull(r?.agriQty);
  return Number.isFinite(v) ? v : null;
}

function agriSWForDisplay(r) {
  if (r?.agriSideWall !== undefined && r.agriSideWall !== null) return r.agriSideWall;
  const type = canonicalBaleTypeForRow(r);
  if (type === "CA") {
    const sw = finiteOrNull(r?.sideWallQty);
    return Number.isFinite(sw) ? sw : null;
  }
  return null;
}

function hcTForDisplay(r) {
  const type = canonicalBaleTypeForRow(r);
  // CN is nylon-only mapping by business rule.
  if (type === "CN") return null;
  if (r?.hcTread !== undefined && r.hcTread !== null) return r.hcTread;
  const hc = finiteOrNull(r?.hcQty);
  if (Number.isFinite(hc)) return hc;
  // Old-format fallback: generic T line belongs to HC when LC is not mentioned.
  const t = finiteOrNull(r?.treadQty);
  if (Number.isFinite(t) && hasHcMention(r) && !hasLcMention(r)) return t;
  return null;
}

function hcSWForDisplay(r) {
  const type = canonicalBaleTypeForRow(r);
  // CN is nylon-only mapping by business rule.
  if (type === "CN") return null;
  if (r?.hcSideWall !== undefined && r.hcSideWall !== null) return r.hcSideWall;
  const sw = finiteOrNull(r?.sideWallQty);
  if (!Number.isFinite(sw)) return null;
  if (type === "CRC") {
    return hasHcMention(r) ? sw : null;
  }
  if (type === "CRS") {
    if (hasHcMention(r)) return sw;
    if (!hasLcMention(r)) return sw;
    return null;
  }
  return null;
}

function hcWholeForDisplay(r) {
  const type = canonicalBaleTypeForRow(r);
  if (type === "CN" || type === "CA") return null;
  const v = finiteOrNull(r?.hcWholeQty);
  return Number.isFinite(v) ? v : null;
}

function isUsableDate(d) {
  if (!d) return false;
  if (!Number.isInteger(d.year) || !Number.isInteger(d.month) || !Number.isInteger(d.day)) return false;
  if (d.month < 1 || d.month > 12) return false;
  if (d.day < 1 || d.day > 31) return false;
  return true;
}

function monthKeyFromDate(d) {
  return `${d.year}-${String(d.month).padStart(2, "0")}`;
}

function groupRowsByMonth(records) {
  const byMonth = new Map();
  const invalid = [];
  for (const r of records) {
    const d = r.chatDateParsed || r.date || null;
    if (!isUsableDate(d)) {
      invalid.push(r);
      continue;
    }
    const key = monthKeyFromDate(d);
    if (!byMonth.has(key)) byMonth.set(key, []);
    byMonth.get(key).push(r);
  }
  return { byMonth, invalid };
}

function appendProductionRow(ws, r) {
  const effectiveRecordType =
    r.recordType ||
    (r.failureType ? "FAILED" : (r.baleTestCode ? "CR_CA" : (r.scrapType ? "SCRAP" : "STANDARD")));

  const baleCode = r.baleNumber || r.baleTestCode || "";
  const balePrefix = extractBalePrefix(baleCode);
  const baleType =
    canonicalBaleTypePrefix(balePrefix) ||
    r.productionType ||
    r.testType ||
    effectiveRecordType;

  const added = ws.addRow([
    "",
    dateToStr(r.chatDateParsed),
    r.machine || "",
    baleType,
    displayBaleNumberOnly(baleCode),
    baleSeriesForDisplay(r),
    r.operator || "",
    r.assistant || "",
    formatTimeWithSeconds(r.startTime),
    formatTimeWithSeconds(r.finishTime),
    formatDurationForRow(r),
    num(r.passengerQty),
    num(r.fourx4Qty),
    num(r.motorcycleQty),
    num(lcWholeForDisplay(r)),
    num(lcTreadForDisplay(r)),
    num(lcSideWallForDisplay(r)),
    num(lcWholeRadialForDisplay(r)),
    num(hcTForDisplay(r)),
    num(hcSWForDisplay(r)),
    num(hcWholeForDisplay(r)),
    num(agriTForDisplay(r)),
    num(agriSWForDisplay(r)),
    num(totalTForDisplay(r)),
    num(totalSWForDisplay(r)),
    num(nylonTForDisplay(r)),
    num(nylonSWForDisplay(r)),
    num(r.tubeQty),        // Total Nylon T — repurposed for Tube qty (TB bales)
    num(resolvedTotalQty(r)),
    num(r.weightKg),
    num(resolvedWeightTons(r)),
    r.rawMessage || "",
    r.bodyDateText || "",
  ]);
  if (baleType === "ConV") {
    for (let col = 1; col <= 33; col += 1) {
      const c = added.getCell(col);
      c.font = { ...(c.font || {}), color: { argb: "FFFF0000" } };
    }
  }
  if (effectiveRecordType === "FAILED" || baleType === "FAILED") {
    const failedCell = added.getCell(4); // Bale Type
    failedCell.font = {
      ...(failedCell.font || {}),
      bold: true,
      color: { argb: "FFFF0000" },
    };
  }
  if (r.usedTimestampFallbackForFutureBodyDate) {
    const dateCell = added.getCell(2);
    dateCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFF2CC" }, // light pastel yellow for future-body-date fallback
    };
  }
}

function sumNumeric(rows, getter) {
  return rows.reduce((acc, row) => {
    const v = Number(getter(row));
    return Number.isFinite(v) ? acc + v : acc;
  }, 0);
}

function appendProductionTotalsRow(ws, rows) {
  const durationValues = rows
    .map((r) => durationSecondsFromRow(r))
    .filter((v) => Number.isFinite(v) && v >= 0);
  const avgDurationSec = durationValues.length
    ? Math.round(durationValues.reduce((a, b) => a + b, 0) / durationValues.length)
    : null;

  const totalPassenger = sumNumeric(rows, (r) => r.passengerQty);
  const total4x4 = sumNumeric(rows, (r) => r.fourx4Qty);
  const totalMotorcycle = sumNumeric(rows, (r) => r.motorcycleQty);
  const totalLc = sumNumeric(rows, (r) => lcWholeForDisplay(r));
  const totalLcT = sumNumeric(rows, (r) => lcTreadForDisplay(r));
  const totalLcSW = sumNumeric(rows, (r) => lcSideWallForDisplay(r));
  const totalLcWholeRadial = sumNumeric(rows, (r) => lcWholeRadialForDisplay(r));
  const totalHc = sumNumeric(rows, (r) => hcTForDisplay(r));
  const totalHcSW = sumNumeric(rows, (r) => hcSWForDisplay(r));
  const totalHcWhole = sumNumeric(rows, (r) => hcWholeForDisplay(r));
  const totalTread = sumNumeric(rows, (r) => totalTForDisplay(r));
  const totalSideWall = sumNumeric(rows, (r) => totalSWForDisplay(r));
  const totalNylonT = sumNumeric(rows, (r) => nylonTForDisplay(r));
  const totalNylonSW = sumNumeric(rows, (r) => nylonSWForDisplay(r));
  const totalAgri = sumNumeric(rows, (r) => agriTForDisplay(r));
  const totalAgriSW = sumNumeric(rows, (r) => agriSWForDisplay(r));
  const totalTube = sumNumeric(rows, (r) => r.tubeQty);
  const totalTyres = sumNumeric(rows, (r) => resolvedTotalQty(r));
  const totalWeightKg = sumNumeric(rows, (r) => r.weightKg);
  const totalWeightTons = Math.round((totalWeightKg / 1000) * 1000) / 1000;

  ws.addRow([
    "",
    "TOTALS",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    avgDurationSec !== null ? formatSecondsAsHHMMSS(avgDurationSec) : "",
    totalPassenger || "",
    total4x4 || "",
    totalMotorcycle || "",
    totalLc || "",
    totalLcT || "",
    totalLcSW || "",
    totalLcWholeRadial || "",
    totalHc || "",   // HC T column
    totalHcSW || "",
    totalHcWhole || "",
    totalAgri || "",
    totalAgriSW || "",
    totalTread || "",
    totalSideWall || "",
    totalNylonT || "",
    totalNylonSW || "",
    totalTube || "", // Tube column
    totalTyres || "",
    totalWeightKg || "",
    totalWeightTons || "",
    "",
    "",
  ]);
}

function appendMonthlySummariesTable(ws, summaries) {
  if (!Array.isArray(summaries) || summaries.length === 0) return;

  ws.addRow([]);
  ws.addRow(["", "Daily Summaries"]);
  ws.addRow([
    "",
    "Chat Date Parsed",
    "Summary Type",
    "Baling Machine Number",
    "Bale Count",
    "Weight Kg",
    "Tons",
    "Passenger Qty",
    "4x4 Qty",
    "LC Qty",
    "Motorcycle Qty",
    "SR Qty",
    "Tread Qty",
    "Side Wall Qty",
    "Agri Qty",
    "Total Tyres",
    "Machine 1 Start Hour",
    "Machine 1 Finish Hour",
    "Machine 2 Start Hour",
    "Machine 2 Finish Hour",
    "Notes / Flags",
    "Raw Message",
  ]);

  for (const r of sortByDateAndTime(summaries)) {
    ws.addRow([
      "",
      dateToStr(r.chatDateParsed),
      r.summaryType || "BALING_DAILY_SUMMARY",
      r.machine || "",
      num(r.baleCount),
      num(r.weightKg),
      num(r.tons),
      num(r.passengerQty),
      num(r.fourx4Qty),
      num(r.lcQty),
      num(r.motorcycleQty),
      num(r.srQty),
      num(r.treadQty),
      num(r.sideWallQty),
      num(r.agriQty),
      num(r.totalTyres),
      r.machine1StartHour || "",
      r.machine1FinishHour || "",
      r.machine2StartHour || "",
      r.machine2FinishHour || "",
      r.notesFlags || "",
      r.rawMessage || "",
    ]);
  }
}

export async function createBalingWorkbook(data = {}) {
  const {
    standardRecords = [],
    failedRecords = [],
    scrapRecords = [],
    crcaRecords = [],
    summaryRecords = [],
    validationLog = [],
  } = data;

  const { default: ExcelJS } = await import("exceljs");
  const wb = new ExcelJS.Workbook();
  const styles = baseStyles();

  const monthSheets = [];
  const logEntries = [...validationLog];
  const allProductionRows = [
    ...standardRecords,
    ...failedRecords,
    ...scrapRecords,
    ...crcaRecords,
  ];
  const { byMonth, invalid: invalidDatedProductionRows } = groupRowsByMonth(allProductionRows);
  const { byMonth: summariesByMonth, invalid: invalidDatedSummaries } = groupRowsByMonth(summaryRecords);
  const sortedMonthKeys = [...byMonth.keys()].sort();

  for (const key of sortedMonthKeys) {
    const ws = wb.addWorksheet(monthLabel(key));
    addProductionGroupedHeader(ws);
    const monthRows = sortByDateAndTime(byMonth.get(key));
    for (const r of monthRows) appendProductionRow(ws, r);
    appendProductionTotalsRow(ws, monthRows);
    appendMonthlySummariesTable(ws, summariesByMonth.get(key) || []);
    monthSheets.push(ws);
  }

  let reviewWs = null;
  if (invalidDatedProductionRows.length > 0) {
    reviewWs = wb.addWorksheet("Review_Missing_Date");
    reviewWs.addRow([...EXAMPLE_PROD_HEADER_ROW2, "Review Reason"]);
    for (const r of invalidDatedProductionRows) {
      reviewWs.addRow([
        "",
        "",
        r.machine || "",
        r.productionType || "Production",
        r.baleNumber || "",
        "",
        r.operator || "",
        r.assistant || "",
        formatTimeWithSeconds(r.startTime),
        formatTimeWithSeconds(r.finishTime),
        formatDurationForRow(r),
        num(r.passengerQty),
        num(r.fourx4Qty),
        num(r.motorcycleQty),
        num(lcWholeForDisplay(r)),
        num(lcTreadForDisplay(r)),
        num(lcSideWallForDisplay(r)),
        num(lcWholeRadialForDisplay(r)),
        num(hcTForDisplay(r)),
        num(hcSWForDisplay(r)),
        num(hcWholeForDisplay(r)),
        num(agriTForDisplay(r)),
        num(agriSWForDisplay(r)),
        num(totalTForDisplay(r)),
        num(totalSWForDisplay(r)),
        num(nylonTForDisplay(r)),
        num(nylonSWForDisplay(r)),
        num(r.tubeQty),
        num(resolvedTotalQty(r)),
        num(r.weightKg),
        num(resolvedWeightTons(r)),
        r.rawMessage || "",
        r.bodyDateText || "",
        "Missing or invalid parsed date - month routing skipped",
      ]);

      logEntries.push({
        severity: "WARNING",
        issueType: "MISSING_OR_INVALID_DATE",
        chatDateParsed: "",
        sourceMessageTimestamp: r.sourceTimestamp || "",
        machine: r.machine || "",
        baleNumberCode: r.baleNumber || r.baleTestCode || "",
        sheetTargetAttempted: "Monthly Production Sheet",
        problemDescription: "Row could not be routed to a month sheet due to missing/invalid parsed date",
        rawMessage: r.rawMessage || "",
        date: "",
      });
    }
  }

  for (const r of invalidDatedSummaries) {
    logEntries.push({
      severity: "WARNING",
      issueType: "MISSING_OR_INVALID_DATE",
      chatDateParsed: "",
      sourceMessageTimestamp: r.sourceTimestamp || "",
      machine: r.machine || "",
      baleNumberCode: "",
      sheetTargetAttempted: "Monthly Summary Table",
      problemDescription: "Daily summary could not be routed to a month sheet due to missing/invalid parsed date",
      rawMessage: r.rawMessage || "",
      date: "",
    });
  }

  const failedWs = wb.addWorksheet(BALING_SHEET_NAMES.failed);
  failedWs.addRow(BALING_FAILED_HEADERS);
  for (const r of sortByDateAndTime(failedRecords)) {
    failedWs.addRow([
      dateToStr(r.chatDateParsed),
      r.machine || "",
      displayBaleNumberOnly(r.baleNumber || ""),
      r.failureType || "FAILED_BALE",
      r.failureReason || "",
      r.operator || "",
      r.assistant || "",
      formatTimeWithSeconds(r.startTime),
      formatTimeWithSeconds(r.finishTime),
      num(r.passengerQty),
      num(r.fourx4Qty),
      num(r.lcQty),
      num(r.motorcycleQty),
      num(r.srQty),
      num(r.agriQty),
      num(r.treadQty),
      num(r.sideWallQty),
      num(resolvedTotalQty(r)),
      num(r.weightKg),
      r.notesFlags || "",
      r.rawMessage || "",
    ]);
  }

  const scrapWs = wb.addWorksheet(BALING_SHEET_NAMES.scrap);
  scrapWs.addRow(BALING_SCRAP_HEADERS);
  for (const r of sortByDateAndTime(scrapRecords)) {
    scrapWs.addRow([
      dateToStr(r.chatDateParsed),
      r.machine || "",
      r.productionLabel || "",
      displayBaleNumberOnly(r.baleNumber || ""),
      r.scrapType || "Scrap",
      num(r.scrapQty),
      num(r.weightKg),
      r.operator || "",
      r.assistant || "",
      formatTimeWithSeconds(r.startTime),
      formatTimeWithSeconds(r.finishTime),
      r.notesFlags || "",
      r.rawMessage || "",
    ]);
  }

  const crcaWs = wb.addWorksheet(BALING_SHEET_NAMES.crca);
  crcaWs.addRow(BALING_CRCA_HEADERS);
  for (const r of sortByDateAndTime(crcaRecords)) {
    crcaWs.addRow([
      dateToStr(r.chatDateParsed),
      r.machine || "",
      r.baleTestCode || r.baleNumber || "",
      r.testType || "Test",
      r.recordType || "CR_CA",
      r.operator || "",
      r.assistant || "",
      formatTimeWithSeconds(r.startTime),
      formatTimeWithSeconds(r.finishTime),
      formatDurationForRow(r),
      num(r.treadQty),
      num(r.sideWallQty),
      num(r.passengerQty),
      num(r.fourx4Qty),
      num(r.lcQty),
      num(r.motorcycleQty),
      num(r.srQty),
      num(r.agriQty),
      num(r.weightKg),
      r.notesFlags || "",
      r.rawMessage || "",
    ]);
  }

  const logWs = wb.addWorksheet(BALING_SHEET_NAMES.validation);
  logWs.addRow([
    "Severity",
    "Issue Type",
    "Chat Date Parsed",
    "Baling Machine Number",
    "Bale Number / Code",
    "Sheet Target Attempted",
    "Problem Description",
    "Raw Message",
  ]);

  for (const e of logEntries) {
    logWs.addRow([
      e.severity || "WARNING",
      e.issueType || "PARSER_WARNING",
      e.chatDateParsed || e.date || "",
      e.machine || "",
      e.baleNumberCode || "",
      e.sheetTargetAttempted || "",
      e.problemDescription || e.issue || "",
      e.rawMessage || "",
    ]);
  }

  const sheets = [...monthSheets, failedWs, scrapWs, crcaWs, logWs];
  if (reviewWs) sheets.push(reviewWs);
  for (const ws of sheets) {
    if (monthSheets.includes(ws)) {
      styleProductionHeader(ws, styles);
      styleBodyRows(ws, 4, ws.rowCount, styles.baseBorder);
      ws.views = [{ state: "frozen", xSplit: 1, ySplit: 3, topLeftCell: "B4" }];
      ws.autoFilter = { from: { row: 2, column: 2 }, to: { row: 2, column: 33 } };
    } else {
      styleHeaderRow(ws.getRow(1), styles, false);
      styleBodyRows(ws, 2, ws.rowCount, styles.baseBorder);
      ws.views = [{ state: "frozen", ySplit: 1 }];
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: ws.columnCount } };
    }
  }

  for (const ws of monthSheets) {
    applyColumnWidths(ws, [
      13,
      10.6640625,
      13,
      10.1640625,
      12.5,
      13,
      14.5,
      16.1640625,
      11.5,
      12.5,
      16.83203125,
      10.83203125,
      13,
      11.5,
      17.33203125,
      17.5,
      19.1640625,
      16,
      18.6640625,
      16,
      21,
      17.5,
      18.6640625,
      14.1640625,
      15.1640625,
      18.6640625,
      20.33203125,
      20.83203125,
      14,
      14,
      13,
      65,
      16,
    ]);
    applyNumericFormatting(ws, [12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]);
    ws.getColumn(31).numFmt = "0.000";
  }
  applyColumnWidths(failedWs, [14, 14, 10, 14, 22, 18, 18, 10, 10, 10, 10, 10, 10, 12, 10, 10, 12, 10, 10, 28, 80]);
  applyColumnWidths(scrapWs, [14, 14, 24, 10, 20, 10, 10, 18, 18, 10, 10, 28, 80]);
  applyColumnWidths(crcaWs, [14, 14, 18, 10, 10, 18, 18, 10, 10, 10, 10, 12, 10, 10, 10, 12, 10, 10, 10, 28, 80]);
  applyColumnWidths(logWs, [10, 22, 14, 14, 18, 20, 44, 90]);
  if (reviewWs) applyColumnWidths(reviewWs, [13, 10.6640625, 13, 10.1640625, 12.5, 13, 14.5, 16.1640625, 11.5, 12.5, 16.83203125, 10.83203125, 13, 11.5, 17.33203125, 17.5, 19.1640625, 16, 18.6640625, 16, 21, 17.5, 18.6640625, 14.1640625, 15.1640625, 18.6640625, 20.33203125, 20.83203125, 14, 14, 13, 65, 16, 42]);

  // Numeric formatting for qty/weight/duration columns.
  applyNumericFormatting(failedWs, [10, 11, 12, 13, 14, 15, 16, 17, 18, 19]);
  applyNumericFormatting(scrapWs, [6, 7]);
  applyNumericFormatting(crcaWs, [11, 12, 13, 14, 15, 16, 17, 18, 19]);
  if (reviewWs) {
    applyNumericFormatting(reviewWs, [12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]);
    reviewWs.getColumn(31).numFmt = "0.000";
  }

  for (const ws of monthSheets) applyDurationTrafficLight(ws, 11, 4);
  applyDurationTrafficLight(crcaWs, 10);
  if (reviewWs) applyDurationTrafficLight(reviewWs, 11);

  return wb;
}
