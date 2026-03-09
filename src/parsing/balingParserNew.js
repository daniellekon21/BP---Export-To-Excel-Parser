/**
 * balingParserNew.js
 *
 * New WhatsApp baling production format parser (Jan 2026+).
 *
 * Message structure:
 *   DD/MM/YYYY
 *   BM #: N
 *
 *   PREFIX###-MM/YYYY
 *   Operator: Name
 *   Ass: Name
 *   Start: HH:MM
 *   Finish: HH:MM
 *   Total Time: N minutes
 *   <material blocks>
 *   Weight: N kg
 *
 * Exports:
 *   parseBalingMessagesNew(text)          — new format first, old format fallback
 *   parseBalingMessagesOldWithFallback(text) — old format first, new format fallback
 */

import { splitWhatsAppMessages, dateToStr, formatTime } from "../helpers.js";
import {
  normalizeBalingText,
  parseLabelTime,
  minutesBetween,
  parseWeightKg,
} from "./commonParsingUtils.js";
import {
  normalizeBalePrefix,
  processBalingMessageOldFormat,
  isFailedBaleMessage,
  extractFailedBaleReason,
} from "./balingParser.js";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function tsToString(tsDate) {
  if (!tsDate) return "";
  return `${tsDate.year}/${String(tsDate.month).padStart(2, "0")}/${String(tsDate.day).padStart(2, "0")}`;
}

function isValidDateObject(d) {
  if (!d) return false;
  if (!Number.isInteger(d.year) || !Number.isInteger(d.month) || !Number.isInteger(d.day)) return false;
  if (d.month < 1 || d.month > 12 || d.day < 1 || d.day > 31) return false;
  return true;
}

function dateObjToSortable(d) {
  return d.year * 10000 + d.month * 100 + d.day;
}

function todayDateObj() {
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth() + 1, day: now.getDate() };
}

function emptyResult() {
  return {
    standardRecords: [],
    failedRecords: [],
    scrapRecords: [],
    crcaRecords: [],
    summaryRecords: [],
    validationLog: [],
    ignoredMessages: 0,
    allRecords: [],
  };
}

// ---------------------------------------------------------------------------
// Validation log helper (mirrors balingParser.js logValidation)
// ---------------------------------------------------------------------------

function logValidation(result, msg, row, details) {
  const entry = {
    severity: details.severity || "WARNING",
    issueType: details.issueType || "PARSER_WARNING",
    chatDateParsed: row?.chatDateParsed ? dateToStr(row.chatDateParsed) : "",
    sourceMessageTimestamp: row?.sourceTimestamp || tsToString(msg?.tsDate),
    machine: row?.machine || "",
    baleNumberCode: details.baleNumberCode || row?.baleNumber || "",
    sheetTargetAttempted: details.sheetTargetAttempted || "Bales_Production",
    problemDescription: details.problemDescription || "",
    rawMessage: msg?.body || row?.rawMessage || "",
    date: row?.chatDateParsed ? dateToStr(row.chatDateParsed) : "",
    messageType: details.issueType || "PARSER_WARNING",
    issue: details.problemDescription || "",
    action: details.action || "Logged",
  };
  result.validationLog.push(entry);
}

// ---------------------------------------------------------------------------
// Duplicate detection (mirrors balingParser.js validateRecord for production)
// ---------------------------------------------------------------------------

function checkDuplicate(row, result, msg, seenProductionBales) {
  if (!row.baleNumber) return true;
  const key = `${dateToStr(row.chatDateParsed)}|${row.baleNumber}`;
  if (seenProductionBales.has(key)) {
    logValidation(result, msg, row, {
      severity: "WARNING",
      issueType: "DUPLICATE_BALE_NUMBER",
      sheetTargetAttempted: "Bales_Production",
      baleNumberCode: row.baleNumber,
      problemDescription: `Duplicate bale number on same date (${row.baleNumber}); latest kept`,
    });
    return false;
  }
  seenProductionBales.add(key);
  return true;
}

function validateNewFormatRecord(row, result, msg, seenProductionBales) {
  if (row.startTime && row.finishTime && row.durationMinutes === null) {
    logValidation(result, msg, row, {
      severity: "WARNING",
      issueType: "TIME_SEQUENCE_INVALID",
      sheetTargetAttempted: "Bales_Production",
      problemDescription: `Finish before start (${formatTime(row.startTime)} -> ${formatTime(row.finishTime)})`,
    });
  }
  if (!row.weightKg || row.weightKg <= 0) {
    logValidation(result, msg, row, {
      severity: "WARNING",
      issueType: "MISSING_WEIGHT",
      sheetTargetAttempted: "Bales_Production",
      problemDescription: "Missing weight on successful bale",
    });
  }
  return checkDuplicate(row, result, msg, seenProductionBales);
}

// ---------------------------------------------------------------------------
// New-format bale code detection
// ---------------------------------------------------------------------------

// All production prefixes (longer before shorter to avoid partial matches).
const NEW_FORMAT_BALE_CODE_RE =
  /^(PShrB|ConV|CRC|CRS|PCR|PB|CA|TB|CN|SR|CR)(\d{1,4})-(\d{1,2}\/\d{4})$/im;

const BM_RE = /\bBM\b[\s#:=\-]*(\d{1,2})/i;

function detectNewFormatBaleCode(normalized) {
  const m = normalized.match(NEW_FORMAT_BALE_CODE_RE);
  if (!m) return null;
  const prefix = normalizeBalePrefix(m[1]);
  const seq = m[2].padStart(3, "0");
  return { baleNumber: `${prefix}${seq}`, suffix: m[3], canonicalPrefix: prefix };
}

function detectNewFormatMachine(normalized) {
  const m = normalized.match(BM_RE);
  if (!m) return "";
  const num = parseInt(m[1], 10);
  if (!Number.isFinite(num) || num <= 0) return "";
  return `BM - ${num}`;
}

// ---------------------------------------------------------------------------
// Date extraction — top-of-message date takes priority
// ---------------------------------------------------------------------------

function extractTopDate(normalized) {
  // Match a standalone date line near the top (first 5 non-empty lines)
  const lines = normalized.split("\n").map((s) => s.trim()).filter(Boolean).slice(0, 8);
  for (const line of lines) {
    const m = line.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      return {
        day: parseInt(m[1], 10),
        month: parseInt(m[2], 10),
        year: parseInt(m[3], 10),
      };
    }
  }
  return null;
}

function resolveDate(bodyDate, tsDate) {
  if (isValidDateObject(bodyDate)) {
    if (dateObjToSortable(bodyDate) > dateObjToSortable(todayDateObj())) {
      return tsDate || bodyDate;
    }
    return bodyDate;
  }
  return tsDate || null;
}

// ---------------------------------------------------------------------------
// Operator / Assistant detection
// ---------------------------------------------------------------------------

function cleanPerson(raw) {
  let s = String(raw || "").trim();
  s = s.replace(/^[\s:,\-]+|[\s:,\-]+$/g, "").trim();
  if (/^(?:n\/?a|na|null|none)$/i.test(s)) return "";
  return s;
}

function detectNewOperator(normalized) {
  const m = normalized.match(/\boperator\s*[:=\-]\s*([^\n]+)/i);
  if (!m) return "";
  return cleanPerson(m[1].split(/\bassistant\b|\bass\b/i)[0]);
}

function detectNewAssistant(normalized) {
  // Match "Ass:" or "Assistant:" variants
  const m = normalized.match(/\b(?:ass(?:istant)?)\s*[:=\-]\s*([^\n]+)/i);
  if (!m) return "";
  return cleanPerson(m[1].split(/\b(?:start\s*time|finish\s*time|start|finish|weight|total|item|date)\b/i)[0]);
}

// ---------------------------------------------------------------------------
// Total Time parsing
// ---------------------------------------------------------------------------

function parseTotalTimeMinutes(normalized) {
  const m = normalized.match(/\btotal\s*time\s*[:\-]?\s*(\d+)\s*min/i);
  return m ? parseInt(m[1], 10) : null;
}

function parseTbQty(normalized) {
  // Quantity line for TB bales: "TB: 14", "tb-14", "TB=14" (case-insensitive)
  // Requires a separator so bale code lines like TB001-01/2026 won't match.
  const m = String(normalized || "").match(/(?:^|\n)\s*tb\s*[:=\-]\s*(\d+)\s*(?:\n|$)/i);
  if (!m) return null;
  const v = parseInt(m[1], 10);
  return Number.isFinite(v) ? v : null;
}

// ---------------------------------------------------------------------------
// Material block parser
// ---------------------------------------------------------------------------

/**
 * Category quantity rule: max(T, floor(SW/2))
 * Returns null if both T and SW are absent.
 */
function categoryQty(t, sw) {
  const hasT = Number.isFinite(t) && t > 0;
  const hasSW = Number.isFinite(sw) && sw > 0;
  if (!hasT && !hasSW) return null;
  if (hasT && hasSW) return Math.max(t, Math.floor(sw / 2));
  if (hasT) return t;
  return Math.floor(sw / 2);
}

// Category header patterns
const CAT_AGRI_RE = /^agri(?:cultural)?$/i;
const CAT_LC_RE = /^(?:lc|light\s*commercial)$/i;
const CAT_HC_RE = /^(?:hc|heavy\s*commercial)$/i;
const CAT_LC_WHOLE_RE = /^(?:lc|light\s*commercial)(?:\s+full)?\s*[:=\-]\s*(\d+)$/i;
const CAT_HC_FULL_RE = /^(?:hc|heavy\s*commercial)\s+full\s*[:\-]?\s*(\d+)$/i;
const CAT_HC_CUT_RE = /^(?:hc|heavy\s*commercial)\s+cut$/i;
const CAT_PCR_RE = /^pcr$/i;
const CAT_CRC_RE = /^crc$/i;
const TB_QTY_RE = /^tb\s*[:=\-]\s*(\d+)$/i;

// Sub-value patterns
const T_RE = /^\bT\s*[:\-]\s*(\d+)$/i;
const SW_RE = /^SW\s*[:\-]\s*(\d+)$/i;
const LC_T_RE = /^(?:lc|light\s*commercial)\s*t\s*[:=\-]\s*(\d+)$/i;
const LC_SW_RE = /^(?:lc|light\s*commercial)\s*sw\s*[:=\-]\s*(\d+)$/i;
const HC_T_RE = /^(?:hc|heavy\s*commercial)\s*t\s*[:=\-]\s*(\d+)$/i;
const HC_SW_RE = /^(?:hc|heavy\s*commercial)\s*sw\s*[:=\-]\s*(\d+)$/i;
const AGRI_T_RE = /^(?:agri|agricultural)\s*t\s*[:=\-]\s*(\d+)$/i;
const AGRI_SW_RE = /^(?:agri|agricultural)\s*sw\s*[:=\-]\s*(\d+)$/i;
const MC_RE = /^(?:mc|motorcycle)\s*[:\-]\s*(\d+)$/i;
const PASS_RE = /^(?:pass(?:enger)?)\s*[:\-]\s*(\d+)$/i;
const FOURX4_RE = /^(?:4\s*[xX×]\s*4)\s*[:\-]\s*(\d+)$/i;

function parseInt10(s) {
  const n = parseInt(s, 10);
  return Number.isFinite(n) ? n : null;
}

/**
 * Parse material blocks from the normalized body.
 * Returns a partial row object with per-category fields and computed total.
 */
function parseMaterialBlocks(lines, canonicalPrefix = "") {
  const result = {
    // Per-category T/SW fields (new-format specific)
    lcTread: null,
    lcSideWall: null,
    hcTread: null,
    hcSideWall: null,
    hcWhole: null,
    lcWhole: null,
    agriTread: null,
    agriSideWall: null,
    tubeQty: null,
    // Standard whole-tyre fields (reuse existing row schema)
    passengerQty: null,
    fourx4Qty: null,
    motorcycleQty: null,
    // Validation
    unknownCategories: [],
  };

  let currentCat = null; // "agri" | "lc" | "hc" | "hcFull" | "pcr" | "crc" | null

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) continue;

    // Skip meta lines that should not be interpreted as category headers
    if (/^(?:operator|ass(?:istant)?|start|finish|total\s*time|weight|bm\b)/i.test(line)) {
      currentCat = null;
      continue;
    }

    // Explicit per-category T/SW lines
    const lcTMatch = line.match(LC_T_RE);
    if (lcTMatch) {
      const v = parseInt10(lcTMatch[1]);
      if (v !== null) result.lcTread = (result.lcTread || 0) + v;
      currentCat = "lc";
      continue;
    }
    const lcSWMatch = line.match(LC_SW_RE);
    if (lcSWMatch) {
      const v = parseInt10(lcSWMatch[1]);
      if (v !== null) result.lcSideWall = (result.lcSideWall || 0) + v;
      currentCat = "lc";
      continue;
    }
    const hcTMatch = line.match(HC_T_RE);
    if (hcTMatch) {
      const v = parseInt10(hcTMatch[1]);
      if (v !== null) result.hcTread = (result.hcTread || 0) + v;
      currentCat = "hc";
      continue;
    }
    const hcSWMatch = line.match(HC_SW_RE);
    if (hcSWMatch) {
      const v = parseInt10(hcSWMatch[1]);
      if (v !== null) result.hcSideWall = (result.hcSideWall || 0) + v;
      currentCat = "hc";
      continue;
    }
    const agriTMatch = line.match(AGRI_T_RE);
    if (agriTMatch) {
      const v = parseInt10(agriTMatch[1]);
      if (v !== null) result.agriTread = (result.agriTread || 0) + v;
      currentCat = "agri";
      continue;
    }
    const agriSWMatch = line.match(AGRI_SW_RE);
    if (agriSWMatch) {
      const v = parseInt10(agriSWMatch[1]);
      if (v !== null) result.agriSideWall = (result.agriSideWall || 0) + v;
      currentCat = "agri";
      continue;
    }

    // T: N
    const tMatch = line.match(T_RE);
    if (tMatch) {
      const v = parseInt10(tMatch[1]);
      if (currentCat === "agri" && v !== null) result.agriTread = (result.agriTread || 0) + v;
      else if (currentCat === "lc" && v !== null) result.lcTread = (result.lcTread || 0) + v;
      else if ((currentCat === "hc" || currentCat === "hcFull") && v !== null) result.hcTread = (result.hcTread || 0) + v;
      else if (String(canonicalPrefix || "").toUpperCase() === "CA" && v !== null) result.agriTread = (result.agriTread || 0) + v;
      continue;
    }

    // SW: N
    const swMatch = line.match(SW_RE);
    if (swMatch) {
      const v = parseInt10(swMatch[1]);
      if (currentCat === "agri" && v !== null) result.agriSideWall = (result.agriSideWall || 0) + v;
      else if (currentCat === "lc" && v !== null) result.lcSideWall = (result.lcSideWall || 0) + v;
      else if ((currentCat === "hc" || currentCat === "hcFull") && v !== null) result.hcSideWall = (result.hcSideWall || 0) + v;
      else if (String(canonicalPrefix || "").toUpperCase() === "CA" && v !== null) result.agriSideWall = (result.agriSideWall || 0) + v;
      continue;
    }

    // MC: N / Motorcycle: N
    const mcMatch = line.match(MC_RE);
    if (mcMatch) {
      const v = parseInt10(mcMatch[1]);
      if (v !== null) result.motorcycleQty = (result.motorcycleQty || 0) + v;
      continue;
    }

    // Pass: N / Passenger: N
    const passMatch = line.match(PASS_RE);
    if (passMatch) {
      const v = parseInt10(passMatch[1]);
      if (v !== null) result.passengerQty = (result.passengerQty || 0) + v;
      continue;
    }

    // 4X4: N / 4x4: N
    const fxMatch = line.match(FOURX4_RE);
    if (fxMatch) {
      const v = parseInt10(fxMatch[1]);
      if (v !== null) result.fourx4Qty = (result.fourx4Qty || 0) + v;
      continue;
    }

    // TB: N (tube quantity for TB bale type)
    const tbQtyMatch = line.match(TB_QTY_RE);
    if (tbQtyMatch) {
      const v = parseInt10(tbQtyMatch[1]);
      if (v !== null) result.tubeQty = (result.tubeQty || 0) + v;
      continue;
    }

    // Category whole-qty lines
    const lcWholeMatch = line.match(CAT_LC_WHOLE_RE);
    if (lcWholeMatch) {
      const v = parseInt10(lcWholeMatch[1]);
      if (v !== null) result.lcWhole = (result.lcWhole || 0) + v;
      currentCat = "lc";
      continue;
    }

    // Category header detection
    if (CAT_HC_FULL_RE.test(line)) {
      const hcFullMatch = line.match(CAT_HC_FULL_RE);
      const v = parseInt10(hcFullMatch[1]);
      if (v !== null) result.hcWhole = (result.hcWhole || 0) + v;
      currentCat = "hcFull";
      continue;
    }

    if (CAT_HC_CUT_RE.test(line)) {
      currentCat = "hc";
      continue;
    }

    if (CAT_HC_RE.test(line)) {
      currentCat = "hc";
      continue;
    }

    if (CAT_AGRI_RE.test(line)) {
      currentCat = "agri";
      continue;
    }

    if (CAT_LC_RE.test(line)) {
      currentCat = "lc";
      continue;
    }

    if (CAT_PCR_RE.test(line)) {
      currentCat = "pcr";
      continue;
    }

    if (CAT_CRC_RE.test(line)) {
      currentCat = "crc";
      continue;
    }

    // Skip bale code lines and Weight lines — they are header/footer, not categories
    if (/^(PShrB|ConV|CRC|CRS|PCR|PB|CA|TB|CN|SR|CR)\d{1,4}-\d{1,2}\/\d{4}$/i.test(line)) continue;
    if (/^weight\s*[:\-]/i.test(line)) continue;

    // If the line has a digit and looks like a category/value but wasn't matched, note it
    if (/\d/.test(line) && /^[A-Za-z]/.test(line) && !/^\d{1,2}\/\d{2}\/\d{4}/.test(line)) {
      result.unknownCategories.push(line);
    }
  }

  return result;
}

/**
 * Compute the total qty from material block results.
 * Rule: max(T, floor(SW/2)) per T/SW category + whole-tyre counts.
 */
function computeNewFormatTotalQty(mat, canonicalPrefix = "") {
  const p = String(canonicalPrefix || "").toUpperCase();
  if (p === "TB") {
    return Number.isFinite(mat.tubeQty) && mat.tubeQty > 0 ? mat.tubeQty : null;
  }
  let total = 0;

  const agriContrib = categoryQty(mat.agriTread, mat.agriSideWall);
  if (agriContrib !== null) total += agriContrib;

  if (p === "CN") {
    // CN uses LC as nylon tread source; do not add LC whole separately.
    const cnTread = Number.isFinite(mat.lcTread) ? mat.lcTread : mat.lcWhole;
    const cnContrib = categoryQty(cnTread, mat.lcSideWall);
    if (cnContrib !== null) total += cnContrib;
  } else {
    const lcContrib = categoryQty(mat.lcTread, mat.lcSideWall);
    if (lcContrib !== null) total += lcContrib;
    if (Number.isFinite(mat.lcWhole) && mat.lcWhole > 0) total += mat.lcWhole;
  }

  const hcContrib = categoryQty(mat.hcTread, mat.hcSideWall);
  if (hcContrib !== null) total += hcContrib;

  if (Number.isFinite(mat.hcWhole) && mat.hcWhole > 0) total += mat.hcWhole;

  if (Number.isFinite(mat.passengerQty) && mat.passengerQty > 0) total += mat.passengerQty;
  if (Number.isFinite(mat.fourx4Qty) && mat.fourx4Qty > 0) total += mat.fourx4Qty;
  if (Number.isFinite(mat.motorcycleQty) && mat.motorcycleQty > 0) total += mat.motorcycleQty;

  return total > 0 ? total : null;
}

function mapWholeFieldsByBaleType(canonicalPrefix, mat) {
  const p = String(canonicalPrefix || "").toUpperCase();
  if (p === "PCR") {
    return {
      lcQty: Number.isFinite(mat.lcWhole) ? mat.lcWhole : null,
      lcWholeQty: null,
      hcWholeQty: null,
    };
  }
  return {
    lcQty: null,
    lcWholeQty: Number.isFinite(mat.lcWhole) ? mat.lcWhole : null,
    hcWholeQty: Number.isFinite(mat.hcWhole) ? mat.hcWhole : null,
  };
}

// ---------------------------------------------------------------------------
// New-format production message parser
// ---------------------------------------------------------------------------

/**
 * Attempt to parse a single message as a new-format production bale.
 * Returns a row object on success, null if the message doesn't match.
 */
function tryParseNewFormatProduction(msg, normalized) {
  // Quick-reject: must have a BM line and a new-format bale code line
  if (!BM_RE.test(normalized)) return null;
  const baleInfo = detectNewFormatBaleCode(normalized);
  if (!baleInfo) return null;

  const machine = detectNewFormatMachine(normalized);
  if (!machine) return null;

  const bodyDate = extractTopDate(normalized);
  const parsedDate = resolveDate(bodyDate, msg.tsDate || null);

  const startTime = parseLabelTime(normalized, "start") || parseLabelTime(normalized, "starting");
  const finishTime = parseLabelTime(normalized, "finish") || parseLabelTime(normalized, "end");
  const durationMinutes = minutesBetween(startTime, finishTime) ?? parseTotalTimeMinutes(normalized);

  const weightKg = parseWeightKg(normalized);

  const operator = detectNewOperator(normalized);
  const assistant = detectNewAssistant(normalized);

  // Parse material blocks from lines (skip header metadata lines)
  const allLines = normalized.split("\n");
  const mat = parseMaterialBlocks(allLines, baleInfo.canonicalPrefix);
  if (baleInfo.canonicalPrefix.toUpperCase() === "TB") {
    if (!Number.isFinite(mat.tubeQty)) mat.tubeQty = parseTbQty(normalized);
  }
  const newFormatTotalQty = computeNewFormatTotalQty(mat, baleInfo.canonicalPrefix);
  const baleTypeMapping = mapWholeFieldsByBaleType(baleInfo.canonicalPrefix, mat);

  const row = {
    sourceTimestamp: tsToString(msg.tsDate),
    chatDateParsed: parsedDate,
    date: parsedDate,
    bodyDate: bodyDate || null,
    bodyDateText: bodyDate ? dateToStr(bodyDate) : "",
    usedTimestampFallbackForFutureBodyDate:
      isValidDateObject(bodyDate) &&
      dateObjToSortable(bodyDate) > dateObjToSortable(todayDateObj()),
    machine,
    operator,
    assistant,
    startTime,
    finishTime,
    durationMinutes,
    weightKg,
    totalQty: newFormatTotalQty,
    rawMessage: msg.body,
    normalizedMessage: normalized,

    // Bale identification
    baleNumber: baleInfo.baleNumber,
    baleSeries: baleInfo.suffix || "",
    productionType: "Production",
    recordType: "STANDARD",

    // Per-category fields (new-format specific — drive display functions in createBalingWorkbook)
    lcTread: mat.lcTread,
    lcSideWall: mat.lcSideWall,
    lcWholeQty: baleTypeMapping.lcWholeQty,
    hcTread: mat.hcTread,
    hcSideWall: mat.hcSideWall,
    hcWhole: baleTypeMapping.hcWholeQty,
    hcWholeQty: baleTypeMapping.hcWholeQty,
    agriTread: mat.agriTread,
    agriSideWall: mat.agriSideWall,

    // Standard whole-tyre fields (PCR sub-counts)
    passengerQty: mat.passengerQty,
    fourx4Qty: mat.fourx4Qty,
    motorcycleQty: mat.motorcycleQty,

    // Fields that old-format display functions may reference — set to null for new format
    // (actual values are in the per-category fields above)
    lcQty: baleTypeMapping.lcQty,
    hcQty: null,
    srQty: null,
    agriQty: null,
    treadQty: null,
    sideWallQty: null,
    tubeQty: Number.isFinite(mat.tubeQty) ? mat.tubeQty : null,
    otherItemRaw: mat.unknownCategories.join(" | "),

    // Pre-computed total qty (used by resolvedTotalQty in createBalingWorkbook)
    newFormatTotalQty,
    parsedByNewFormat: true,
    newFormatFallbackUsed: null,

    notesFlags: "",
  };

  return row;
}

// ---------------------------------------------------------------------------
// Validation + warning emission for new-format rows
// ---------------------------------------------------------------------------

function warnIfMissingFields(row, result, msg) {
  if (!row.operator) {
    logValidation(result, msg, row, {
      severity: "WARNING",
      issueType: "MISSING_OPERATOR",
      problemDescription: "Operator field absent",
    });
  }
  if (!row.assistant) {
    logValidation(result, msg, row, {
      severity: "WARNING",
      issueType: "MISSING_ASSISTANT",
      problemDescription: "Assistant (Ass) field absent",
    });
  }
  if (!row.startTime) {
    logValidation(result, msg, row, {
      severity: "WARNING",
      issueType: "MISSING_START_TIME",
      problemDescription: "Start time absent",
    });
  }
  if (!row.finishTime) {
    logValidation(result, msg, row, {
      severity: "WARNING",
      issueType: "MISSING_FINISH_TIME",
      problemDescription: "Finish time absent",
    });
  }
  if (row.otherItemRaw) {
    logValidation(result, msg, row, {
      severity: "WARNING",
      issueType: "UNKNOWN_MATERIAL_CATEGORY",
      problemDescription: `Unrecognised material category lines: ${row.otherItemRaw}`,
    });
  }
}

// ---------------------------------------------------------------------------
// Public parser: new format first, old format fallback
// ---------------------------------------------------------------------------

export function parseBalingMessagesNew(text) {
  const messages = splitWhatsAppMessages(text);
  const result = emptyResult();
  const seenProductionBales = new Set();

  for (const msg of messages) {
    const normalized = normalizeBalingText(msg.body);

    // Try new format first on every message
    const row = tryParseNewFormatProduction(msg, normalized);
    if (row) {
      warnIfMissingFields(row, result, msg);
      const keep = validateNewFormatRecord(row, result, msg, seenProductionBales);
      if (keep) result.standardRecords.push(row);
      continue;
    }

    // New format produced no result → fallback to old format
    const handled = processBalingMessageOldFormat(msg, normalized, result, seenProductionBales);
    if (handled) {
      logValidation(result, msg, null, {
        severity: "INFO",
        issueType: "PARSER_FALLBACK_USED",
        problemDescription: "New-format parse produced no result; old format fallback used",
        sheetTargetAttempted: "",
      });
    }
    // handled=false → both parsers found nothing → silent skip
  }

  result.allRecords = [
    ...result.standardRecords,
    ...result.failedRecords,
    ...result.scrapRecords,
    ...result.crcaRecords,
  ];

  return result;
}

// ---------------------------------------------------------------------------
// Public parser: old format first, new format fallback
// ---------------------------------------------------------------------------

export function parseBalingMessagesOldWithFallback(text) {
  const messages = splitWhatsAppMessages(text);
  const result = emptyResult();
  const seenProductionBales = new Set();

  for (const msg of messages) {
    const normalized = normalizeBalingText(msg.body);

    // Try old format first; returns true if it produced any record
    const handled = processBalingMessageOldFormat(msg, normalized, result, seenProductionBales);
    if (handled) continue;

    // Old format ignored this message → try new format as fallback
    const row = tryParseNewFormatProduction(msg, normalized);
    if (row) {
      warnIfMissingFields(row, result, msg);
      logValidation(result, msg, row, {
        severity: "INFO",
        issueType: "PARSER_FALLBACK_USED",
        problemDescription: "Old format could not parse; new format fallback used",
        sheetTargetAttempted: "Bales_Production",
      });
      const keep = validateNewFormatRecord(row, result, msg, seenProductionBales);
      if (keep) result.standardRecords.push(row);
    }
    // If both fail → silent skip
  }

  result.allRecords = [
    ...result.standardRecords,
    ...result.failedRecords,
    ...result.scrapRecords,
    ...result.crcaRecords,
  ];

  return result;
}
