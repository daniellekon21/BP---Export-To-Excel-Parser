import { splitWhatsAppMessages, dateToStr, formatTime } from "../helpers.js";
import {
  normalizeBalingText,
  parseLabelTime,
  minutesBetween,
  parseWeightKg,
  parseTotalQty,
} from "./commonParsingUtils.js";
import { BALING_CATEGORY_ALIASES } from "../config/balingSchemas.js";

function tsToString(tsDate) {
  if (!tsDate) return "";
  return `${tsDate.year}/${String(tsDate.month).padStart(2, "0")}/${String(tsDate.day).padStart(2, "0")}`;
}

function extractBodyDate(body) {
  const text = String(body || "");
  let m = text.match(/\bdate\s*[:\-]\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/i);
  // Common real-world variant: a standalone line with date only (no "Date -" label).
  if (!m) m = text.match(/(?:^|\n)\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*(?:\n|$)/i);
  if (!m) return null;

  const parsedYear = parseInt(m[3], 10);
  const year = m[3].length === 2 ? (2000 + parsedYear) : parsedYear;
  return {
    day: parseInt(m[1], 10),
    month: parseInt(m[2], 10),
    year,
  };
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
  return {
    year: now.getFullYear(),
    month: now.getMonth() + 1,
    day: now.getDate(),
  };
}

function resolveParsedDate(bodyDate, tsDate) {
  if (isValidDateObject(bodyDate)) {
    // If the operator-entered date is in the future (e.g. typo year 2034),
    // fall back to WhatsApp timestamp date.
    if (dateObjToSortable(bodyDate) > dateObjToSortable(todayDateObj())) {
      return tsDate || bodyDate;
    }
    return bodyDate;
  }
  return tsDate || null;
}

function isFutureBodyDate(bodyDate) {
  return isValidDateObject(bodyDate) && dateObjToSortable(bodyDate) > dateObjToSortable(todayDateObj());
}

function detectMachine(body) {
  const src = String(body || "");
  const m = src.match(/\b(?:machine|baler|bm)\s*#?\s*[:=\-]?\s*(\d{1,2}|one|two|three|four|five|six|seven|eight|nine|ten)\b/i);
  if (!m) return "";

  const raw = m[1].toLowerCase();
  const wordToNum = {
    one: 1, two: 2, three: 3, four: 4, five: 5,
    six: 6, seven: 7, eight: 8, nine: 9, ten: 10,
  };

  const num = Number.isFinite(Number(raw)) ? Number(raw) : wordToNum[raw];
  if (!Number.isFinite(num) || num <= 0) return "";
  return `BM - ${num}`;
}

// Bale code prefixes in priority order (longer prefixes before shorter ones
// to avoid PShrB being swallowed by the P of Passenger, etc.).
const BALE_CODE_RE = /\b(PShrB|CRC|CRS|PCR|PB|CA|TB|CN|CR)\s*[- ]?\s*(\d{1,4})(?:\s*[-–]\s*\d{1,2}\/\d{2,4})?\b/i;

// Map legacy PB → current PCR label.
// Exported so balingParserNew.js can reuse the same normalisation logic.
export function normalizeBalePrefix(raw) {
  const up = raw.toUpperCase();
  return up === "PB" ? "PCR" : up;
}

function detectBaleNumber(body) {
  const src = String(body || "");
  // New-format codes: CA282-02/2026, PB013-02/2025, CRC001-02/2025, etc.
  const m = src.match(BALE_CODE_RE);
  if (m) {
    const prefix = normalizeBalePrefix(m[1]);
    const seq = m[2].padStart(3, "0");
    return `${prefix}${seq}`;
  }
  // Legacy B-number format: B505, B662
  const old = src.match(/\bB\s*[- ]?(\d{1,4})\b/i);
  if (!old) return "";
  return `B${old[1]}`;
}

function detectProductionType(body) {
  if (/\bproduction\b/i.test(body)) return "Production";
  if (/\btest\b/i.test(body)) return "Test";
  return "";
}

function cleanPersonField(raw) {
  let s = String(raw || "").trim();
  s = s.replace(/^[\s:,\-]+|[\s:,\-]+$/g, "").trim();
  if (/^(?:n\/?a|na|null|none)$/i.test(s)) return "";
  return s;
}

function detectOperator(body) {
  const m = String(body || "").match(/\boperator\s*[:=\-]\s*([^\n]+)/i);
  if (!m) return "";
  const onlyOperator = m[1].split(/\bassistant\b/i)[0];
  return cleanPersonField(onlyOperator);
}

function detectAssistant(body) {
  const m = String(body || "").match(/\b(?:assistant|ass)\s*[:=\-]\s*([^\n]+)/i);
  if (!m) return "";
  const onlyAssistant = m[1].split(/\b(?:start\s*time|finish\s*time|item|total|weight|date|process)\b/i)[0];
  return cleanPersonField(onlyAssistant);
}

function detectSummaryType(body) {
  if (/scrap/i.test(body) && /side\s*wall|\bsr\b/i.test(body)) return "SCRAP_SIDEWALL_DAILY_SUMMARY";
  if (/\bcr\b|\bca\b/i.test(body) && /summary/i.test(body)) return "CR_CA_DAILY_SUMMARY";
  return "BALING_DAILY_SUMMARY";
}

const FAILED_PHRASE_PATTERNS = [
  /\bfailed?\s*bales?\b/i,
  /\bfail\s*bale\b/i,
  /\bbale\s*failed\b/i,
  /\bfailed?\s*bails?\b/i,
  /\bwire'?s?\s*snapp?ed\b/i,
  /\bwires?\s*snapp?ed\b/i,
  /\brope\s*snapp?ed\b/i,
  /\bstraps?\s*snapp?ed\b/i,
  /\bbale\s*popped\b/i,
  /\bpopped\s*bale\b/i,
  /\bsnapp?ed\s*during\s*baling\b/i,
  /\bnot\s+a\s+bale\b/i,
  /\breport\s+failed\s+separately\b/i,
];

const FAILED_REASON_PATTERNS = [
  { pattern: /\bwire'?s?\s*snapp?ed\b/i, reason: "wire snapped" },
  { pattern: /\bwires?\s*snapp?ed\b/i, reason: "wires snapped" },
  { pattern: /\brope\s*snapp?ed\b/i, reason: "rope snapped" },
  { pattern: /\bstraps?\s*snapp?ed\b/i, reason: "straps snapped" },
  { pattern: /\bbale\s*popped\b|\bpopped\s*bale\b/i, reason: "bale popped" },
  { pattern: /\bsnapp?ed\s+during\s+baling\b/i, reason: "snapped during baling" },
  { pattern: /\bnot\s+a\s+bale\b/i, reason: "not a bale" },
  { pattern: /\bfailed?\s+during\s+baling\b/i, reason: "failed during baling" },
];

const FAILED_CONTEXT_PATTERN = /\bbale|baling|machine|operator|assistant|start|finish|item|qty|weight|date|wire|rope|strap|popped\b/i;
const FAILED_ADMIN_HINT_PATTERN = /\breport\b|\bseparately\b|\bchange\b|\baccount\b|\bconsumables\b/i;

export function extractFailedBaleReason(text) {
  const source = String(text || "");
  for (const entry of FAILED_REASON_PATTERNS) {
    if (entry.pattern.test(source)) return entry.reason;
  }
  const explicit = source.match(/\bfailed?\s*bales?(?:\s*due\s*to)?\s*[:\-]?\s*([^\n.,;]+)/i);
  if (explicit) return explicit[1].trim().toLowerCase();
  return "";
}

export function isFailedBaleMessage(text) {
  const source = String(text || "");
  const hasFailurePhrase = FAILED_PHRASE_PATTERNS.some((p) => p.test(source));
  const hasContext = FAILED_CONTEXT_PATTERN.test(source);
  if (!hasFailurePhrase || !hasContext) {
    return { isFailed: false, uncertain: false, conflictingSignals: false, reason: "" };
  }

  const hasBaleNumber = /\bB\s*\d{1,4}\b/i.test(source);
  const hasStructuredBaleData = /\bmachine|operator|assistant|start|finish|item|weight|total|date\b/i.test(source);
  const hasProductionSignal = /\bproduction\b/i.test(source);
  const hasFailureReason = extractFailedBaleReason(source) !== "";
  const uncertain = !hasStructuredBaleData || (FAILED_ADMIN_HINT_PATTERN.test(source) && !hasBaleNumber && !hasFailureReason);
  const conflictingSignals = hasProductionSignal && !hasFailureReason;

  return {
    isFailed: true,
    uncertain,
    conflictingSignals,
    reason: extractFailedBaleReason(source),
  };
}

function detectMessageSubtypeNonFailed(normalized) {
  const t = normalized.toLowerCase();
  if (!t || t.length < 12) return "ignore";
  if (/<media omitted>|this message was deleted|deleted this message/.test(t)) return "ignore";

  const hasSummary = /\bdaily\s*summary\b|\bsummary\b/.test(t);
  const hasScrap = /\bscrap\b|radial\s*side\s*walls?|side\s*wall\s*radial/i.test(t);
  const hasExplicitTest = /\btest\b/i.test(t);

  // New-format production bale codes (CA, PCR, PB→PCR, CRC, CRS, TB, CN, PShrB, CR).
  const hasNewBaleCode = BALE_CODE_RE.test(normalized);
  // Legacy B-number format: B505, B662.
  const hasOldBaleCode = /\bB\s*\d{1,4}\b/i.test(normalized);

  // Old-style CR/CA code that is NOT one of the recognised new production bale codes
  // (e.g. a bare "CR-01" test reference without a full bale record).
  const hasLegacyCrcaCode = !hasNewBaleCode && !hasOldBaleCode && /\b(?:cr|ca)\s*[- ]?\d{1,3}\b/i.test(normalized);

  if (hasSummary) return "summary";
  if (hasScrap && !hasOldBaleCode && !hasNewBaleCode) return "scrap";
  if (hasLegacyCrcaCode) return "crca";
  if ((hasOldBaleCode || hasNewBaleCode) && hasExplicitTest) return "test";
  if (hasOldBaleCode || hasNewBaleCode) return "standard";
  return "ignore";
}

function mapAliasToKey(segment) {
  for (const alias of BALING_CATEGORY_ALIASES) {
    if (alias.pattern.test(segment)) return alias.key;
  }
  return null;
}

function extractItemPairs(normalized) {
  const items = {
    passenger: 0,
    fourx4: 0,
    lc: 0,
    lcWhole: 0,
    hc: 0,
    hcWhole: 0,
    motorcycle: 0,
    sr: 0,
    agri: 0,
    tread: 0,
    sideWall: 0,
    tube: 0,
    otherRaw: [],
  };
  const warnings = [];

  const lines = normalized.split("\n").map((s) => s.trim()).filter(Boolean);
  const itemLikeLines = lines.filter((line) => {
    const hasItemSignal = /\bitem\b|\bpassengers?\b|\bPass\.|\b4x4\b|\blight\s*commercial\b|\blc\b|\bmotorcycle\b|\bheavy\s*commercial\b|\bHC\b|\bagri\b|\bsr\b|\bside\s*wall\b|\bsw\b|\btreads?\b|\blct\b|\bhct\b|\btubes?\b|\bT\s*[-:]\s*\d/i.test(line);
    const hasStrongItemPrefix = /\bitem\b/i.test(line);
    const isMetaLine = /\bdate\b|\boperator\b|\bassistant\b|\bstart\b|\bfinish\b|\bmachine\b/i.test(line) && !hasStrongItemPrefix;
    return hasItemSignal && !isMetaLine;
  });

  const sourceText = itemLikeLines.length > 0 ? itemLikeLines.join(", ") : normalized;

  const cleaned = sourceText
    .replace(/\bweight\s*[:\-]?\s*\d+(?:\.\d+)?\s*kg\b/gi, "")
    .replace(/\btotal\s*(?:qty)?\s*[:\-]?\s*\d+\b/gi, "")
    .replace(/\bitem\s*[:\-]/gi, "")
    .replace(/\bqty\b/gi, "")
    .replace(/\s-\s*(?=(?:passengers?|4x4|light commercial|lc\b|heavy commercial|HC\b|motorcycle|sr\b|agri\b|treads?\b|side wall|sw\b|lct\b|hct\b|tubes?\b))/gi, ", ")
    .replace(/\n/g, ", ");

  const chunks = cleaned.split(",").map((s) => s.trim()).filter(Boolean);
  for (const chunk of chunks) {
    const hcWholeMatch = chunk.match(/\bhc\s*full\b\s*[:\-]?\s*(\d+)\b/i);
    if (hcWholeMatch) {
      items.hcWhole += Number(hcWholeMatch[1]);
      continue;
    }
    const lcWholeMatch = chunk.match(/\blc\s*full\b\s*[:\-]?\s*(\d+)\b/i);
    if (lcWholeMatch) {
      items.lcWhole += Number(lcWholeMatch[1]);
      continue;
    }

    const key = mapAliasToKey(chunk);
    const nums = chunk.match(/\d+/g);
    const qty = nums ? Number(nums[nums.length - 1]) : null;

    if (key && Number.isFinite(qty)) {
      items[key] += qty;
      continue;
    }

    const reverse = chunk.match(/^(\d+)\s+(.+)$/);
    if (reverse) {
      const revKey = mapAliasToKey(reverse[2]);
      if (revKey) {
        items[revKey] += Number(reverse[1]);
        continue;
      }
    }

    if (/\d/.test(chunk) && /item|qty|pass|4x4|lc|motor|side|tread|agri|sr/i.test(chunk)) {
      warnings.push(`Unknown category segment: "${chunk}"`);
      items.otherRaw.push(chunk);
    }
  }

  return { items, warnings };
}

function parseSummaryMetrics(normalized) {
  const { items, warnings } = extractItemPairs(normalized);
  const baleCountMatch = normalized.match(/\b(?:bales?|no\.?\s*of\s*bales?)\s*[:\-]?\s*(\d+)\b/i);
  const tonsMatch = normalized.match(/\b(\d+(?:\.\d+)?)\s*tons?\b/i);
  const m1 = normalized.match(/machine\s*1[^\n]*?(\d{1,2}:\d{2})[^\n]*?(\d{1,2}:\d{2})/i);
  const m2 = normalized.match(/machine\s*2[^\n]*?(\d{1,2}:\d{2})[^\n]*?(\d{1,2}:\d{2})/i);

  return {
    items,
    warnings,
    baleCount: baleCountMatch ? Number(baleCountMatch[1]) : null,
    tons: tonsMatch ? Number(tonsMatch[1]) : null,
    machine1Start: m1 ? m1[1] : "",
    machine1Finish: m1 ? m1[2] : "",
    machine2Start: m2 ? m2[1] : "",
    machine2Finish: m2 ? m2[2] : "",
  };
}

function parseCrcaCode(normalized) {
  const m = normalized.match(/\b(CR|CA)\s*[- ]?\s*(\d{1,3})(?:\s+(\d{1,2}\/\d{2,4}))?/i);
  if (!m) return "";
  const serial = String(m[2]).padStart(2, "0");
  const series = m[3] ? ` ${m[3]}` : "";
  return `${m[1].toUpperCase()}-${serial}${series}`;
}

function baleTypeFromCode(code) {
  const m = String(code || "").match(/^(PSHRB|CONV|CRC|CRS|PCR|PB|CA|TB|CN|SR|CR|B)/i);
  if (!m) return "";
  const p = m[1].toUpperCase();
  if (p === "PB" || p === "B") return "PCR";
  if (p === "SR") return "CRS";
  if (p === "CONV") return "ConV";
  if (p === "PSHRB") return "PShrB";
  return p;
}

function parseTbQty(normalized) {
  const m = String(normalized || "").match(/\btb\s*[:=\-]\s*(\d+)\b/i);
  if (!m) return null;
  const n = Number(m[1]);
  return Number.isFinite(n) ? n : null;
}

function detectScrapQty(body) {
  const patterns = [
    /\bscrap[^\n]*?[:\-]\s*(\d+)\b/i,
    /\bside\s*walls?\s*[:\-]\s*(\d+)\b/i,
    /\bsr\s*[:\-]\s*(\d+)\b/i,
    /\b(\d+)\s*x\s+.*\bscrap\s*bales?\b/i,
  ];
  for (const p of patterns) {
    const m = String(body || "").match(p);
    if (m) return Number(m[1]);
  }
  return null;
}

function parseFailure(normalized) {
  const reason = extractFailedBaleReason(normalized);
  return {
    failureType: "FAILED_BALE",
    failureReason: reason || "",
  };
}

function makeBaseRow(msg, normalized) {
  const bodyDate = extractBodyDate(normalized);
  const usedTimestampFallbackForFutureBodyDate = isFutureBodyDate(bodyDate);
  const parsedDate = resolveParsedDate(bodyDate, msg.tsDate || null);
  const startTime = parseLabelTime(normalized, "start") || parseLabelTime(normalized, "starting");
  const finishTime = parseLabelTime(normalized, "finish") || parseLabelTime(normalized, "end");
  const durationMinutes = minutesBetween(startTime, finishTime);

  return {
    sourceTimestamp: tsToString(msg.tsDate),
    chatDateParsed: parsedDate,
    date: parsedDate,
    bodyDate: bodyDate || null,
    bodyDateText: bodyDate ? dateToStr(bodyDate) : "",
    usedTimestampFallbackForFutureBodyDate,
    machine: detectMachine(normalized),
    operator: detectOperator(normalized),
    assistant: detectAssistant(normalized),
    startTime,
    finishTime,
    durationMinutes,
    weightKg: parseWeightKg(normalized),
    totalQty: parseTotalQty(normalized),
    rawMessage: msg.body,
    normalizedMessage: normalized,
  };
}

function logValidation(result, msg, row, details) {
  const entry = {
    severity: details.severity || "WARNING",
    issueType: details.issueType || "PARSER_WARNING",
    chatDateParsed: row?.chatDateParsed ? dateToStr(row.chatDateParsed) : "",
    sourceMessageTimestamp: row?.sourceTimestamp || tsToString(msg.tsDate),
    machine: row?.machine || "",
    baleNumberCode: details.baleNumberCode || row?.baleNumber || row?.baleTestCode || "",
    sheetTargetAttempted: details.sheetTargetAttempted || "",
    problemDescription: details.problemDescription || "",
    rawMessage: msg.body || row?.rawMessage || "",
    // Compatibility fields used by current UI filter logic.
    date: row?.chatDateParsed ? dateToStr(row.chatDateParsed) : "",
    messageType: details.issueType || "PARSER_WARNING",
    issue: details.problemDescription || "",
    action: details.action || "Logged",
  };
  result.validationLog.push(entry);
}

function validateRecord(row, subtype, result, msg, seenProductionBales) {
  const categorySum =
    (row.passengerQty || 0) +
    (row.fourx4Qty || 0) +
    (row.lcQty || 0) +
    (row.motorcycleQty || 0) +
    (row.srQty || 0) +
    (row.agriQty || 0) +
    (row.treadQty || 0) +
    (row.sideWallQty || 0);

  const shouldCheckTotalMismatch = subtype === "Bales_Production" || subtype === "Failed_Bales";
  if (shouldCheckTotalMismatch && row.totalQty !== null && row.totalQty !== undefined && categorySum > 0 && row.totalQty !== categorySum) {
    // Treat operator-entered total as advisory: auto-correct to parsed category sum.
    row.totalQty = categorySum;
  }

  if (subtype === "Bales_Production" && (!row.weightKg || row.weightKg <= 0)) {
    logValidation(result, msg, row, {
      severity: "WARNING",
      issueType: "MISSING_WEIGHT",
      sheetTargetAttempted: subtype,
      problemDescription: "Missing weight on successful bale",
    });
  }

  if (row.startTime && row.finishTime && row.durationMinutes === null) {
    logValidation(result, msg, row, {
      severity: "WARNING",
      issueType: "TIME_SEQUENCE_INVALID",
      sheetTargetAttempted: subtype,
      problemDescription: `Finish before start (${formatTime(row.startTime)} -> ${formatTime(row.finishTime)})`,
    });
  }

  if (subtype === "Bales_Production" && row.baleNumber) {
    const key = `${dateToStr(row.chatDateParsed)}|${row.baleNumber}`;
    if (seenProductionBales.has(key)) {
      logValidation(result, msg, row, {
        severity: "WARNING",
        issueType: "DUPLICATE_BALE_NUMBER",
        sheetTargetAttempted: subtype,
        baleNumberCode: row.baleNumber,
        problemDescription: `Duplicate bale number on same date (${row.baleNumber}); latest kept`,
      });
      return false;
    }
    seenProductionBales.add(key);
  }

  return true;
}

/**
 * Process a single WhatsApp message using old-format baling logic.
 * Mutates `result` and `seenProductionBales` in place.
 * Returns true if the message was recognised and produced at least one record
 * (including summaries/scraps/failed); returns false if the message was ignored.
 * Exported so balingParserNew.js can reuse this as a fallback.
 */
export function processBalingMessageOldFormat(msg, normalized, result, seenProductionBales) {
  const failedClass = isFailedBaleMessage(normalized);
  const subtype = failedClass.isFailed ? "failed" : detectMessageSubtypeNonFailed(normalized);

  if (subtype === "ignore") {
    result.ignoredMessages += 1;
    return false;
  }

  const base = makeBaseRow(msg, normalized);

  if (subtype === "summary") {
    const metrics = parseSummaryMetrics(normalized);
    const summary = {
      ...base,
      summaryType: detectSummaryType(normalized),
      machine: base.machine || "",
      baleCount: metrics.baleCount,
      tons: metrics.tons,
      passengerQty: metrics.items.passenger || null,
      fourx4Qty: metrics.items.fourx4 || null,
      lcQty: metrics.items.lc || null,
      motorcycleQty: metrics.items.motorcycle || null,
      srQty: metrics.items.sr || null,
      agriQty: metrics.items.agri || null,
      treadQty: metrics.items.tread || null,
      sideWallQty: metrics.items.sideWall || null,
      totalTyres: parseTotalQty(normalized),
      machine1StartHour: metrics.machine1Start,
      machine1FinishHour: metrics.machine1Finish,
      machine2StartHour: metrics.machine2Start,
      machine2FinishHour: metrics.machine2Finish,
      notesFlags: metrics.warnings.join(" | "),
    };
    result.summaryRecords.push(summary);

    if (metrics.warnings.length > 0) {
      logValidation(result, msg, summary, {
        severity: "WARNING",
        issueType: "SUMMARY_PARSE_PARTIAL",
        sheetTargetAttempted: "Daily_Summaries",
        problemDescription: metrics.warnings.join(" | "),
        action: "Summary row emitted with partial mapping",
      });
    }
    return true;
  }

  if (subtype === "scrap") {
    const scrap = {
      ...base,
      productionLabel: (() => {
        const m = normalized.match(/\bproduction\s*[:\-]\s*([^\n]+)/i);
        return m ? m[1].trim() : "";
      })(),
      baleNumber: detectBaleNumber(normalized),
      scrapType: /radial/i.test(normalized) ? "Scrap Radial Sidewall" : "Scrap",
      scrapQty: detectScrapQty(normalized),
      notesFlags: "",
    };

    const issues = [];
    if (!scrap.scrapQty) issues.push("Missing scrap quantity");
    if (!scrap.weightKg) issues.push("Missing weight for scrap record");
    scrap.notesFlags = issues.join(" | ");
    result.scrapRecords.push(scrap);

    for (const issue of issues) {
      logValidation(result, msg, scrap, {
        severity: "WARNING",
        issueType: issue.includes("quantity") ? "SCRAP_QTY_MISSING" : "MISSING_WEIGHT",
        sheetTargetAttempted: "Scrap_Sidewalls",
        problemDescription: issue,
      });
    }
    return true;
  }

  const { items, warnings } = extractItemPairs(normalized);
  const rowBase = {
    ...base,
    baleNumber: detectBaleNumber(normalized),
    productionType: detectProductionType(normalized),
    passengerQty: items.passenger || null,
    fourx4Qty: items.fourx4 || null,
    lcQty: items.lc || null,
      hcQty: items.hc || null,
      hcWholeQty: items.hcWhole || null,
      lcWholeQty: items.lcWhole || null,
      motorcycleQty: items.motorcycle || null,
    srQty: items.sr || null,
    agriQty: items.agri || null,
    treadQty: items.tread || null,
    sideWallQty: items.sideWall || null,
    tubeQty: items.tube || null,
    otherItemRaw: items.otherRaw.join(" | "),
    recordType: subtype.toUpperCase(),
    notesFlags: warnings.join(" | "),
  };

  const baleType = baleTypeFromCode(rowBase.baleNumber);
  if (baleType === "TB") {
    const tbQty = parseTbQty(normalized);
    rowBase.tubeQty = tbQty;
    // TB rows should not populate non-TB tyre categories from free-form T/SW/LC text.
    rowBase.passengerQty = null;
    rowBase.fourx4Qty = null;
    rowBase.lcQty = null;
    rowBase.hcQty = null;
    rowBase.hcWholeQty = null;
    rowBase.lcWholeQty = null;
    rowBase.motorcycleQty = null;
    rowBase.srQty = null;
    rowBase.agriQty = null;
    rowBase.treadQty = null;
    rowBase.sideWallQty = null;
    if (tbQty !== null && (rowBase.totalQty === null || rowBase.totalQty === undefined)) {
      rowBase.totalQty = tbQty;
    }
    if (tbQty === null) {
      logValidation(result, msg, rowBase, {
        severity: "WARNING",
        issueType: "MALFORMED_TB",
        sheetTargetAttempted: subtype === "failed" ? "Failed_Bales" : "Bales_Production",
        problemDescription: "TB bale missing TB quantity (expected TB: N / TB- N / TB= N)",
      });
    }
  }

  for (const warn of warnings) {
    logValidation(result, msg, rowBase, {
      severity: "WARNING",
      issueType: "UNKNOWN_CATEGORY",
      sheetTargetAttempted: subtype === "failed" ? "Failed_Bales" : "Bales_Production",
      problemDescription: warn,
    });
  }

  if (subtype === "failed") {
    const failure = parseFailure(normalized);
    const failed = {
      ...rowBase,
      failureType: failure.failureType,
      failureReason: failure.failureReason || failedClass.reason || "",
    };
    result.failedRecords.push(failed);
    logValidation(result, msg, failed, {
      severity: "INFO",
      issueType: "FAILED_BALE_REROUTED",
      sheetTargetAttempted: "Failed_Bales",
      problemDescription: "Failed bale detected and routed to Failed_Bales",
    });
    if (failedClass.uncertain) {
      logValidation(result, msg, failed, {
        severity: "WARNING",
        issueType: "FAILED_BALE_UNCERTAIN",
        sheetTargetAttempted: "Failed_Bales",
        problemDescription: "Failed-bale detection uncertain (admin/conflicting/low-structure text)",
      });
    }
    if (failedClass.conflictingSignals) {
      logValidation(result, msg, failed, {
        severity: "WARNING",
        issueType: "FAILED_BALE_CONFLICTING_SIGNALS",
        sheetTargetAttempted: "Failed_Bales",
        problemDescription: "Message contains both failure and production signals",
      });
    }
    validateRecord(failed, "Failed_Bales", result, msg, seenProductionBales);
    return true;
  }

  if (subtype === "crca" || subtype === "test") {
    const crca = {
      ...rowBase,
      baleTestCode: parseCrcaCode(normalized) || rowBase.baleNumber,
      testType: rowBase.productionType || "Test",
      recordType: subtype === "test" ? "TEST" : "CR_CA",
    };
    result.crcaRecords.push(crca);

    if (!crca.baleTestCode) {
      logValidation(result, msg, crca, {
        severity: "WARNING",
        issueType: "MALFORMED_CRCA",
        sheetTargetAttempted: "CR_CA_Tests",
        problemDescription: "Malformed CR/CA/Test code",
      });
    }

    if (subtype === "test") {
      logValidation(result, msg, crca, {
        severity: "INFO",
        issueType: "ROW_EMITTED_FALLBACK",
        sheetTargetAttempted: "CR_CA_Tests",
        problemDescription: "Explicit test bale rerouted from standard production",
      });
    }

    validateRecord(crca, "CR_CA_Tests", result, msg, seenProductionBales);
    return true;
  }

  const production = { ...rowBase, recordType: "STANDARD", productionType: rowBase.productionType || "Production" };
  const keep = validateRecord(production, "Bales_Production", result, msg, seenProductionBales);
  if (keep) result.standardRecords.push(production);
  return true;
}

export function parseBalingMessages(text) {
  const messages = splitWhatsAppMessages(text);
  const result = {
    standardRecords: [],
    failedRecords: [],
    scrapRecords: [],
    crcaRecords: [],
    summaryRecords: [],
    validationLog: [],
    ignoredMessages: 0,
    allRecords: [],
  };

  const seenProductionBales = new Set();

  for (const msg of messages) {
    const normalized = normalizeBalingText(msg.body);
    processBalingMessageOldFormat(msg, normalized, result, seenProductionBales);
  }

  result.allRecords = [
    ...result.standardRecords,
    ...result.failedRecords,
    ...result.scrapRecords,
    ...result.crcaRecords,
  ];

  return result;
}
