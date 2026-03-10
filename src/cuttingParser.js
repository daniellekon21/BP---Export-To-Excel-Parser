// ─── Cutting Parser ──────────────────────────────────────────────────────────
//
// Two parsers for the two WhatsApp chat formats:
//   parseCuttingMessagesNew — structured new format (Cutter 1 / Tyre Type: LC / ...)
//   parseCuttingMessages    — freeform old format  (CM1-25 Agri / 08:00-09:00 / ...)
//
// Shared utility functions live in ./cutting/cuttingUtils.js.

import { parseTime, formatTime, dateToStr, splitWhatsAppMessages } from "./helpers.js";
import {
  extractBodyDate,
  normalizeCuttingLine,
  classifyLine,
  mapTyreType,
  mapTyreTypeNew,
  mapTreadTypeNew,
  parseMachineLine,
  makeMachineRow,
  isValidNewType,
  parseSummaryBlocks,
  parseLegacyDailySummaryBlocks,
  parseLegacyCuttingSummaryBlocks,
  inferDailySummaryType,
  resolveUntypedCounts,
} from "./cutting/cuttingUtils.js";

export {
  normalizeCuttingLine,
  classifyLine,
  mapTyreType,
  parseMachineLine,
  makeMachineRow,
} from "./cutting/cuttingUtils.js";
export { flushCutterBlock, mapTyreTypeNew, mapTreadTypeNew } from "./cutting/cuttingUtils.js";

// ─── New-Format Parser ────────────────────────────────────────────────────────
//
// Structured per-hour format:
//   *Time*: 08:00-09:00
//   *Cutter 1*
//   Operator: Jane / Assistant: Bob
//   Tyre Type: LC    Quantity: 45
//   Tread Type: HC   Quantity: 12
//
// Returns { records, summaryRecords, validationLog }

export function parseCuttingMessagesNew(text) {
  const messages = splitWhatsAppMessages(text);
  const records        = [];
  const summaryRecords = [];
  const validationLog  = [];

  for (const msg of messages) {
    const body = msg.body;
    if (body.includes("<Media omitted>")) continue;

    const isDailySummary     = /daily summary/i.test(body);
    const isStructuredSummary = /cutting summary/i.test(body) || (!isDailySummary && /\bsummary\b/i.test(body));
    const isSummary = isStructuredSummary || isDailySummary;
    const isHourly  = !isSummary && /cutter\s+\d/i.test(body);
    if (!isSummary && !isHourly) continue;

    const date = extractBodyDate(body) ?? msg.tsDate;
    if (!date) continue;

    const dateStr = dateToStr(date);
    const series  = `${String(date.month).padStart(2, "0")}/${String(date.year).slice(2)}`;

    // ── Summary messages ────────────────────────────────────────────────────
    if (isSummary) {
      let blocks = parseSummaryBlocks(body);
      if (blocks.length === 0 && isStructuredSummary) blocks = parseLegacyCuttingSummaryBlocks(body);

      if (blocks.length > 0) {
        for (const b of blocks) {
          const hasAnyValue = b.lc !== null || b.hc !== null || b.agri !== null || b.tread_lc !== null || b.tread_hc !== null || b.tread_agri !== null;
          if (!hasAnyValue) {
            validationLog.push({ date: dateStr, time: "", messageType: "Summary", cutter: `CM - ${b.cmNum}`, issue: "Summary block has no parseable tyre/tread values", action: "Summary row skipped" });
            continue;
          }
          summaryRecords.push({ date, series, cmNumber: `CM - ${b.cmNum}`, lc: b.lc, hc: b.hc, agri: b.agri, tread_lc: b.tread_lc, tread_hc: b.tread_hc, tread_agri: b.tread_agri });
        }
        continue;
      }

      if (isDailySummary) {
        for (const b of parseLegacyDailySummaryBlocks(body)) {
          const cmLabel = `CM - ${b.cmNum}`;
          const inferredType = b.type ?? inferDailySummaryType(records, date, cmLabel);
          if (!inferredType) {
            validationLog.push({ date: dateStr, time: "", messageType: "Summary", cutter: cmLabel, issue: `Ambiguous legacy daily summary type in "${b.raw}"`, action: "Summary row skipped" });
            continue;
          }
          summaryRecords.push({ date, series, cmNumber: cmLabel, lc: inferredType === "lc" ? b.qty : null, hc: inferredType === "hc" ? b.qty : null, agri: inferredType === "agri" ? b.qty : null, tread_lc: null, tread_hc: null, tread_agri: null });
        }
      }
      continue;
    }

    // ── Hourly messages ─────────────────────────────────────────────────────
    let startTime = null, finishTime = null;
    const slotMatch = body.match(/\*?time\*?\s*:\s*(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})/i);
    if (slotMatch) { startTime = parseTime(slotMatch[1]); finishTime = parseTime(slotMatch[2]); }
    const timeStr = (startTime && finishTime) ? `${formatTime(startTime)}-${formatTime(finishTime)}` : "";

    const seenCutters     = new Set();
    const producedCutters = new Set();
    let currentBlock = null;
    let lastQtyField = null;
    let isDuplicate  = false;

    function flushBlock() {
      if (!currentBlock || isDuplicate) return;
      const { cmNum, operator, assistant, tyreType, tyreCount, treadType, treadCount } = currentBlock;
      const opStr = [operator, assistant].filter(Boolean).join(" / ");

      if (tyreType !== null && !isValidNewType(tyreType)) {
        validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: `Invalid tyre type "${tyreType}"`, action: "Block skipped — only LC, HC, Agri are valid" });
        return;
      }
      if (treadType !== null && !isValidNewType(treadType)) {
        validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: `Invalid tread type "${treadType}"`, action: "Block skipped — only LC, HC, Agri are valid" });
        return;
      }
      if (tyreType === null) return;
      if (tyreCount === null) {
        validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: "Tyre type present but quantity missing", action: "Block skipped" });
        return;
      }

      const row = makeMachineRow(date, `CM - ${cmNum}`, series, startTime, finishTime, opStr, body);
      const tyreCol = mapTyreTypeNew(tyreType);
      if (tyreCol !== "unknown_type") row[tyreCol] = tyreCount;

      if (treadType === null) {
        validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: "Tread section missing", action: "Partial parse — tyre data written, tread omitted" });
      } else if (treadCount === null) {
        validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: "Tread type present but quantity missing", action: "Tread omitted" });
      } else {
        const treadCol = mapTreadTypeNew(treadType);
        if (treadCol !== "unknown_type") row[treadCol] = treadCount;
      }

      records.push(row);
      producedCutters.add(cmNum);
    }

    for (const rawLine of body.split("\n")) {
      const line = rawLine.trim();
      if (!line) continue;

      const cutterMatch = line.match(/^\*?cutter\s+(\d+)\*?/i);
      if (cutterMatch) {
        flushBlock();
        const cmNum = parseInt(cutterMatch[1], 10);
        isDuplicate = seenCutters.has(cmNum);
        if (isDuplicate) {
          validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: `Duplicate cutter CM - ${cmNum} in same message`, action: "Kept first occurrence, ignored later duplicate" });
        } else {
          seenCutters.add(cmNum);
        }
        currentBlock = { cmNum, operator: "", assistant: "", tyreType: null, tyreCount: null, treadType: null, treadCount: null };
        lastQtyField = null;
        continue;
      }

      if (!currentBlock || isDuplicate) continue;

      const m_op   = line.match(/^operator\s*:\s*(.*)/i);
      const m_ast  = line.match(/^assistant\s*:\s*(.*)/i);
      const m_tyre = line.match(/^t[iy]re\s+type\s*:\s*(.*)/i);
      const m_trd  = line.match(/^tread\s+type\s*:\s*(.*)/i);
      const m_qty  = line.match(/^quantit[yi]e?s?\s*:\s*(\d+)/i);

      if (m_op)   { currentBlock.operator  = m_op[1].trim();  continue; }
      if (m_ast)  { currentBlock.assistant = m_ast[1].trim(); continue; }
      if (m_tyre) { currentBlock.tyreType  = m_tyre[1].trim(); lastQtyField = "tyre";  continue; }
      if (m_trd)  { currentBlock.treadType = m_trd[1].trim();  lastQtyField = "tread"; continue; }
      if (m_qty) {
        const qty = parseInt(m_qty[1], 10);
        if (lastQtyField === "tyre")  currentBlock.tyreCount  = qty;
        if (lastQtyField === "tread") currentBlock.treadCount = qty;
      }
    }
    flushBlock();

    if (producedCutters.size > 0) {
      for (const n of [1, 2, 3]) {
        if (!producedCutters.has(n)) {
          const p = makeMachineRow(date, `CM - ${n}`, series, startTime, finishTime, "", body);
          p._syntheticPlaceholder = true;
          records.push(p);
        }
      }
    }
  }

  return { records, summaryRecords, validationLog };
}

// ─── Old-Format Parser ────────────────────────────────────────────────────────
//
// Freeform per-hour format (one message = one or more time intervals):
//   09/10/2025 / 08:00-09:00 / CM1-25 Agriculture / CM2-18 Radial(LC)
//
// Returns { records, summaryRecords, validationLog }

export function parseCuttingMessages(text) {
  const messages = splitWhatsAppMessages(text);
  const records = [];
  const summaryRecords = [];
  const validationLog = [];

  for (const msg of messages) {
    const body = msg.body;
    if (body.includes("<Media omitted>")) continue;

    const date = extractBodyDate(body) ?? msg.tsDate;
    if (!date) continue;

    const series  = `${String(date.month).padStart(2, "0")}/${String(date.year).slice(2)}`;
    const dateStr = dateToStr(date);
    const isDailySummary     = /daily summary/i.test(body);
    const isStructuredSummary = /cutting summary/i.test(body) || (!isDailySummary && /\bsummary\b/i.test(body));

    if (isStructuredSummary) {
      let blocks = parseSummaryBlocks(body);
      if (blocks.length === 0) blocks = parseLegacyCuttingSummaryBlocks(body);
      for (const b of blocks) {
        const hasAnyValue = b.lc !== null || b.hc !== null || b.agri !== null || b.tread_lc !== null || b.tread_hc !== null || b.tread_agri !== null;
        if (!hasAnyValue && !b._hasMarker) {
          validationLog.push({ date: dateStr, time: "", messageType: "Summary", cutter: `CM - ${b.cmNum}`, issue: "Summary block has no parseable tyre/tread values", action: "Summary row skipped" });
          continue;
        }
        summaryRecords.push({ date, series, cmNumber: `CM - ${b.cmNum}`, lc: b.lc, hc: b.hc, agri: b.agri, tread_lc: b.tread_lc, tread_hc: b.tread_hc, tread_agri: b.tread_agri });
      }
      continue;
    }

    if (isDailySummary) {
      for (const b of parseLegacyDailySummaryBlocks(body)) {
        const cmLabel = `CM - ${b.cmNum}`;
        const inferredType = b.type ?? inferDailySummaryType(records, date, cmLabel);
        if (!inferredType) {
          validationLog.push({ date: dateStr, time: "", messageType: "Summary", cutter: cmLabel, issue: `Ambiguous legacy daily summary type in "${b.raw}"`, action: "Summary row skipped" });
          continue;
        }
        summaryRecords.push({ date, series, cmNumber: cmLabel, lc: inferredType === "lc" ? b.qty : null, hc: inferredType === "hc" ? b.qty : null, agri: inferredType === "agri" ? b.qty : null, tread_lc: null, tread_hc: null, tread_agri: null });
      }
      continue;
    }

    if (!/CM\s*\d|Machine\s+\d/i.test(body)) continue;

    // Split body into interval blocks at each time-range line
    const intervals = [];
    let cur = { startTime: null, finishTime: null, lines: [] };
    for (const rawLine of body.split("\n")) {
      const timeMatch = rawLine.match(/(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})/);
      if (timeMatch) {
        intervals.push(cur);
        cur = { startTime: parseTime(timeMatch[1]), finishTime: parseTime(timeMatch[2]), lines: [] };
      } else {
        cur.lines.push(rawLine);
      }
    }
    intervals.push(cur);

    for (const interval of intervals) {
      if (interval.lines.length === 0) continue;

      const { startTime, finishTime } = interval;
      const seenCMs    = new Set();
      const machineRows = new Map();
      let hasIntervalProduction = false;
      let intervalHintColumn = null;
      let m;

      const getRow = (cmNum) => {
        if (!machineRows.has(cmNum)) machineRows.set(cmNum, makeMachineRow(date, `CM - ${cmNum}`, series, startTime, finishTime, "", body));
        return machineRows.get(cmNum);
      };

      const intervalBody = interval.lines.join("\n");

      // Phase 1: body-level handlers for parenthesised formats
      // Format D: CM1-(HC)=17 / (LC)=12  (dual-type, may span two lines)
      const fmtD = /CM\s*(\d)\s*-\s*\(HC\)\s*=?\s*(\d+)[\s\S]*?\(LC\)\s*=?\s*(\d+)/gi;
      while ((m = fmtD.exec(intervalBody)) !== null) {
        const cmNum = parseInt(m[1], 10);
        if (seenCMs.has(cmNum)) continue;
        seenCMs.add(cmNum);
        const hc = parseInt(m[2], 10), lc = parseInt(m[3], 10);
        const row = getRow(cmNum);
        if (!isNaN(hc) && hc > 0) row.heavy_commercial_t = hc;
        if (!isNaN(lc) && lc > 0) row.tread_lc = lc;
        if ((!isNaN(hc) && hc > 0) || (!isNaN(lc) && lc > 0)) hasIntervalProduction = true;
      }

      // Single-(HC) lines
      const fmtHC = /CM\s*(\d)\s*-\s*\(HC\)\s*=?\s*(\d+)/gi;
      while ((m = fmtHC.exec(intervalBody)) !== null) {
        const cmNum = parseInt(m[1], 10);
        if (seenCMs.has(cmNum)) continue;
        seenCMs.add(cmNum);
        const count = parseInt(m[2], 10);
        const row = getRow(cmNum);
        if (!isNaN(count)) row.heavy_commercial_t = count;
        if (!isNaN(count) && count > 0) hasIntervalProduction = true;
      }

      // Single-(LC) lines → RADIALS Light Commercial T
      const fmtLC = /CM\s*(\d)\s*-\s*\(LC\)\s*=?\s*(\d+)/gi;
      while ((m = fmtLC.exec(intervalBody)) !== null) {
        const cmNum = parseInt(m[1], 10);
        if (seenCMs.has(cmNum)) continue;
        seenCMs.add(cmNum);
        const count = parseInt(m[2], 10);
        const row = getRow(cmNum);
        if (!isNaN(count)) row.tread_lc = count;
        if (!isNaN(count) && count > 0) hasIntervalProduction = true;
      }

      // Phase 2: line-by-line for all other formats
      const pendingRecords = [];
      for (const rawLine of interval.lines) {
        if (classifyLine(rawLine) !== "machine_line") continue;
        const normalized = normalizeCuttingLine(rawLine);
        if (!intervalHintColumn) {
          const hinted = mapTyreType(normalized);
          if (hinted && hinted !== "unknown_type" && hinted !== "treads") intervalHintColumn = hinted;
        }
        const parsed = parseMachineLine(normalized);
        if (parsed.length === 0) continue;
        const cmNum = parsed[0].cmNum;
        if (seenCMs.has(cmNum)) continue;
        seenCMs.add(cmNum);
        for (const p of parsed) pendingRecords.push(p);
      }

      // Infer type for bare-count lines from sibling CMs in the same interval
      const knownColumns = pendingRecords.filter(p => p.column !== null && p.column !== "unknown_type" && p.column !== "").map(p => p.column);
      const inferredColumn = [...new Set(knownColumns)].length === 1 ? [...new Set(knownColumns)][0] : null;

      for (const p of pendingRecords) {
        if (p.count === null && !p.isStatus) continue;
        let col = p.column;
        if ((col === null || col === "unknown_type") && p.count !== null && inferredColumn) col = inferredColumn;
        if ((col === null || col === "unknown_type") && p.count !== null && intervalHintColumn) col = intervalHintColumn;
        const row = getRow(p.cmNum);
        if (col && col !== "unknown_type" && col !== "" && p.count !== null) {
          row[col] = p.count;
          if (p.count > 0) hasIntervalProduction = true;
        } else if (p.count !== null) {
          row._untypedCount = (row._untypedCount ?? 0) + p.count;
          if (!row._hintColumn && intervalHintColumn) row._hintColumn = intervalHintColumn;
          if (p.count > 0) hasIntervalProduction = true;
        }
      }

      if (!hasIntervalProduction) continue;

      for (const row of machineRows.values()) records.push(row);

      for (const n of [1, 2, 3]) {
        if (!machineRows.has(n)) {
          const p = makeMachineRow(date, `CM - ${n}`, series, startTime, finishTime, "", body);
          p._syntheticPlaceholder = true;
          records.push(p);
        }
      }
    }
  }

  resolveUntypedCounts(records);
  return { records, summaryRecords, validationLog };
}
