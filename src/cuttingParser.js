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
  parseStandaloneTreadLine,
  makeMachineRow,
  isValidNewType,
  parseSummaryBlocks,
  parseLegacyDailySummaryBlocks,
  parseLegacyCuttingSummaryBlocks,
  inferDailySummaryType,
  resolveUntypedCounts,
  extractSummaryTotalTreads,
} from "./cutting/cuttingUtils.js";

export {
  normalizeCuttingLine,
  classifyLine,
  mapTyreType,
  parseMachineLine,
  parseStandaloneTreadLine,
  makeMachineRow,
} from "./cutting/cuttingUtils.js";
export { flushCutterBlock, mapTyreTypeNew, mapTreadTypeNew } from "./cutting/cuttingUtils.js";

function cmSortNumber(row) {
  const m = String(row?.cmNumber ?? "").match(/(\d+)/);
  if (!m) return Number.MAX_SAFE_INTEGER;
  const n = parseInt(m[1], 10);
  return Number.isNaN(n) ? Number.MAX_SAFE_INTEGER : n;
}

function sortRowsByCm(rows) {
  rows.sort((a, b) => {
    const cmCmp = cmSortNumber(a) - cmSortNumber(b);
    if (cmCmp !== 0) return cmCmp;
    const ai = Number.isInteger(a?._rowOrderInText) ? a._rowOrderInText : 0;
    const bi = Number.isInteger(b?._rowOrderInText) ? b._rowOrderInText : 0;
    return ai - bi;
  });
}

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
  let parseOrder = 0;

  function pushRecord(row) {
    row._parseOrder = parseOrder++;
    records.push(row);
  }

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
      const totalTreadsQty = extractSummaryTotalTreads(body);
      let blocks = parseSummaryBlocks(body);
      if (blocks.length === 0 && isStructuredSummary) blocks = parseLegacyCuttingSummaryBlocks(body);

      if (blocks.length > 0) {
        for (const b of blocks) {
          const cmLabel = `CM - ${b.cmNum}`;
          // Infer type for untyped quantities (bare "Tyres" / "Radials" or _untypedQty)
          const hasUntypedTotal = b.radialsTotal != null && b.radialsLC == null && b.radialsHC == null && b.radialsAgri == null;
          const needsInference = hasUntypedTotal || b._untypedQty != null;
          const inferredType = needsInference ? inferDailySummaryType(records, date, cmLabel) : null;
          const untypedQty = b.radialsTotal ?? b._untypedQty ?? null;

          let totalLC = b.radialsLC ?? (inferredType === "radialsLC" ? untypedQty : null);
          let totalHC = b.radialsHC ?? (inferredType === "radialsHC" ? untypedQty : null);
          const baseAgri = b.radialsAgri ?? (inferredType === "radialsAgri" ? untypedQty : null);
          // In daily summary, treads count as agri — sum them together
          const treadAgri = b.radialsAgriTreads ?? 0;
          const totalAgri = baseAgri != null ? baseAgri + treadAgri : (treadAgri > 0 ? treadAgri : null);
          // If bare "Radials" label wasn't inferred to a specific type, store in Radials Total column
          let totalRadials = (totalLC !== null && totalHC !== null) ? ((totalLC ?? 0) + (totalHC ?? 0)) : null;
          if (totalRadials === null && hasUntypedTotal && inferredType === null && b.radialsTotal != null) {
            totalRadials = b.radialsTotal;
          }
          const totalAgriTreads = b.radialsAgriTreads ?? null;
          // Only _untypedQty (bare number, no type label) → Unknown column; labeled "Radials" goes to Radials Total
          const _unresolved = b._untypedQty != null && b._untypedQty > 0 && inferredType === null;
          const totalUnknown = _unresolved ? b._untypedQty : null;
          const hasAnyValue = totalLC !== null || totalHC !== null || totalAgri !== null || totalAgriTreads !== null || totalRadials !== null || totalUnknown !== null;
          if (!hasAnyValue) {
            // N/A or 0 blocks are expected — no need to log
            continue;
          }
          summaryRecords.push({ date, series, cmNumber: cmLabel, totalLC, totalHC, totalRadials, totalAgri, totalAgriTreads, totalUnknown, _unresolved });
        }
        if (totalTreadsQty !== null) summaryRecords.push({ date, series, cmNumber: "Total Treads", totalLC: null, totalHC: null, totalRadials: null, totalAgri: totalTreadsQty, totalAgriTreads: null, totalUnknown: null });
        continue;
      }

      if (isDailySummary) {
        for (const b of parseLegacyDailySummaryBlocks(body)) {
          const cmLabel = `CM - ${b.cmNum}`;
          const inferredType = b.type ?? inferDailySummaryType(records, date, cmLabel);
          // Treads in daily summary count as Agri
          const treadAgri = b.treadQty ?? 0;
          if (!inferredType) {
            summaryRecords.push({
              date, series, cmNumber: cmLabel,
              totalLC: null, totalHC: null, totalRadials: null,
              totalAgri: null, totalAgriTreads: null,
              totalUnknown: (b.qty + treadAgri) || null, _unresolved: (b.qty + treadAgri) > 0,
            });
            continue;
          }
          summaryRecords.push({
            date,
            series,
            cmNumber: cmLabel,
            totalLC: inferredType === "radialsLC" ? b.qty : null,
            totalHC: inferredType === "radialsHC" ? b.qty : null,
            totalRadials: null,
            totalAgri: (inferredType === "radialsAgri" ? b.qty : 0) + treadAgri || null,
            totalAgriTreads: null,
            totalUnknown: null,
          });
        }
      }
      if (totalTreadsQty !== null) summaryRecords.push({ date, series, cmNumber: "Total Treads", totalLC: null, totalHC: null, totalRadials: null, totalAgri: totalTreadsQty, totalAgriTreads: null, totalUnknown: null });
      continue;
    }

    // ── Hourly messages ─────────────────────────────────────────────────────
    let startTime = null, finishTime = null;
    const slotMatch = body.match(/\*?time\*?\s*:\s*(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})/i);
    if (slotMatch) { startTime = parseTime(slotMatch[1]); finishTime = parseTime(slotMatch[2]); }
    const timeStr = (startTime && finishTime) ? `${formatTime(startTime)}-${formatTime(finishTime)}` : "";

    const seenCutters     = new Set();
    const producedCutters = new Set();
    const messageRows     = [];
    let currentBlock = null;
    let lastQtyField = null;
    let rowOrderInText = 0;

    function flushBlock() {
      if (!currentBlock) return;
      const { cmNum, operator, assistant, tyreType, tyreCount, treadType, treadCount } = currentBlock;
      const opStr = [operator, assistant].filter(Boolean).join(" / ");

      if (tyreType !== null && !isValidNewType(tyreType)) {
        validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: `Invalid tyre type "${tyreType}"`, action: "Block skipped — only LC, HC, Agri are valid", rawText: body });
        return;
      }
      if (treadType !== null && !isValidNewType(treadType)) {
        validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: `Invalid tread type "${treadType}"`, action: "Block skipped — only LC, HC, Agri are valid", rawText: body });
        return;
      }
      if (tyreType === null) return;
      if (tyreCount === null) {
        validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: "Tyre type present but quantity missing", action: "Block skipped", rawText: body });
        return;
      }

      const row = makeMachineRow(date, `CM - ${cmNum}`, series, startTime, finishTime, opStr, body);
      const tyreCol = mapTyreTypeNew(tyreType);
      if (tyreCol !== "unknown_type") row[tyreCol] = (row[tyreCol] ?? 0) + tyreCount;

      if (treadType === null) {
        validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: "Tread section missing", action: "Partial parse — tyre data written, tread omitted", rawText: body });
      } else if (treadCount === null) {
        validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: "Tread type present but quantity missing", action: "Tread omitted", rawText: body });
      } else {
        const treadCol = mapTreadTypeNew(treadType);
        if (treadCol !== "unknown_type") row[treadCol] = (row[treadCol] ?? 0) + treadCount;
      }

      row._rowOrderInText = rowOrderInText++;
      messageRows.push(row);
      producedCutters.add(cmNum);
    }

    for (const rawLine of body.split("\n")) {
      const line = rawLine.trim();
      if (!line) continue;

      const cutterMatch = line.match(/^\*?cutter\s+(\d+)\*?/i);
      if (cutterMatch) {
        flushBlock();
        const cmNum = parseInt(cutterMatch[1], 10);
        if (seenCutters.has(cmNum)) {
          validationLog.push({ date: dateStr, time: timeStr, messageType: "Hourly", cutter: `CM - ${cmNum}`, issue: `Duplicate cutter CM - ${cmNum} in same message`, action: "Kept all occurrences", rawText: body });
        }
        seenCutters.add(cmNum);
        currentBlock = { cmNum, operator: "", assistant: "", tyreType: null, tyreCount: null, treadType: null, treadCount: null };
        lastQtyField = null;
        continue;
      }

      if (!currentBlock) continue;

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
          p._rowOrderInText = rowOrderInText++;
          messageRows.push(p);
        }
      }
    }

    sortRowsByCm(messageRows);
    for (const row of messageRows) pushRecord(row);
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
  const unattributedRecords = []; // "CM-13 Agri Treads" style lines with no machine digit
  let parseOrder = 0;

  function pushRecord(row) {
    row._parseOrder = parseOrder++;
    records.push(row);
  }

  for (const msg of messages) {
    const body = msg.body;
    if (body.includes("<Media omitted>")) continue;

    const date = extractBodyDate(body) ?? msg.tsDate;
    if (!date) continue;

    const series  = `${String(date.month).padStart(2, "0")}/${String(date.year).slice(2)}`;
    const isDailySummary     = /daily summary/i.test(body);
    const isStructuredSummary = /cutting summary/i.test(body) || (!isDailySummary && /\bsummary\b/i.test(body));

    if (isStructuredSummary) {
      let blocks = parseSummaryBlocks(body);
      if (blocks.length === 0) blocks = parseLegacyCuttingSummaryBlocks(body);
      for (const b of blocks) {
        const cmLabel = `CM - ${b.cmNum}`;
        const hasUntypedTotal = b.radialsTotal != null && b.radialsLC == null && b.radialsHC == null && b.radialsAgri == null;
        const needsInference = hasUntypedTotal || b._untypedQty != null;
        const inferredType = needsInference ? inferDailySummaryType(records, date, cmLabel) : null;
        const untypedQty = b.radialsTotal ?? b._untypedQty ?? null;

        let totalLC = b.radialsLC ?? (inferredType === "radialsLC" ? untypedQty : null);
        let totalHC = b.radialsHC ?? (inferredType === "radialsHC" ? untypedQty : null);
        const baseAgri = b.radialsAgri ?? (inferredType === "radialsAgri" ? untypedQty : null);
        // In daily summary, treads count as agri — sum them together
        const treadAgri = b.radialsAgriTreads ?? 0;
        const totalAgri = baseAgri != null ? baseAgri + treadAgri : (treadAgri > 0 ? treadAgri : null);
        // If bare "Radials" label wasn't inferred to a specific type, store in Radials Total column
        let totalRadials = (totalLC !== null && totalHC !== null) ? ((totalLC ?? 0) + (totalHC ?? 0)) : null;
        if (totalRadials === null && hasUntypedTotal && inferredType === null && b.radialsTotal != null) {
          totalRadials = b.radialsTotal;
        }
        const totalAgriTreads = b.radialsAgriTreads ?? null;
        // Only _untypedQty (bare number, no type label) → Unknown column; labeled "Radials" goes to Radials Total
        const _unresolved = b._untypedQty != null && b._untypedQty > 0 && inferredType === null;
        const totalUnknown = _unresolved ? b._untypedQty : null;
        const hasAnyValue = totalLC !== null || totalHC !== null || totalAgri !== null || totalAgriTreads !== null || totalRadials !== null || totalUnknown !== null;
        if (!hasAnyValue && !b._hasMarker) {
          // N/A or 0 blocks are expected — no need to log
          continue;
        }
        summaryRecords.push({ date, series, cmNumber: cmLabel, totalLC, totalHC, totalRadials, totalAgri, totalAgriTreads, totalUnknown, _unresolved });
      }
      const totalTreadsQty = extractSummaryTotalTreads(body);
      if (totalTreadsQty !== null) summaryRecords.push({ date, series, cmNumber: "Total Treads", totalLC: null, totalHC: null, totalRadials: null, totalAgri: totalTreadsQty, totalAgriTreads: null, totalUnknown: null });
      continue;
    }

    if (isDailySummary) {
      for (const b of parseLegacyDailySummaryBlocks(body)) {
        const cmLabel = `CM - ${b.cmNum}`;
        const inferredType = b.type ?? inferDailySummaryType(records, date, cmLabel);
        // Treads in daily summary count as Agri
        const treadAgri = b.treadQty ?? 0;
        if (!inferredType) {
          summaryRecords.push({
            date, series, cmNumber: cmLabel,
            totalLC: null, totalHC: null, totalRadials: null,
            totalAgri: null, totalAgriTreads: null,
            totalUnknown: b.qty + treadAgri, _unresolved: true,
          });
          continue;
        }
        summaryRecords.push({
          date,
          series,
          cmNumber: cmLabel,
          totalLC: inferredType === "radialsLC" ? b.qty : null,
          totalHC: inferredType === "radialsHC" ? b.qty : null,
          totalRadials: null,
          totalAgri: (inferredType === "radialsAgri" ? b.qty : 0) + treadAgri || null,
          totalAgriTreads: null,
          totalUnknown: null,
        });
      }
      const totalTreadsQty = extractSummaryTotalTreads(body);
      if (totalTreadsQty !== null) summaryRecords.push({ date, series, cmNumber: "Total Treads", totalLC: null, totalHC: null, totalRadials: null, totalAgri: totalTreadsQty, totalAgriTreads: null, totalUnknown: null });
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
      const seenCMs = new Set();
      const machineRows = [];
      let hasIntervalProduction = false;
      let m;
      let rowOrderInText = 0;
      const intervalTimeLabel = (startTime && finishTime)
        ? `${formatTime(startTime)}-${formatTime(finishTime)}`
        : "";
      const intervalLines = interval.lines.map((l) => l.trim()).filter(Boolean);
      const intervalRawMessage = [dateToStr(date), intervalTimeLabel, ...intervalLines]
        .filter(Boolean)
        .join("\n");

      const createRow = (cmNum) => {
        const row = makeMachineRow(date, `CM - ${cmNum}`, series, startTime, finishTime, "", intervalRawMessage);
        row._rowOrderInText = rowOrderInText++;
        machineRows.push(row);
        seenCMs.add(cmNum);
        return row;
      };

      const intervalBody = interval.lines.join("\n");

      // Phase 1: body-level handlers for parenthesised formats
      // Format D: CM1-(HC)=17 / (LC)=12  (dual-type, may span two lines)
      const fmtDCMs = new Set();
      // [^\n]*? = rest of HC line; (?:\n(?!CM\s*\d)[^\n]*?)? = optional next line only if it doesn't start a new CM entry
      const fmtD = /CM\s*(\d)\s*-\s*\(HC\)\s*=?\s*(\d+)[^\n]*?(?:\n(?!CM\s*\d)[^\n]*?)?\(LC\)\s*=?\s*(\d+)/gi;
      while ((m = fmtD.exec(intervalBody)) !== null) {
        const cmNum = parseInt(m[1], 10);
        const hc = parseInt(m[2], 10), lc = parseInt(m[3], 10);
        const row = createRow(cmNum);
        if (!isNaN(hc) && hc > 0) row.radialsHC = hc;
        if (!isNaN(lc) && lc > 0) row.radialsLC = lc;
        if ((!isNaN(hc) && hc > 0) || (!isNaN(lc) && lc > 0)) hasIntervalProduction = true;
        fmtDCMs.add(cmNum);
      }

      // Single-(HC) lines
      const fmtHC = /CM\s*(\d)\s*-\s*\(HC\)\s*=?\s*(\d+)/gi;
      while ((m = fmtHC.exec(intervalBody)) !== null) {
        const cmNum = parseInt(m[1], 10);
        if (fmtDCMs.has(cmNum)) continue;
        const count = parseInt(m[2], 10);
        const row = createRow(cmNum);
        if (!isNaN(count)) row.radialsHC = count;
        if (!isNaN(count) && count > 0) hasIntervalProduction = true;
      }

      // Single-(LC) lines → RADIALS Light Commercial
      const fmtLC = /CM\s*(\d)\s*-\s*\(LC\)\s*=?\s*(\d+)/gi;
      while ((m = fmtLC.exec(intervalBody)) !== null) {
        const cmNum = parseInt(m[1], 10);
        if (fmtDCMs.has(cmNum)) continue;
        const count = parseInt(m[2], 10);
        const row = createRow(cmNum);
        if (!isNaN(count)) row.radialsLC = count;
        if (!isNaN(count) && count > 0) hasIntervalProduction = true;
      }

      // Phase 2: line-by-line for all other formats.
      // Rule: once we see CMX, every subsequent tyre line belongs to CMX until a new CMY appears.
      const pendingRecords = [];
      let currentCM = null;
      for (const rawLine of interval.lines) {
        const normalized = normalizeCuttingLine(rawLine);

        // Machine line → parse it and update the current CM block
        if (classifyLine(rawLine) === "machine_line") {
          const parsed = parseMachineLine(normalized);
          if (parsed.length === 0) continue;
          pendingRecords.push(parsed);
          currentCM = parsed[0].cmNum;
          continue;
        }

        // Ambiguous CM line: "CM-{count} {type}" with no machine digit (e.g. "CM-13 Agri Treads")
        const ambiguousMatch = normalized.match(/^CM\s*-\s*(\d+)\s+([A-Za-z][A-Za-z. ()]*?)\s*$/i);
        if (ambiguousMatch) {
          const count = parseInt(ambiguousMatch[1], 10);
          const col = mapTyreType(ambiguousMatch[2]);
          if (!isNaN(count) && col !== "unknown_type") {
            unattributedRecords.push({
              date, col, count, rawType: ambiguousMatch[2].trim(),
              startTime, finishTime, series, rawMessage: intervalRawMessage,
            });
            hasIntervalProduction = true;
          }
          continue;
        }

        // Not a machine line and no CM seen yet — skip
        if (currentCM === null) continue;

        // Standalone tread count (bare number on its own line)
        const standaloneTreadCount = parseStandaloneTreadLine(normalized);
        if (standaloneTreadCount !== null) {
          const lastParsedLine = pendingRecords[pendingRecords.length - 1];
          if (lastParsedLine && lastParsedLine[0]?.cmNum === currentCM) {
            lastParsedLine.push({ cmNum: currentCM, column: "radialsAgriTreads", count: standaloneTreadCount });
          } else {
            pendingRecords.push([{ cmNum: currentCM, column: "radialsAgriTreads", count: standaloneTreadCount }]);
          }
          continue;
        }

        // General continuation: extract (type, count) in any format.
        // Covers: TYPE=COUNT, (TYPE)=COUNT, -TYPE=COUNT, .TYPE=COUNT,
        //         =COUNT TYPE, COUNT TYPE — regardless of separators or parens.
        let contRawType = null, contCount = null;
        let mc;
        // TYPE = COUNT with optional leading separator and optional trailing qualifier
        mc = normalized.match(/^[-./=]?\s*\(?([A-Za-z][A-Za-z. ]*?)\)?\s*=\s*(\d+)\s*(\([A-Za-z][A-Za-z. ]*\))?\s*$/i);
        if (mc) { contRawType = mc[3] ? `${mc[1]} ${mc[3]}` : mc[1]; contCount = parseInt(mc[2], 10); }
        // = COUNT TYPE
        if (!mc) {
          mc = normalized.match(/^=\s*(\d+)\s+([A-Za-z][A-Za-z. ()]*?)\s*$/i);
          if (mc) { contRawType = mc[2]; contCount = parseInt(mc[1], 10); }
        }
        // COUNT TYPE (bare)
        if (!mc) {
          mc = normalized.match(/^(\d+)\s+([A-Za-z][A-Za-z. ()]*?)\s*$/i);
          if (mc) { contRawType = mc[2]; contCount = parseInt(mc[1], 10); }
        }
        if (mc && contRawType !== null && !isNaN(contCount)) {
          const col = mapTyreType(contRawType);
          if (col !== "unknown_type") {
            const lastParsedLine = pendingRecords[pendingRecords.length - 1];
            if (lastParsedLine && lastParsedLine[0]?.cmNum === currentCM) {
              lastParsedLine.push({ cmNum: currentCM, column: col, count: contCount });
            } else {
              pendingRecords.push([{ cmNum: currentCM, column: col, count: contCount }]);
            }
          }
        }
      }

      // Build per-CM type inference (same machine only — no cross-machine inference)
      const columnsByCm = new Map();
      for (const parsedLine of pendingRecords) {
        const cm = parsedLine[0].cmNum;
        for (const p of parsedLine) {
          if (p.column && p.column !== "unknown_type" && p.column !== "") {
            if (!columnsByCm.has(cm)) columnsByCm.set(cm, new Set());
            columnsByCm.get(cm).add(p.column);
          }
        }
      }

      // Suppress blank status-only rows for CMs that have real tyre production
      const cmsWithRealData = new Set();
      for (const parsedLine of pendingRecords) {
        if (parsedLine.some(p => p.count != null && p.column && p.column !== "" && p.column !== "unknown_type")) {
          cmsWithRealData.add(parsedLine[0].cmNum);
        }
      }

      for (const parsedLine of pendingRecords) {
        const isPureStatus = parsedLine.every(p => p.isStatus && p.count == null);
        if (isPureStatus && cmsWithRealData.has(parsedLine[0].cmNum)) continue;
        const cm = parsedLine[0].cmNum;
        const row = createRow(cm);
        const cmCols = columnsByCm.get(cm);
        const cmInferred = cmCols && cmCols.size === 1 ? [...cmCols][0] : null;
        for (const p of parsedLine) {
          if (p.count === null && !p.isStatus) continue;
          let col = p.column;
          if ((col === null || col === "unknown_type") && p.count !== null && cmInferred) col = cmInferred;
          if (col && col !== "unknown_type" && col !== "" && p.count !== null) {
            row[col] = (row[col] ?? 0) + p.count;
            if (p.count > 0) hasIntervalProduction = true;
          } else if (p.count !== null) {
            row._untypedCount = (row._untypedCount ?? 0) + p.count;
            if (p.isSideWall) {
              row._sideWallCount = (row._sideWallCount ?? 0) + (p.sideWallCount ?? 0);
            }
            if (p.count > 0) hasIntervalProduction = true;
          }
        }
      }

      if (!hasIntervalProduction) continue;

      const intervalRows = [...machineRows];

      for (const n of [1, 2, 3]) {
        if (!seenCMs.has(n)) {
          const p = makeMachineRow(date, `CM - ${n}`, series, startTime, finishTime, "", intervalRawMessage);
          p._syntheticPlaceholder = true;
          p._rowOrderInText = rowOrderInText++;
          intervalRows.push(p);
        }
      }

      sortRowsByCm(intervalRows);
      for (const row of intervalRows) pushRecord(row);
    }
  }

  // Resolve ambiguous CM lines (e.g. "CM-13 Agri Treads" with no machine digit)
  {
    const dk = (d) => d ? `${d.year}-${String(d.month).padStart(2,"0")}-${String(d.day).padStart(2,"0")}` : "";
    for (const u of unattributedRecords) {
      const dateStr = dk(u.date);
      // Find which CMs have data in u.col for this date (across all intervals)
      const cmsWithCol = new Set();
      for (const r of records) {
        if (dk(r.date) !== dateStr) continue;
        if (r[u.col] != null) {
          const m = String(r.cmNumber ?? "").match(/(\d+)/);
          if (m) cmsWithCol.add(parseInt(m[1], 10));
        }
      }

      if (cmsWithCol.size === 1) {
        // Exactly one CM does this tyre type today → attribute to it
        const cmNum = [...cmsWithCol][0];
        const cmLabel = `CM - ${cmNum}`;
        const existing = records.find(r =>
          dk(r.date) === dateStr && r.cmNumber === cmLabel &&
          r.startTime === u.startTime && r.finishTime === u.finishTime
        );
        if (existing) {
          existing[u.col] = (existing[u.col] ?? 0) + u.count;
        } else {
          const row = makeMachineRow(u.date, cmLabel, u.series, u.startTime, u.finishTime, "", u.rawMessage);
          row[u.col] = u.count;
          records.push(row);
        }
      } else {
        // Can't infer → mark ALL rows of this interval orange, lose the qty, log it
        for (const r of records) {
          if (dk(r.date) === dateStr && r.startTime === u.startTime && r.finishTime === u.finishTime) {
            r._ambiguousLine = true;
          }
        }
        const candidates = cmsWithCol.size === 0 ? "none" : [...cmsWithCol].map(n => `CM${n}`).join(", ");
        validationLog.push({
          date: u.date,
          time: `${formatTime(u.startTime)}-${formatTime(u.finishTime)}`,
          messageType: "hourly",
          cutter: "CM - ?",
          issue: `Ambiguous CM line — could not infer machine for "${u.rawType}" (${u.count} tyres). Candidates: ${candidates}`,
          action: "Rows coloured orange. Quantity lost — assign manually.",
          rawText: u.rawMessage,
        });
      }
    }
  }

  resolveUntypedCounts(records, summaryRecords, validationLog);
  return { records, summaryRecords, validationLog };
}
