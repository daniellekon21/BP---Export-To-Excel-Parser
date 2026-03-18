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
          // Compute totalRadials only when we have typed LC+HC values
          const totalRadials = (totalLC !== null && totalHC !== null) ? ((totalLC ?? 0) + (totalHC ?? 0)) : null;
          const totalAgriTreads = b.radialsAgriTreads ?? null;
          // If inference failed, put the untyped value in "Unknown" column as last resort
          // Don't mark as unresolved if the quantity is 0 — nothing to classify
          const _unresolved = untypedQty > 0 && (hasUntypedTotal || (b._untypedQty != null && inferredType == null));
          const totalUnknown = _unresolved ? untypedQty : null;
          const hasAnyValue = totalLC !== null || totalHC !== null || totalAgri !== null || totalAgriTreads !== null || totalRadials !== null || totalUnknown !== null;
          if (!hasAnyValue) {
            // N/A or 0 blocks are expected — no need to log
            continue;
          }
          summaryRecords.push({ date, series, cmNumber: cmLabel, totalLC, totalHC, totalRadials, totalAgri, totalAgriTreads, totalUnknown, _unresolved });
        }
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
        const totalRadials = (totalLC !== null && totalHC !== null) ? ((totalLC ?? 0) + (totalHC ?? 0)) : null;
        const totalAgriTreads = b.radialsAgriTreads ?? null;
        const _unresolved = untypedQty > 0 && (hasUntypedTotal || (b._untypedQty != null && inferredType == null));
        const totalUnknown = _unresolved ? untypedQty : null;
        const hasAnyValue = totalLC !== null || totalHC !== null || totalAgri !== null || totalAgriTreads !== null || totalRadials !== null || totalUnknown !== null;
        if (!hasAnyValue && !b._hasMarker) {
          // N/A or 0 blocks are expected — no need to log
          continue;
        }
        summaryRecords.push({ date, series, cmNumber: cmLabel, totalLC, totalHC, totalRadials, totalAgri, totalAgriTreads, totalUnknown, _unresolved });
      }
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
      const fmtD = /CM\s*(\d)\s*-\s*\(HC\)\s*=?\s*(\d+)[\s\S]*?\(LC\)\s*=?\s*(\d+)/gi;
      while ((m = fmtD.exec(intervalBody)) !== null) {
        const cmNum = parseInt(m[1], 10);
        const hc = parseInt(m[2], 10), lc = parseInt(m[3], 10);
        const row = createRow(cmNum);
        if (!isNaN(hc) && hc > 0) row.radialsHC = hc;
        if (!isNaN(lc) && lc > 0) row.radialsLC = lc;
        if ((!isNaN(hc) && hc > 0) || (!isNaN(lc) && lc > 0)) hasIntervalProduction = true;
      }

      // Single-(HC) lines
      const fmtHC = /CM\s*(\d)\s*-\s*\(HC\)\s*=?\s*(\d+)/gi;
      while ((m = fmtHC.exec(intervalBody)) !== null) {
        const cmNum = parseInt(m[1], 10);
        const count = parseInt(m[2], 10);
        const row = createRow(cmNum);
        if (!isNaN(count)) row.radialsHC = count;
        if (!isNaN(count) && count > 0) hasIntervalProduction = true;
      }

      // Single-(LC) lines → RADIALS Light Commercial
      const fmtLC = /CM\s*(\d)\s*-\s*\(LC\)\s*=?\s*(\d+)/gi;
      while ((m = fmtLC.exec(intervalBody)) !== null) {
        const cmNum = parseInt(m[1], 10);
        const count = parseInt(m[2], 10);
        const row = createRow(cmNum);
        if (!isNaN(count)) row.radialsLC = count;
        if (!isNaN(count) && count > 0) hasIntervalProduction = true;
      }

      // Phase 2: line-by-line for all other formats
      const pendingRecords = [];
      let lastMachineForStandaloneTreads = null;
      for (const rawLine of interval.lines) {
        const normalized = normalizeCuttingLine(rawLine);
        const standaloneTreadCount = parseStandaloneTreadLine(normalized);
        if (standaloneTreadCount !== null && lastMachineForStandaloneTreads !== null) {
          const lastParsedLine = pendingRecords[pendingRecords.length - 1];
          if (lastParsedLine && lastParsedLine[0]?.cmNum === lastMachineForStandaloneTreads) {
            lastParsedLine.push({ cmNum: lastMachineForStandaloneTreads, column: "radialsAgriTreads", count: standaloneTreadCount });
          } else {
            pendingRecords.push([{ cmNum: lastMachineForStandaloneTreads, column: "radialsAgriTreads", count: standaloneTreadCount }]);
          }
          continue;
        }
        if (classifyLine(rawLine) !== "machine_line") continue;
        const parsed = parseMachineLine(normalized);
        if (parsed.length === 0) continue;
        pendingRecords.push(parsed);
        if (parsed.some((p) => p.count !== null && p.column !== "")) {
          lastMachineForStandaloneTreads = parsed[0].cmNum;
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

      for (const parsedLine of pendingRecords) {
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

  resolveUntypedCounts(records, summaryRecords, validationLog);
  return { records, summaryRecords, validationLog };
}
