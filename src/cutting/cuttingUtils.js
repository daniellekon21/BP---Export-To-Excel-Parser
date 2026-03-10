// ─── Cutting Parser Utilities ─────────────────────────────────────────────────
// Shared helpers used by both the old and new cutting parsers.

import { parseTime, formatTime, dateToStr, splitWhatsAppMessages } from "../helpers.js";
export { parseTime, formatTime, dateToStr, splitWhatsAppMessages };

// ─── Body Date Extractor ──────────────────────────────────────────────────────

// Try "Date - DD/MM/YYYY" first, then bare DD/MM/YYYY on its own line.
// Rejects future dates (e.g. operator typo "2036" instead of "2026").
export function extractBodyDate(body) {
  let m = body.match(/\bDate\s*[-–]\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/i);
  if (!m) m = body.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{4})\s*$/m);
  if (!m) return null;
  const d = { day: parseInt(m[1], 10), month: parseInt(m[2], 10), year: parseInt(m[3], 10) };
  if (new Date(d.year, d.month - 1, d.day) > new Date()) return null;
  return d;
}

// ─── Normalize + Classify ─────────────────────────────────────────────────────

// Fix common operator typos before pattern matching.
export function normalizeCuttingLine(line) {
  let s = line;
  s = s.replace(/[–—−]/g, "-");
  s = s.replace(/\b(?:CM\s+|Machine\s+)(\d)/gi, "CM$1");
  s = s.replace(/(\d)([A-Za-z])/g, "$1 $2");
  s = s.replace(/(\d+)\s+x\s+(?=[A-Za-z])/gi, "$1 ");
  s = s.replace(/[^\S\n]+/g, " ").trim();
  return s;
}

// Classify a single line: time_line | machine_line | summary_line | note_line
export function classifyLine(line) {
  const t = line.trim();
  if (!t) return "note_line";
  if (/(?:cutting\s+)?summary\b/i.test(t)) return "summary_line";
  if (/tread/i.test(t) && !/CM\s*\d|Machine\s+\d/i.test(t)) return "note_line";
  if (/(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})/.test(t)) return "time_line";
  if (/\b(?:CM\s*\d|Machine\s+\d)/i.test(t)) {
    const afterCM = t.replace(/^.*?(?:CM\s*\d|Machine\s+\d)/i, "");
    if (/\band\b/i.test(afterCM) && !/\d/.test(afterCM)) return "note_line";
    return "machine_line";
  }
  return "note_line";
}

// ─── Tyre Type Mappers ────────────────────────────────────────────────────────

// Old-format: raw string → canonical Excel column name.
// Longest match first to avoid "Radial(LC)" landing in the HC column.
export function mapTyreType(raw) {
  const t = raw.trim().toLowerCase().replace(/[()]/g, " ").replace(/\s+/g, " ").trim();
  if (t.includes("agri"))                                                    return "agricultural_t";
  if (t.includes("passenger") || t === "pcr")                               return "passenger";
  if (/4.?x.?4/i.test(t))                                                   return "fourx4";
  if (t.includes("radial") && (t.includes("lc") || t.includes("light")))   return "tread_lc";
  if (t.includes("light") || t.includes("lc"))                              return "light_commercial";
  if (t.includes("heavy") || t.includes("hc") || t.includes("truck") || t.includes("radial")) return "heavy_commercial_t";
  if (t.includes("motor"))                                                   return "motorcycle";
  if (t.includes("tread"))                                                   return "treads";
  return "unknown_type";
}

// New-format tyre abbreviation → column key.
export function mapTyreTypeNew(raw) {
  const t = raw.replace(/\*/g, "").trim().toLowerCase();
  if (t === "lc" || t.includes("light")) return "light_commercial";
  if (t === "hc" || t.includes("heavy")) return "heavy_commercial_t";
  if (t.startsWith("agri"))              return "agricultural_t";
  return "unknown_type";
}

// New-format tread abbreviation → column key.
export function mapTreadTypeNew(raw) {
  const t = raw.replace(/\*/g, "").trim().toLowerCase();
  if (t === "lc" || t.includes("light")) return "tread_lc";
  if (t === "hc" || t.includes("heavy")) return "tread_hc";
  if (t.startsWith("agri"))              return "tread_agri";
  return "unknown_type";
}

// ─── Parse Machine Line ───────────────────────────────────────────────────────

// Parse one normalized machine line into [{cmNum, column, count}].
// Format D (HC+LC dual-type) is handled at body level before this is called.
export function parseMachineLine(line) {
  let m;

  // Strip trailing status commentary after a parsed count
  line = line.replace(/(CM\s*\d\s*[-=]\s*\d+)\s*-\s*(?:stopped|idle|waiting|offloading|breakdown|cleaning|paused|maintenance|no\s+production|not\s+(?:in\s+use|cutting|working))\b.*$/i, "$1");

  // Compound: "CM3-6 Agri + 11 treads"
  m = line.match(/CM\s*(\d)\s*[-=]\s*(\d+)\s+(.+?)\s*\+\s*(\d+)\s+([A-Za-z]+)/i);
  if (m) {
    const cmNum = parseInt(m[1], 10);
    const results = [];
    const c1 = parseInt(m[2], 10); if (!isNaN(c1)) results.push({ cmNum, column: mapTyreType(m[3]), count: c1 });
    const c2 = parseInt(m[4], 10); if (!isNaN(c2)) results.push({ cmNum, column: mapTyreType(m[5]), count: c2 });
    if (results.length) return results;
  }

  // Tread Cut: "CM2-14 (Agri) - Tread Cut 21"
  m = line.match(/CM\s*(\d)\s*[-=]\s*(\d+)\s*\(?([A-Za-z]+)\)?\s*-\s*Tread\s*Cut\s*(\d+)/i);
  if (m) {
    return [
      { cmNum: parseInt(m[1], 10), column: mapTyreType(m[3]), count: parseInt(m[2], 10) },
      { cmNum: parseInt(m[1], 10), column: "treads",           count: parseInt(m[4], 10) },
    ];
  }

  // Implicit compound: "CM2-20 Agri 18 Treads"
  m = line.match(/CM\s*(\d)\s*[-=]\s*(\d+)\s+([A-Za-z][A-Za-z.]*)\s+(\d+)\s+([A-Za-z]+)/i);
  if (m) {
    const cmNum = parseInt(m[1], 10);
    const results = [];
    const c1 = parseInt(m[2], 10); if (!isNaN(c1)) results.push({ cmNum, column: mapTyreType(m[3]), count: c1 });
    const c2 = parseInt(m[4], 10); if (!isNaN(c2)) results.push({ cmNum, column: mapTyreType(m[5]), count: c2 });
    if (results.length) return results;
  }

  // Format G: CM1 TypeName - Count
  m = line.match(/CM\s*(\d)\s+([A-Za-z][A-Za-z ]*?)\s+-\s*(\d+)\s*$/i);
  if (m) {
    const count = parseInt(m[3], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[2]), count }];
  }

  // Format A: CM1 = TypeName - Count
  m = line.match(/CM\s*(\d)\s*=\s*([A-Za-z][A-Za-z ]*?)\s*-\s*(\d+|N\/?A)\b/i);
  if (m) {
    const countStr = m[3].trim();
    const count = /n\/?a/i.test(countStr) ? null : parseInt(countStr, 10);
    if (count === null || !isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[2]), count }];
  }

  // NA: CM1 = N/A / CM1-not in use / etc.
  m = line.match(/CM\s*(\d)\s*[-=\s]+(N\/?A|not\s+in\s+use|breakdown|paused|off|offloading|beltrim|maintenance)\b/i);
  if (m) return [{ cmNum: parseInt(m[1], 10), column: "", count: null, isStatus: true }];

  // Format E: CM1 - TypeName = Count
  m = line.match(/CM\s*(\d)\s*-\s*([A-Za-z][A-Za-z ]*?)\s*=\s*(\d+)/i);
  if (m) {
    const count = parseInt(m[3], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[2]), count }];
  }

  // Format C: CM1 - Count TypeName (handles optional parens + LC qualifier)
  m = line.match(/CM\s*(\d)\s*-\s*(\d+)\s*[-\s]+\(?([A-Za-z][A-Za-z\s()]*?)\s*$/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[3]), count }];
  }

  // Format B: CM1 = Count TypeName
  m = line.match(/CM\s*(\d)\s*=\s*(\d+)\s+([A-Za-z][A-Za-z\s()]*?)\s*$/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[3]), count }];
  }

  // Format F: CM1 Count TypeName (no separator)
  m = line.match(/^CM\s*(\d)\s+(\d+)\s+([A-Za-z][A-Za-z\s()]*?)\s*$/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[3]), count }];
  }

  // Format Z: CM1 = Count (no tyre type)
  m = line.match(/^CM\s*(\d)\s*[-=\s]\s*(\d+)\s*$/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: null, count }];
  }

  return [];
}

// ─── Machine Row Factory ──────────────────────────────────────────────────────

// Create a blank machine row — one per machine per timeframe.
export function makeMachineRow(date, cmNumber, series, startTime, finishTime, operator = "", rawMessage = "") {
  return {
    date, cmNumber, series, startTime, finishTime, operator, rawMessage,
    passenger: null, fourx4: null, motorcycle: null,
    light_commercial: null,   // col 9  — LC tyre (PCR)
    heavy_commercial_t: null, // col 11 — HC tyre (RADIALS)
    agricultural_t: null,     // col 12 — Agri tyre (NYLONS)
    tread_lc: null,           // col 10 — LC tread (RADIALS)
    tread_hc: null,           // col 13 — HC tread (NYLONS)
    tread_agri: null,         // col 12 — Agri tread (NYLONS, overrides agricultural_t)
    treads: null,             // col 15 — old-format tread cut total
  };
}

// ─── New-Format Flush Helper ─────────────────────────────────────────────────

// Flush a completed cutter block into records (for external/test use).
export function flushCutterBlock(block, meta, records) {
  const { date, series, startTime, finishTime } = meta;
  const { cmNum, operator, assistant, tyreType, tyreCount, treadType, treadCount } = block;
  const opStr = [operator, assistant].filter(Boolean).join(" / ");
  if (tyreType !== null && tyreCount !== null) {
    records.push({ date, cmNumber: `CM - ${cmNum}`, series, startTime, finishTime, operator: opStr, column: mapTyreTypeNew(tyreType), count: tyreCount });
  }
  if (treadType !== null && treadCount !== null) {
    records.push({ date, cmNumber: `CM - ${cmNum}`, series, startTime, finishTime, operator: "", column: mapTreadTypeNew(treadType), count: treadCount });
  }
}

// ─── New-Format Type Validation ───────────────────────────────────────────────

export function isValidNewType(raw) {
  const t = raw.replace(/\*/g, "").trim().toLowerCase();
  return t === "lc" || t === "hc" || t === "agri";
}

// ─── Summary Block Parsers ────────────────────────────────────────────────────

// Parse cutter blocks from a Cutting Summary message body.
// Returns [{cmNum, lc, hc, agri, tread_lc, tread_hc, tread_agri}]
export function parseSummaryBlocks(body) {
  const blocks = [];
  let cur = null;
  let section = null;
  let lastTypeField = null;

  function applyTypeQty(rawType, qty, forceTread = false) {
    const t = rawType.trim().toLowerCase();
    if (forceTread || /^tread\s+lc$/.test(t)) { cur.tread_lc = qty; return; }
    if (forceTread || /^tread\s+hc$/.test(t)) { cur.tread_hc = qty; return; }
    if (forceTread || /^tread\s+agri/.test(t)) { cur.tread_agri = qty; return; }
    const col = mapTyreType(rawType);
    if (col === "light_commercial") cur.lc = qty;
    else if (col === "heavy_commercial_t") cur.hc = qty;
    else if (col === "agricultural_t") cur.agri = qty;
  }

  function parseInlinePayload(payload) {
    const s = payload.trim();
    if (!s || /n\/?a|not\s+in\s+use|offloading|off\b/i.test(s)) return;
    const pairs = [...s.matchAll(/([A-Za-z][A-Za-z ]+?)\s*[-=:]\s*(\d+)/g)];
    if (pairs.length > 0) { for (const m of pairs) applyTypeQty(m[1], parseInt(m[2], 10)); return; }
    const loose = s.match(/^([A-Za-z][A-Za-z ]+)\s+(\d+)$/);
    if (loose) applyTypeQty(loose[1], parseInt(loose[2], 10));
  }

  for (const rawLine of body.split("\n")) {
    const line = rawLine.trim();
    if (!line || /(?:cutting\s+)?summary\b/i.test(line)) continue;

    const cutterMatch = line.match(/^\*?(?:cutter|cm)\s*[- ]?(\d+)\*?(?:\s*[:=\-]\s*(.*))?$/i);
    if (cutterMatch) {
      if (cur) blocks.push(cur);
      cur = { cmNum: parseInt(cutterMatch[1], 10), lc: null, hc: null, agri: null, tread_lc: null, tread_hc: null, tread_agri: null };
      section = null; lastTypeField = null;
      if (cutterMatch[2]) parseInlinePayload(cutterMatch[2]);
      continue;
    }
    if (!cur) continue;

    if (/^total\s+tyre/i.test(line)) { section = "tyre"; continue; }
    if (/^total\s+tread/i.test(line)) { section = "tread"; continue; }

    const inlineTread = line.match(/^tread\s+(lc|hc|agri)\s*:\s*(\d+)/i);
    if (inlineTread) {
      const t = inlineTread[1].toLowerCase(), qty = parseInt(inlineTread[2], 10);
      if (t === "lc") cur.tread_lc = qty; else if (t === "hc") cur.tread_hc = qty; else cur.tread_agri = qty;
      continue;
    }

    const typeQty = line.match(/^(lc|hc|agri)\s*[:\-–]\s*(\d+)/i);
    if (typeQty) {
      const t = typeQty[1].toLowerCase(), qty = parseInt(typeQty[2], 10);
      if (section === "tread") {
        if (t === "lc") cur.tread_lc = qty; else if (t === "hc") cur.tread_hc = qty; else cur.tread_agri = qty;
      } else {
        if (t === "lc") cur.lc = qty; else if (t === "hc") cur.hc = qty; else cur.agri = qty;
      }
      continue;
    }

    parseInlinePayload(line);

    const m_tyre = line.match(/^t[iy]re\s+type\s*:\s*(.*)/i);
    const m_trd  = line.match(/^tread\s+type\s*:\s*(.*)/i);
    const m_qty  = line.match(/^quantit[yi]e?s?\s*:\s*(\d+)/i);
    if (m_tyre) { cur._tyreType  = m_tyre[1].trim().toLowerCase(); lastTypeField = "tyre"; continue; }
    if (m_trd)  { cur._treadType = m_trd[1].trim().toLowerCase();  lastTypeField = "tread"; continue; }
    if (m_qty) {
      const qty = parseInt(m_qty[1], 10);
      if (lastTypeField === "tyre" && cur._tyreType) {
        const t = cur._tyreType;
        if (t === "lc") cur.lc = qty; else if (t === "hc") cur.hc = qty; else if (t === "agri") cur.agri = qty;
      } else if (lastTypeField === "tread" && cur._treadType) {
        const t = cur._treadType;
        if (t === "lc") cur.tread_lc = qty; else if (t === "hc") cur.tread_hc = qty; else if (t === "agri") cur.tread_agri = qty;
      }
    }
  }
  if (cur) blocks.push(cur);
  return blocks.map(({ _tyreType, _treadType, ...rest }) => rest);
}

export function parseLegacyDailySummaryBlocks(body) {
  const blocks = [];
  function parseCm(raw) {
    const t = raw.trim().toLowerCase();
    if (t === "1" || t === "one") return 1;
    if (t === "2" || t === "two") return 2;
    if (t === "3" || t === "three") return 3;
    return null;
  }
  function parseType(raw) {
    const t = raw.toLowerCase();
    if (t.includes("light") || /\blc\b/.test(t)) return "lc";
    if (t.includes("truck") || t.includes("heavy") || /\bhc\b/.test(t)) return "hc";
    if (t.includes("agri")) return "agri";
    return null;
  }
  for (const rawLine of body.split("\n")) {
    const line = rawLine.trim();
    if (!line || /daily summary/i.test(line) || /cutting machines?/i.test(line) || /^total\b/i.test(line)) continue;
    const m = line.match(/^machine\s+(one|two|three|\d)\s*[-:]\s*(.+)$/i);
    if (!m) continue;
    const cmNum = parseCm(m[1]);
    if (!cmNum) continue;
    const payload = m[2].trim();
    const qtyMatch = payload.match(/(\d+)/);
    if (!qtyMatch) continue;
    const qty = parseInt(qtyMatch[1], 10);
    if (!isNaN(qty)) blocks.push({ cmNum, qty, type: parseType(payload), raw: line });
  }
  return blocks;
}

export function parseLegacyCuttingSummaryBlocks(body) {
  const byCm = new Map();
  let currentCm = null;

  function ensure(cmNum) {
    if (!byCm.has(cmNum)) byCm.set(cmNum, { cmNum, lc: null, hc: null, agri: null, tread_lc: null, tread_hc: null, tread_agri: null, _hasMarker: true });
    return byCm.get(cmNum);
  }
  function applyTypeQty(block, rawType, qty) {
    const col = mapTyreType(rawType);
    if (col === "light_commercial") block.lc = qty;
    else if (col === "heavy_commercial_t") block.hc = qty;
    else if (col === "agricultural_t") block.agri = qty;
    else if (/tread\s*lc/i.test(rawType)) block.tread_lc = qty;
    else if (/tread\s*hc/i.test(rawType)) block.tread_hc = qty;
    else if (/tread\s*agri/i.test(rawType)) block.tread_agri = qty;
  }
  function parseSegment(segment, block) {
    const s = segment.trim();
    if (!s || /n\/?a|not\s+in\s+use|offloading|off\b/i.test(s) || /^total\b/i.test(s)) return;
    const explicit = [...s.matchAll(/([A-Za-z][A-Za-z ]+?)\s*[-=:]\s*(\d+)/g)];
    if (explicit.length > 0) { for (const m of explicit) applyTypeQty(block, m[1].trim(), parseInt(m[2], 10)); return; }
    const loose = s.match(/^([A-Za-z][A-Za-z ]+)\s+(\d+)$/);
    if (loose) applyTypeQty(block, loose[1].trim(), parseInt(loose[2], 10));
  }
  for (const rawLine of body.split("\n")) {
    const line = rawLine.trim();
    if (!line || /(?:cutting\s+)?summary\b/i.test(line) || /^(\d{1,2}\/\d{1,2}\/\d{4}|date\s*[-:])/i.test(line) || /^total\b/i.test(line)) continue;
    const cmHeader = line.match(/^CM\s*[- ]?(\d)\s*=\s*(.*)$/i);
    if (cmHeader) {
      currentCm = parseInt(cmHeader[1], 10);
      parseSegment(cmHeader[2], ensure(currentCm));
      continue;
    }
    if (currentCm) parseSegment(line, ensure(currentCm));
  }
  return [...byCm.values()];
}

export function inferDailySummaryType(records, date, cmNumberLabel) {
  const dateKey = `${date.year}-${date.month}-${date.day}`;
  const cols = new Set();
  for (const r of records) {
    if (!r.date || `${r.date.year}-${r.date.month}-${r.date.day}` !== dateKey || r.cmNumber !== cmNumberLabel) continue;
    if (r.light_commercial != null) cols.add("lc");
    if (r.heavy_commercial_t != null) cols.add("hc");
    if (r.agricultural_t != null) cols.add("agri");
  }
  return cols.size === 1 ? [...cols][0] : null;
}

// ─── Old-Format Untyped Count Resolution ──────────────────────────────────────

export function resolveUntypedCounts(records) {
  const typesByDay = new Map();
  for (const r of records) {
    const d = r?.date ? `${r.date.year}-${String(r.date.month).padStart(2,"0")}-${String(r.date.day).padStart(2,"0")}` : "";
    if (!d) continue;
    if (!typesByDay.has(d)) typesByDay.set(d, new Set());
    for (const col of ["light_commercial","heavy_commercial_t","agricultural_t"]) {
      if (r[col] != null) typesByDay.get(d).add(col);
    }
  }
  for (const r of records) {
    if (r._untypedCount == null) continue;
    const d = r?.date ? `${r.date.year}-${String(r.date.month).padStart(2,"0")}-${String(r.date.day).padStart(2,"0")}` : "";
    let target = r._hintColumn ?? null;
    if (!target) { const dayTypes = typesByDay.get(d); if (dayTypes?.size === 1) target = [...dayTypes][0]; }
    if (target) r[target] = (r[target] ?? 0) + r._untypedCount;
    delete r._untypedCount;
    delete r._hintColumn;
  }
}
