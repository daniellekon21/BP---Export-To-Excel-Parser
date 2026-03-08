// ─── Cutting Parser Pipeline ───────────────────────────────────────────────────
//
// Pipeline stages:
//   1. splitWhatsAppMessages  — raw text → [{sender, body, tsDate}]
//   2. normalizeCuttingLine   — fix spacing/punctuation typos per line
//   3. classifyLine           — identify time_line / machine_line / note_line
//   4. mapTyreType            — raw type string → canonical column name
//   5. parseMachineLine       — normalized machine line → [{cmNum, column, count}]
//   6. parseCuttingMessages   — orchestrates all stages, outputs structured records

import { parseTime, formatTime, dateToStr, splitWhatsAppMessages } from "./helpers.js";

// ─── Body Date Extractor ──────────────────────────────────────────────────────

/**
 * Try to extract an explicit date from the message body (e.g. "Date - 08/12/2025").
 * Format: DD/MM/YYYY (South African convention used by operators).
 * Returns a {day, month, year} object, or null if not found.
 */
function extractBodyDate(body) {
  const m = body.match(/\bDate\s*[-–]\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/i);
  if (!m) return null;
  return {
    day: parseInt(m[1], 10),
    month: parseInt(m[2], 10),
    year: parseInt(m[3], 10),
  };
}

// ─── Stage 2 – Normalize ──────────────────────────────────────────────────────

/**
 * Normalize a single line before pattern matching.
 *
 * Handles common operator typos:
 *   "CM 1"  / "Machine 1"   → "CM1"
 *   "CM1=29Agri"            → "CM1=29 Agri"   (digit run glued to letters)
 *   "CM3=23HC"              → "CM3=23 HC"
 *   unicode/curly dashes    → plain ASCII hyphen
 */
export function normalizeCuttingLine(line) {
  let s = line;
  // Unicode dash variants → plain hyphen
  s = s.replace(/[–—−]/g, "-");
  // "CM 1" / "Machine 1" → "CM1"
  s = s.replace(/\b(?:CM\s+|Machine\s+)(\d)/gi, "CM$1");
  // Insert space between a digit run and the first letter (e.g. "29Agri" → "29 Agri")
  s = s.replace(/(\d)([A-Za-z])/g, "$1 $2");
  // Remove standalone "x" used as a multiplier suffix (e.g. "18 x LC" → "18 LC")
  // but preserve "4 x 4" where x is followed by a digit.
  s = s.replace(/(\d+)\s+x\s+(?=[A-Za-z])/gi, "$1 ");
  // Collapse multiple spaces (preserve newlines)
  s = s.replace(/[^\S\n]+/g, " ").trim();
  return s;
}

// ─── Stage 3 – Classify ──────────────────────────────────────────────────────

/**
 * Classify a single line so we know whether to parse it.
 *
 *   time_line    – "08:00-09:00"
 *   machine_line – starts with CM1, CM 1, Machine 1 …
 *   summary_line – "Cutting Summary"
 *   note_line    – tread lines, staffing assignments, date stamps, blanks (ignored)
 *
 * Tread lines: any line mentioning "tread" is skipped, allowing
 * mixed cutting+tread messages to yield only the cutting rows.
 *
 * Staffing lines: "CM1-Bhekinkosi and Lindani" — matches the CM
 * pattern but contains "and" with no digits after the CM ref → note_line.
 */
export function classifyLine(line) {
  const t = line.trim();
  if (!t) return "note_line";
  if (/cutting summary/i.test(t)) return "summary_line";
  // Only filter tread-related lines that have no machine reference
  if (/tread/i.test(t) && !/CM\s*\d|Machine\s+\d/i.test(t)) return "note_line";
  if (/(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})/.test(t)) return "time_line";
  if (/\b(?:CM\s*\d|Machine\s+\d)/i.test(t)) {
    // Staffing lines contain "and" but no digits after the CM reference
    const afterCM = t.replace(/^.*?(?:CM\s*\d|Machine\s+\d)/i, "");
    if (/\band\b/i.test(afterCM) && !/\d/.test(afterCM)) return "note_line";
    return "machine_line";
  }
  return "note_line";
}

// ─── Stage 4 – Type Mapping ──────────────────────────────────────────────────

/**
 * Map a raw tyre-type string to a canonical Excel column name.
 *
 * Returns "unknown_type" for unrecognised strings so that miscategorised data
 * appears blank rather than landing in the wrong column.
 */
export function mapTyreType(raw) {
  // Strip parens/punctuation so "Radial (LC)" → "radial lc", "(Agri)" → "agri"
  const t = raw.trim().toLowerCase().replace(/[()]/g, " ").replace(/\s+/g, " ").trim();
  if (t.includes("agri"))                         return "agricultural_t";
  if (t.includes("passenger") || t === "pcr")     return "passenger";
  if (/4.?x.?4/i.test(t))                         return "fourx4";
  // Check LC before generic "radial" so "radial lc" / "lc radial" → light_commercial
  if (t.includes("light") || t.includes("lc"))    return "light_commercial";
  if (t.includes("heavy") || t.includes("hc") || t.includes("truck") || t.includes("radial")) return "heavy_commercial_t";
  if (t.includes("motor"))                        return "motorcycle";
  if (t.includes("tread"))                        return "treads";
  return "unknown_type";
}

// ─── Stage 5 – Parse Machine Line ────────────────────────────────────────────

/**
 * Parse one normalized machine line into production records.
 *
 * Supported formats (normalization already inserted digit-letter spaces):
 *   A: CM1 = Type - Count      (Jan 2026+, also handles Type - N/A)
 *   B: CM1 = Count Type        (Nov–Dec 2025, space guaranteed after normalization)
 *   C: CM1 - Count Type        (Sept–Oct 2025)
 *   E: CM1 - Type = Count      (Oct 2025 mid-month)
 *   F: CM1 Count Type          (no separator, e.g. "CM2 45 LC")
 *   Z: CM1 = Count / CM1 Count (no tyre type — count-only row, column=null)
 *  NA: CM1 = N/A               (idle; also: not in use / breakdown / paused / off)
 *
 * Format D (HC + LC dual-type across two lines) is handled at body level
 * in parseCuttingMessages before this function is called.
 *
 * Returns [] if the line cannot be parsed.
 */
export function parseMachineLine(line) {
  let m;

  // ── Pre-step: strip trailing status commentary ────────────────────────────
  // "CM1-6 - stopped due to offloading of scrap truck" → "CM1-6"
  // Only strips when the status phrase follows a parsed count (CM + digit +
  // separator + count), so pure-status lines like "CM3-not in use" or
  // "CM2-offloading truck" are left intact for the NA regex below.
  line = line.replace(/(CM\s*\d\s*[-=]\s*\d+)\s*-\s*(?:stopped|idle|waiting|offloading|breakdown|cleaning|paused|maintenance|no\s+production|not\s+(?:in\s+use|cutting|working))\b.*$/i, "$1");

  // ── Compound lines (two categories in one line) ───────────────────────────
  // Explicit + separator: "CM3-6 Agri + 11 treads", "CM2-20 Radial (LC) +28 Agricultural"
  m = line.match(/CM\s*(\d)\s*[-=]\s*(\d+)\s+(.+?)\s*\+\s*(\d+)\s+([A-Za-z]+)/i);
  if (m) {
    const cmNum = parseInt(m[1], 10);
    const c1 = parseInt(m[2], 10), col1 = mapTyreType(m[3]);
    const c2 = parseInt(m[4], 10), col2 = mapTyreType(m[5]);
    const results = [];
    if (!isNaN(c1)) results.push({ cmNum, column: col1, count: c1 });
    if (!isNaN(c2)) results.push({ cmNum, column: col2, count: c2 });
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

  // Implicit compound (space-separated, no +): "CM2-20 Agri 18 Treads", "CM3-11 HC 12 Agri"
  m = line.match(/CM\s*(\d)\s*[-=]\s*(\d+)\s+([A-Za-z][A-Za-z.]*)\s+(\d+)\s+([A-Za-z]+)/i);
  if (m) {
    const cmNum = parseInt(m[1], 10);
    const c1 = parseInt(m[2], 10), col1 = mapTyreType(m[3]);
    const c2 = parseInt(m[4], 10), col2 = mapTyreType(m[5]);
    const results = [];
    if (!isNaN(c1)) results.push({ cmNum, column: col1, count: c1 });
    if (!isNaN(c2)) results.push({ cmNum, column: col2, count: c2 });
    if (results.length) return results;
  }

  // ── Format G: CMN TypeName - Count (space after CM num, type before dash) ─
  // "CM1 Agricultural - 10", "CM3 Truck Radials - 37"
  m = line.match(/CM\s*(\d)\s+([A-Za-z][A-Za-z ]*?)\s+-\s*(\d+)\s*$/i);
  if (m) {
    const count = parseInt(m[3], 10);
    if (!isNaN(count))
      return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[2]), count }];
  }

  // ── Existing single-record formats ─────────────────────────────────────────

  // Format A: CM1 = TypeName - Count  (or TypeName - N/A)
  m = line.match(/CM\s*(\d)\s*=\s*([A-Za-z][A-Za-z ]*?)\s*-\s*(\d+|N\/?A)\b/i);
  if (m) {
    const countStr = m[3].trim();
    const count = /n\/?a/i.test(countStr) ? null : parseInt(countStr, 10);
    if (count === null || !isNaN(count))
      return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[2]), count }];
  }

  // NA: CM1 = N/A  /  CM1-not in use  /  CM1-offloading  /  CM1-beltrim  etc.
  m = line.match(/CM\s*(\d)\s*[-=\s]+(N\/?A|not\s+in\s+use|breakdown|paused|off|offloading|beltrim|maintenance)\b/i);
  if (m) return [{ cmNum: parseInt(m[1], 10), column: "", count: null, isStatus: true }];

  // Format E: CM1 - TypeName = Count
  m = line.match(/CM\s*(\d)\s*-\s*([A-Za-z][A-Za-z ]*?)\s*=\s*(\d+)/i);
  if (m) {
    const count = parseInt(m[3], 10);
    if (!isNaN(count))
      return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[2]), count }];
  }

  // Format C: CM1 - Count TypeName (handles optional parens + double-dash)
  // "CM2-25 Agriculture", "CM2- 21 (Agri)", "CM1-9 - Agri", "Machine 1-45 Trucks"
  m = line.match(/CM\s*(\d)\s*-\s*(\d+)\s*[-\s]+\(?([A-Za-z]+)\)?/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count))
      return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[3]), count }];
  }

  // Format B: CM1 = Count TypeName  (space guaranteed after normalizeCuttingLine)
  m = line.match(/CM\s*(\d)\s*=\s*(\d+)\s+([A-Za-z]+)/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count))
      return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[3]), count }];
  }

  // Format F: CM1 Count TypeName  (no separator, e.g. "CM2 45 LC")
  m = line.match(/^CM\s*(\d)\s+(\d+)\s+([A-Za-z]+)/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count))
      return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[3]), count }];
  }

  // Format Z: CM1 = Count / CM1-Count / CM1 Count — no tyre type
  m = line.match(/^CM\s*(\d)\s*[-=\s]\s*(\d+)\s*$/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count))
      return [{ cmNum: parseInt(m[1], 10), column: null, count }];
  }

  return [];
}

// ─── New-Format Type Mappers ─────────────────────────────────────────────────

/**
 * Map the new format's tyre-type abbreviation to the internal column key.
 *   LC → light_commercial
 *   HC → heavy_commercial_t
 *   Agri → agricultural_t
 */
export function mapTyreTypeNew(raw) {
  const t = raw.replace(/\*/g, "").trim().toLowerCase();
  if (t === "lc" || t.includes("light")) return "light_commercial";
  if (t === "hc" || t.includes("heavy")) return "heavy_commercial_t";
  if (t.startsWith("agri"))              return "agricultural_t";
  return "unknown_type";
}

/**
 * Map the new format's tread-type abbreviation to the internal column key.
 *   LC → tread_lc  (Light Commercial T, RADIALS col 10)
 *   HC → tread_hc  (Heavy Commercial T, NYLONS col 13)
 *   Agri → tread_agri (Agricultural T, NYLONS col 12)
 */
export function mapTreadTypeNew(raw) {
  const t = raw.replace(/\*/g, "").trim().toLowerCase();
  if (t === "lc" || t.includes("light")) return "tread_lc";
  if (t === "hc" || t.includes("heavy")) return "tread_hc";
  if (t.startsWith("agri"))              return "tread_agri";
  return "unknown_type";
}

// ─── New-Format Flush Helper ─────────────────────────────────────────────────

/**
 * Flush a completed cutter block into the records array.
 * Emits:
 *   - one tyre record if tyre_type and tyre_count are present
 *   - one tread record if tread_type and tread_count are present
 */
export function flushCutterBlock(block, meta, records) {
  const { date, series, startTime, finishTime } = meta;
  const { cmNum, operator, assistant, tyreType, tyreCount, treadType, treadCount } = block;
  const opStr = [operator, assistant].filter(Boolean).join(" / ");

  if (tyreType !== null && tyreCount !== null) {
    records.push({
      date, cmNumber: `CM - ${cmNum}`, series, startTime, finishTime,
      operator: opStr,
      column: mapTyreTypeNew(tyreType), count: tyreCount,
    });
  }
  if (treadType !== null && treadCount !== null) {
    records.push({
      date, cmNumber: `CM - ${cmNum}`, series, startTime, finishTime,
      operator: "",
      column: mapTreadTypeNew(treadType), count: treadCount,
    });
  }
}

// ─── New-Format Type Validation ──────────────────────────────────────────────

/**
 * Strict check: only LC, HC, Agri are valid type tokens.
 * Any other string causes the cutter block to be skipped.
 */
function isValidNewType(raw) {
  const t = raw.replace(/\*/g, "").trim().toLowerCase();
  return t === "lc" || t === "hc" || t === "agri";
}

// ─── New-Format Summary Block Parser ─────────────────────────────────────────

/**
 * Parse cutter blocks from a Cutting Summary message body.
 *
 * Supports three sub-formats:
 *   Section headers: "Total Tyres" / "Total Treads" + "LC - 38" / "HC - 0" (dash)
 *   Inline colon:   "LC: 45" / "Tread LC: 5"
 *   Multi-line:     "Tyre Type: LC" + "Quantity: 45" (same as hourly format)
 *
 * Returns [{cmNum, lc, hc, agri, tread_lc, tread_hc, tread_agri}]
 */
function parseSummaryBlocks(body) {
  const blocks = [];
  let cur = null;
  let section = null;       // "tyre" | "tread" — set by "Total Tyres"/"Total Treads" headers
  let lastTypeField = null; // "tyre" | "tread" — set by "Tyre Type:"/"Tread Type:" lines

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
    if (!s) return;
    if (/n\/?a|not\s+in\s+use|offloading|off\b/i.test(s)) return;

    const pairs = [...s.matchAll(/([A-Za-z][A-Za-z ]+?)\s*[-=:]\s*(\d+)/g)];
    if (pairs.length > 0) {
      for (const m of pairs) applyTypeQty(m[1], parseInt(m[2], 10));
      return;
    }

    const loose = s.match(/^([A-Za-z][A-Za-z ]+)\s+(\d+)$/);
    if (loose) applyTypeQty(loose[1], parseInt(loose[2], 10));
  }

  for (const rawLine of body.split("\n")) {
    const line = rawLine.trim();
    if (!line || /cutting summary/i.test(line)) continue;

    const cutterMatch = line.match(/^\*?(?:cutter|cm)\s*[- ]?(\d+)\*?(?:\s*[:=\-]\s*(.*))?$/i);
    if (cutterMatch) {
      if (cur) blocks.push(cur);
      cur = { cmNum: parseInt(cutterMatch[1], 10), lc: null, hc: null, agri: null, tread_lc: null, tread_hc: null, tread_agri: null };
      section = null;
      lastTypeField = null;
      if (cutterMatch[2]) parseInlinePayload(cutterMatch[2]);
      continue;
    }
    if (!cur) continue;

    // Section headers: "Total Tyres" / "Total Treads"
    if (/^total\s+tyre/i.test(line)) { section = "tyre";  continue; }
    if (/^total\s+tread/i.test(line)) { section = "tread"; continue; }

    // Inline: "Tread LC: 45" (explicit tread prefix — always goes to tread fields)
    const inlineTread = line.match(/^tread\s+(lc|hc|agri)\s*:\s*(\d+)/i);
    if (inlineTread) {
      const t = inlineTread[1].toLowerCase(), qty = parseInt(inlineTread[2], 10);
      if (t === "lc") cur.tread_lc = qty;
      else if (t === "hc") cur.tread_hc = qty;
      else cur.tread_agri = qty;
      continue;
    }

    // "LC: 45" / "LC - 45" etc. — route to tyre vs tread based on active section header
    const typeQty = line.match(/^(lc|hc|agri)\s*[:\-–]\s*(\d+)/i);
    if (typeQty) {
      const t = typeQty[1].toLowerCase(), qty = parseInt(typeQty[2], 10);
      if (section === "tread") {
        if (t === "lc") cur.tread_lc = qty;
        else if (t === "hc") cur.tread_hc = qty;
        else cur.tread_agri = qty;
      } else {
        if (t === "lc") cur.lc = qty;
        else if (t === "hc") cur.hc = qty;
        else cur.agri = qty;
      }
      continue;
    }

    // Legacy old-format summaries often use full names:
    // "Light Commercial - 186", "Agricultural - 73", "Heavy Commercial - 30".
    // Parse these as summary values for the current cutter.
    parseInlinePayload(line);

    // Multi-line: "Tyre Type: LC" then "Quantity: 45"
    const m_tyre = line.match(/^t[iy]re\s+type\s*:\s*(.*)/i);
    const m_trd  = line.match(/^tread\s+type\s*:\s*(.*)/i);
    const m_qty  = line.match(/^quantit[yi]e?s?\s*:\s*(\d+)/i);

    if (m_tyre) { cur._tyreType  = m_tyre[1].trim().toLowerCase(); lastTypeField = "tyre";  continue; }
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

  // Strip internal temp fields before returning
  return blocks.map(({ _tyreType, _treadType, ...rest }) => rest);
}

// ─── Machine Row Factory ──────────────────────────────────────────────────────

/**
 * Create a blank combined machine row — exactly ONE row per machine per timeframe.
 * All quantity fields start as null (rendered as empty in the sheet).
 * Fields mirror the Excel column layout in cuttingSheetBuilder.js.
 */
export function makeMachineRow(date, cmNumber, series, startTime, finishTime, operator = "") {
  return {
    date, cmNumber, series, startTime, finishTime, operator,
    // PCR tyre quantities
    passenger: null, fourx4: null, motorcycle: null,
    // Tyre quantities (by type)
    light_commercial: null,   // col 9  — LC tyre
    heavy_commercial_t: null, // col 11 — HC tyre (RADIALS)
    agricultural_t: null,     // col 12 — Agri tyre (NYLONS, when no tread_agri)
    // Tread quantities (by type)
    tread_lc: null,           // col 10 — LC tread (RADIALS)
    tread_hc: null,           // col 13 — HC tread (NYLONS)
    tread_agri: null,         // col 12 — Agri tread (NYLONS, takes precedence over agricultural_t)
    // Old-format tread cut total
    treads: null,             // col 15
  };
}

// ─── Old-Format Untyped Count Resolution ─────────────────────────────────────

function rowDateKey(row) {
  if (!row?.date) return "";
  return `${row.date.year}-${String(row.date.month).padStart(2, "0")}-${String(row.date.day).padStart(2, "0")}`;
}

function knownTypeColumns(row) {
  const cols = [];
  if (row.light_commercial !== null && row.light_commercial !== undefined) cols.push("light_commercial");
  if (row.heavy_commercial_t !== null && row.heavy_commercial_t !== undefined) cols.push("heavy_commercial_t");
  if (row.agricultural_t !== null && row.agricultural_t !== undefined) cols.push("agricultural_t");
  return cols;
}

function resolveUntypedCounts(records) {
  const typesByDay = new Map();
  for (const r of records) {
    const d = rowDateKey(r);
    if (!d) continue;
    if (!typesByDay.has(d)) typesByDay.set(d, new Set());
    for (const col of knownTypeColumns(r)) typesByDay.get(d).add(col);
  }

  for (const r of records) {
    if (r._untypedCount === null || r._untypedCount === undefined) continue;
    let target = r._hintColumn || null;
    if (!target) {
      const dayTypes = typesByDay.get(rowDateKey(r));
      if (dayTypes && dayTypes.size === 1) target = [...dayTypes][0];
    }
    if (target) r[target] = (r[target] ?? 0) + r._untypedCount;
    delete r._untypedCount;
    delete r._hintColumn;
  }
}

function parseLegacyDailySummaryBlocks(body) {
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
    if (!line) continue;
    if (/daily summary/i.test(line)) continue;
    if (/cutting machines?/i.test(line)) continue;
    if (/^total\b/i.test(line)) continue;

    // Examples:
    // "Machine 1-65"
    // "Machine One - 96 LC"
    // "Machine Two - 159 Agri"
    const m = line.match(/^machine\s+(one|two|three|\d)\s*[-:]\s*(.+)$/i);
    if (!m) continue;
    const cmNum = parseCm(m[1]);
    if (!cmNum) continue;

    const payload = m[2].trim();
    const qtyMatch = payload.match(/(\d+)/);
    if (!qtyMatch) continue;

    const qty = parseInt(qtyMatch[1], 10);
    if (isNaN(qty)) continue;

    blocks.push({
      cmNum,
      qty,
      type: parseType(payload),
      raw: line,
    });
  }

  return blocks;
}

function parseLegacyCuttingSummaryBlocks(body) {
  const byCm = new Map();
  let currentCm = null;

  function ensure(cmNum) {
    if (!byCm.has(cmNum)) {
      byCm.set(cmNum, {
        cmNum,
        lc: null, hc: null, agri: null,
        tread_lc: null, tread_hc: null, tread_agri: null,
        _hasMarker: true,
      });
    }
    return byCm.get(cmNum);
  }

  function applyTypeQty(block, rawType, qty) {
    const col = mapTyreType(rawType);
    if (col === "light_commercial") block.lc = qty;
    else if (col === "heavy_commercial_t") block.hc = qty;
    else if (col === "agricultural_t") block.agri = qty;
    // Legacy summaries usually don't include split treads. If they do include
    // explicit tread labels, map via lightweight pattern checks.
    else if (/tread\s*lc/i.test(rawType)) block.tread_lc = qty;
    else if (/tread\s*hc/i.test(rawType)) block.tread_hc = qty;
    else if (/tread\s*agri/i.test(rawType)) block.tread_agri = qty;
  }

  function parseSegment(segment, block) {
    const s = segment.trim();
    if (!s) return;
    if (/n\/?a|not\s+in\s+use|offloading|off\b/i.test(s)) return;
    if (/^total\b/i.test(s)) return;

    const explicit = [...s.matchAll(/([A-Za-z][A-Za-z ]+?)\s*[-=:]\s*(\d+)/g)];
    if (explicit.length > 0) {
      for (const m of explicit) applyTypeQty(block, m[1].trim(), parseInt(m[2], 10));
      return;
    }

    const loose = s.match(/^([A-Za-z][A-Za-z ]+)\s+(\d+)$/);
    if (loose) applyTypeQty(block, loose[1].trim(), parseInt(loose[2], 10));
  }

  for (const rawLine of body.split("\n")) {
    const line = rawLine.trim();
    if (!line) continue;
    if (/cutting summary/i.test(line)) continue;
    if (/^(\d{1,2}\/\d{1,2}\/\d{4}|date\s*[-:])/i.test(line)) continue;
    if (/^total\b/i.test(line)) continue;

    const cmHeader = line.match(/^CM\s*[- ]?(\d)\s*=\s*(.*)$/i);
    if (cmHeader) {
      currentCm = parseInt(cmHeader[1], 10);
      const block = ensure(currentCm);
      parseSegment(cmHeader[2], block);
      continue;
    }

    if (currentCm) parseSegment(line, ensure(currentCm));
  }

  return [...byCm.values()];
}

function inferDailySummaryType(records, date, cmNumberLabel) {
  const dateKey = `${date.year}-${date.month}-${date.day}`;
  const cols = new Set();

  for (const r of records) {
    if (!r.date) continue;
    if (`${r.date.year}-${r.date.month}-${r.date.day}` !== dateKey) continue;
    if (r.cmNumber !== cmNumberLabel) continue;

    if (r.light_commercial !== null && r.light_commercial !== undefined) cols.add("lc");
    if (r.heavy_commercial_t !== null && r.heavy_commercial_t !== undefined) cols.add("hc");
    if (r.agricultural_t !== null && r.agricultural_t !== undefined) cols.add("agri");
  }

  return cols.size === 1 ? [...cols][0] : null;
}

// ─── New-Format Parser ───────────────────────────────────────────────────────

/**
 * Parser for the new structured WhatsApp cutting format.
 *
 * Hourly format:
 *   *Time*: 08:00-09:00
 *   *Cutter 1*
 *   Operator: Jane
 *   Assistant: Bob
 *   Tyre Type: LC          ← only LC / HC / Agri accepted
 *   Quantity: 45
 *   Tread Type: HC         ← optional; missing tread = partial parse (tyre only)
 *   Quantity: 12
 *
 * Summary format:
 *   Cutting Summary
 *   *Cutter 1*
 *   LC: 45  /  Tread LC: 5   (inline)
 *   — or —
 *   Tyre Type: LC  +  Quantity: 45  (multi-line, same as hourly)
 *
 * Returns { records, summaryRecords, validationLog }
 *   records        – hourly production rows (same schema as old parser)
 *   summaryRecords – one row per cutter per daily summary message
 *   validationLog  – QA log entries for duplicates, invalid types, missing qty
 */
export function parseCuttingMessagesNew(text) {
  const messages = splitWhatsAppMessages(text);
  const records        = [];
  const summaryRecords = [];
  const validationLog  = [];

  for (const msg of messages) {
    const body = msg.body;
    if (body.includes("<Media omitted>")) continue;

    const isStructuredSummary = /cutting summary/i.test(body);
    const isDailySummary = /daily summary/i.test(body);
    const isSummary = isStructuredSummary || isDailySummary;
    const isHourly  = !isSummary && /cutter\s+\d/i.test(body);
    if (!isSummary && !isHourly) continue;

    const date = extractBodyDate(body) ?? msg.tsDate;
    if (!date) continue;

    const dateStr = dateToStr(date);
    const series  = `${String(date.month).padStart(2, "0")}/${String(date.year).slice(2)}`;

    // ── Summary messages ───────────────────────────────────────────────────────
    if (isSummary) {
      let blocks = parseSummaryBlocks(body);
      if (blocks.length === 0 && isStructuredSummary) {
        blocks = parseLegacyCuttingSummaryBlocks(body);
      }

      if (blocks.length > 0) {
        for (const b of blocks) {
          // Keep partial summary blocks; skip only if completely empty.
          const hasAnyValue = (
            b.lc !== null || b.hc !== null || b.agri !== null ||
            b.tread_lc !== null || b.tread_hc !== null || b.tread_agri !== null
          );
          if (!hasAnyValue) {
            validationLog.push({
              date: dateStr, time: "", messageType: "Summary",
              cutter: `CM - ${b.cmNum}`,
              issue: "Summary block has no parseable tyre/tread values",
              action: "Summary row skipped",
            });
            continue;
          }
          summaryRecords.push({
            date, series,
            cmNumber: `CM - ${b.cmNum}`,
            lc: b.lc, hc: b.hc, agri: b.agri,
            tread_lc: b.tread_lc, tread_hc: b.tread_hc, tread_agri: b.tread_agri,
          });
        }
        continue;
      }

      // Legacy "Daily Summary" lines (e.g. "Machine 1-96"), supported in
      // new-mode too so mixed chats still produce the Daily Summary table.
      if (isDailySummary) {
        for (const b of parseLegacyDailySummaryBlocks(body)) {
          const cmLabel = `CM - ${b.cmNum}`;
          const inferredType = b.type ?? inferDailySummaryType(records, date, cmLabel);
          if (!inferredType) {
            validationLog.push({
              date: dateStr, time: "", messageType: "Summary",
              cutter: cmLabel,
              issue: `Ambiguous legacy daily summary type in "${b.raw}"`,
              action: "Summary row skipped",
            });
            continue;
          }

          summaryRecords.push({
            date, series,
            cmNumber: cmLabel,
            lc: inferredType === "lc" ? b.qty : null,
            hc: inferredType === "hc" ? b.qty : null,
            agri: inferredType === "agri" ? b.qty : null,
            tread_lc: null, tread_hc: null, tread_agri: null,
          });
        }
      }
      continue;
    }

    // ── Hourly messages ────────────────────────────────────────────────────────
    let startTime = null, finishTime = null;
    const slotMatch = body.match(/\*?time\*?\s*:\s*(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})/i);
    if (slotMatch) {
      startTime = parseTime(slotMatch[1]);
      finishTime = parseTime(slotMatch[2]);
    }
    const timeStr = (startTime && finishTime)
      ? `${formatTime(startTime)}-${formatTime(finishTime)}` : "";

    const seenCutters    = new Set();
    const producedCutters = new Set(); // tracks CMs that actually emitted at least one row
    let currentBlock  = null;
    let lastQtyField  = null;
    let isDuplicate   = false;

    // Flush + validate the current cutter block into records / log.
    // Emits EXACTLY ONE combined row per machine per timeframe.
    function flushBlock() {
      if (!currentBlock || isDuplicate) return;
      const { cmNum, operator, assistant, tyreType, tyreCount, treadType, treadCount } = currentBlock;
      const opStr = [operator, assistant].filter(Boolean).join(" / ");

      // Invalid type → skip whole block; placeholder row covers this machine
      if (tyreType !== null && !isValidNewType(tyreType)) {
        validationLog.push({
          date: dateStr, time: timeStr, messageType: "Hourly",
          cutter: `CM - ${cmNum}`,
          issue: `Invalid tyre type "${tyreType}"`,
          action: "Block skipped — only LC, HC, Agri are valid",
        });
        return;
      }
      if (treadType !== null && !isValidNewType(treadType)) {
        validationLog.push({
          date: dateStr, time: timeStr, messageType: "Hourly",
          cutter: `CM - ${cmNum}`,
          issue: `Invalid tread type "${treadType}"`,
          action: "Block skipped — only LC, HC, Agri are valid",
        });
        return;
      }

      // No tyre data → nothing to anchor the row; placeholder covers this machine
      if (tyreType === null) return;

      // Tyre quantity missing → log and skip; placeholder covers this machine
      if (tyreCount === null) {
        validationLog.push({
          date: dateStr, time: timeStr, messageType: "Hourly",
          cutter: `CM - ${cmNum}`,
          issue: "Tyre type present but quantity missing",
          action: "Block skipped",
        });
        return;
      }

      // Build one combined row for this machine/timeframe
      const row = makeMachineRow(date, `CM - ${cmNum}`, series, startTime, finishTime, opStr);
      const tyreCol = mapTyreTypeNew(tyreType);
      if (tyreCol !== "unknown_type") row[tyreCol] = tyreCount;

      // Tread — optional; log if missing or quantity absent
      if (treadType === null) {
        validationLog.push({
          date: dateStr, time: timeStr, messageType: "Hourly",
          cutter: `CM - ${cmNum}`,
          issue: "Tread section missing",
          action: "Partial parse — tyre data written, tread omitted",
        });
      } else if (treadCount === null) {
        validationLog.push({
          date: dateStr, time: timeStr, messageType: "Hourly",
          cutter: `CM - ${cmNum}`,
          issue: "Tread type present but quantity missing",
          action: "Tread omitted",
        });
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
          validationLog.push({
            date: dateStr, time: timeStr, messageType: "Hourly",
            cutter: `CM - ${cmNum}`,
            issue: `Duplicate cutter CM - ${cmNum} in same message`,
            action: "Kept first occurrence, ignored later duplicate",
          });
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
    flushBlock(); // flush last block

    // If at least one cutter produced data in this timeframe, include
    // placeholders for missing cutters to preserve CM-1/2/3 structure.
    if (producedCutters.size > 0) {
      for (const n of [1, 2, 3]) {
        if (!producedCutters.has(n)) {
          const placeholder = makeMachineRow(date, `CM - ${n}`, series, startTime, finishTime);
          placeholder._syntheticPlaceholder = true;
          records.push(placeholder);
        }
      }
    }
  }

  return { records, summaryRecords, validationLog };
}

// ─── Old-Format Parser ───────────────────────────────────────────────────────

/**
 * Orchestrate the pipeline for a full chat export.
 *
 * Each message body is first split into interval blocks at time-range lines,
 * so messages reporting multiple cutting periods are handled correctly.
 * seenCMs is per-interval, allowing the same machine to appear in different
 * intervals without being deduplicated.
 *
 * Empty shell rows (no time, no count, no status) are discarded.
 *
 * Returns { records, summaryRecords, validationLog }.
 */
export function parseCuttingMessages(text) {
  const messages = splitWhatsAppMessages(text);
  const records = [];
  const summaryRecords = [];
  const validationLog = [];

  for (const msg of messages) {
    const body = msg.body;

    if (body.includes("<Media omitted>")) continue;

    // Use the WhatsApp system timestamp date as authoritative.
    const date = extractBodyDate(body) ?? msg.tsDate;
    if (!date) continue;

    const series  = `${String(date.month).padStart(2, "0")}/${String(date.year).slice(2)}`;
    const dateStr = dateToStr(date);
    const isStructuredSummary = /cutting summary/i.test(body);
    const isDailySummary = /daily summary/i.test(body);

    if (isStructuredSummary) {
      let blocks = parseSummaryBlocks(body);
      if (blocks.length === 0) blocks = parseLegacyCuttingSummaryBlocks(body);

      for (const b of blocks) {
        const hasAnyValue = (
          b.lc !== null || b.hc !== null || b.agri !== null ||
          b.tread_lc !== null || b.tread_hc !== null || b.tread_agri !== null
        );
        if (!hasAnyValue && !b._hasMarker) {
          validationLog.push({
            date: dateStr, time: "", messageType: "Summary",
            cutter: `CM - ${b.cmNum}`,
            issue: "Summary block has no parseable tyre/tread values",
            action: "Summary row skipped",
          });
          continue;
        }
        summaryRecords.push({
          date, series,
          cmNumber: `CM - ${b.cmNum}`,
          lc: b.lc, hc: b.hc, agri: b.agri,
          tread_lc: b.tread_lc, tread_hc: b.tread_hc, tread_agri: b.tread_agri,
        });
      }
      continue;
    }

    // Legacy "Daily Summary" messages (old format):
    // parse machine totals and infer tyre type from same-day hourly rows if needed.
    if (isDailySummary) {
      for (const b of parseLegacyDailySummaryBlocks(body)) {
        const cmLabel = `CM - ${b.cmNum}`;
        const inferredType = b.type ?? inferDailySummaryType(records, date, cmLabel);
        if (!inferredType) {
          validationLog.push({
            date: dateStr, time: "", messageType: "Summary",
            cutter: cmLabel,
            issue: `Ambiguous legacy daily summary type in "${b.raw}"`,
            action: "Summary row skipped",
          });
          continue;
        }

        summaryRecords.push({
          date, series,
          cmNumber: cmLabel,
          lc: inferredType === "lc" ? b.qty : null,
          hc: inferredType === "hc" ? b.qty : null,
          agri: inferredType === "agri" ? b.qty : null,
          // Old format daily summaries usually omit tread split.
          tread_lc: null, tread_hc: null, tread_agri: null,
        });
      }
      continue;
    }

    // "Tread Cuts" messages are NOT skipped wholesale: mixed messages may also
    // contain valid cutting rows.  The line classifier skips individual tread
    // lines, so only the cutting CM lines are extracted.
    if (!/CM\s*\d|Machine\s+\d/i.test(body)) continue;

    // ── Split body into interval blocks ────────────────────────────────────────
    // Each time-range line (e.g. "08:00-09:00" or "Cutting (12:01-13:00)") starts
    // a new interval.  Machine lines belong to the most recent interval.
    const intervals = [];
    let cur = { startTime: null, finishTime: null, lines: [] };
    for (const rawLine of body.split("\n")) {
      const timeMatch = rawLine.match(/(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})/);
      if (timeMatch) {
        intervals.push(cur);
        cur = {
          startTime: parseTime(timeMatch[1]),
          finishTime: parseTime(timeMatch[2]),
          lines: [],
        };
      } else {
        cur.lines.push(rawLine);
      }
    }
    intervals.push(cur);

    // ── Process each interval block independently ──────────────────────────────
    for (const interval of intervals) {
      if (interval.lines.length === 0) continue;

      const { startTime, finishTime } = interval;
      const seenCMs   = new Set();
      const machineRows = new Map(); // cmNum → one combined row per machine
      let hasIntervalProduction = false;
      let intervalHintColumn = null;
      let m;

      // Helper: get or create the combined row for a CM number
      const getRow = (cmNum) => {
        if (!machineRows.has(cmNum)) {
          machineRows.set(cmNum, makeMachineRow(date, `CM - ${cmNum}`, series, startTime, finishTime));
        }
        return machineRows.get(cmNum);
      };

      // Phase 1: format D (HC + LC per CM, may span two body lines)
      const intervalBody = interval.lines.join("\n");
      const fmtD = /CM\s*(\d)\s*-\s*\(HC\)\s*=?\s*(\d+)[\s\S]*?\(LC\)\s*=?\s*(\d+)/gi;
      while ((m = fmtD.exec(intervalBody)) !== null) {
        const cmNum = parseInt(m[1], 10);
        if (seenCMs.has(cmNum)) continue;
        seenCMs.add(cmNum);
        const hc = parseInt(m[2], 10), lc = parseInt(m[3], 10);
        const row = getRow(cmNum);
        if (!isNaN(hc) && hc > 0) row.heavy_commercial_t = hc;
        if (!isNaN(lc) && lc > 0) row.light_commercial = lc;
        if ((!isNaN(hc) && hc > 0) || (!isNaN(lc) && lc > 0)) hasIntervalProduction = true;
        // Both zero/NaN: row still exists (getRow already registered it)
      }

      // Single-(HC) lines for CMs not captured by format D above
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

      // Phase 2: line-by-line for all other formats
      // Two-pass approach: first collect parsed results, then infer column for
      // any column:null records from sibling CMs in the same interval.
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

        for (const p of parsed) {
          pendingRecords.push(p);
        }
      }

      // Context-aware type inference: if all sibling CMs in this interval that
      // DO have a known column share the same type, assign it to column:null
      // records.  This handles bare-count lines like "CM1-8" that appear
      // alongside typed lines like "CM2-16 Agri".
      const knownColumns = pendingRecords
        .filter(p => p.column !== null && p.column !== "unknown_type" && p.column !== "")
        .map(p => p.column);
      const uniqueKnown = [...new Set(knownColumns)];
      const inferredColumn = uniqueKnown.length === 1 ? uniqueKnown[0] : null;

      for (const p of pendingRecords) {
        // Drop entries that have neither a count nor a status flag —
        // but keep entries with a count even when column is null.
        if (p.count === null && !p.isStatus) continue;

        let col = p.column;
        if ((col === null || col === "unknown_type") && p.count !== null && inferredColumn) {
          col = inferredColumn;
        }
        if ((col === null || col === "unknown_type") && p.count !== null && intervalHintColumn) {
          col = intervalHintColumn;
        }
        const row = getRow(p.cmNum);
        // Accumulate quantity into the field if we have a valid column and count
        if (col && col !== "unknown_type" && col !== "" && p.count !== null) {
          row[col] = p.count;
          if (p.count > 0) hasIntervalProduction = true;
        } else if (p.count !== null) {
          row._untypedCount = (row._untypedCount ?? 0) + p.count;
          if (!row._hintColumn && intervalHintColumn) row._hintColumn = intervalHintColumn;
          if (p.count > 0) hasIntervalProduction = true;
        }
        // Status rows (isStatus, no count) already ensured the CM row exists via getRow
      }

      // No cutter updated with numeric data in this timeframe:
      // skip the interval entirely (avoid empty synthetic lines).
      if (!hasIntervalProduction) continue;

      // Emit one row per machine that appeared in this interval
      for (const row of machineRows.values()) {
        records.push(row);
      }

      // At least one machine produced data in this interval:
      // keep CM-1/2/3 placeholders for missing machines.
      for (const n of [1, 2, 3]) {
        if (!machineRows.has(n)) {
          const placeholder = makeMachineRow(date, `CM - ${n}`, series, startTime, finishTime);
          placeholder._syntheticPlaceholder = true;
          records.push(placeholder);
        }
      }
    }
  }

  resolveUntypedCounts(records);
  return { records, summaryRecords, validationLog };
}
