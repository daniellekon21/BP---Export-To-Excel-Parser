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
  s = s.replace(/\s*<This message was edited>\s*$/i, "");
  s = s.replace(/[–—−]/g, "-");
  s = s.replace(/\b(?:CM\s+|Machine\s+)(\d)/gi, "CM$1");
  s = s.replace(/(\d)(?![xX]\d)([A-Za-z])/g, "$1 $2");
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
  const t = raw.trim().toLowerCase().replace(/[().]/g, " ").replace(/\s+/g, " ").trim();
  // "offloading truck" is a status comment (offloading a physical truck), not a tyre type
  if (t.includes("offloading"))                                              return "unknown_type";
  if (/^(?:s|m|l|small|medium|large)$/.test(t))                              return "radialsAgri";
  if (t.includes("nylon"))                                                  return "nylonsLC";
  if (t.includes("tread"))                                                   return "radialsAgriTreads";
  if (t.includes("agri"))                                                    return "radialsAgri";
  if (/4.?x.?4/i.test(t))                                                   return "radialsLC";
  if (t.includes("radial") && (t.includes("lc") || t.includes("light")))   return "radialsLC";
  if (t.includes("light") || t.includes("lc"))                              return "radialsLC";
  if (t.includes("heavy") || t.includes("hc") || t.includes("truck") || t.includes("radial")) return "radialsHC";
  if (/^tyres?$/.test(t))                                                    return "radialsAgri";
  return "unknown_type";
}

// New-format tyre abbreviation → column key.
export function mapTyreTypeNew(raw) {
  const t = raw.replace(/\*/g, "").trim().toLowerCase();
  if (/^(?:s|m|l|small|medium|large)$/.test(t)) return "radialsAgri";
  if (t.includes("nylon"))              return "nylonsLC";
  if (/4.?x.?4/i.test(t))               return "radialsLC";
  if (t === "lc" || t.includes("light")) return "radialsLC";
  if (t === "hc" || t.includes("heavy")) return "radialsHC";
  if (t.startsWith("agri"))              return "radialsAgri";
  return "unknown_type";
}

// New-format tread abbreviation → column key.
export function mapTreadTypeNew(raw) {
  const t = raw.replace(/\*/g, "").trim().toLowerCase();
  if (t === "lc" || t.includes("light")) return "radialsAgriTreads";
  if (t === "hc" || t.includes("heavy")) return "radialsAgriTreads";
  if (t.startsWith("agri"))              return "radialsAgriTreads";
  return "unknown_type";
}

// ─── Parse Machine Line ───────────────────────────────────────────────────────

// Parse one normalized machine line into [{cmNum, column, count}].
// Format D (HC+LC dual-type) is handled at body level before this is called.
export function parseMachineLine(line) {
  let m;

  // Strip trailing status commentary after a parsed count
  line = line.replace(/(CM\s*\d\s*[-=]\s*\d+\s+[A-Za-z][A-Za-z.\s()]*)\s*-\s*(?:stopped|idle|waiting|offloading|breakdown|cleaning|paused|maintenance|no\s+production|not\s+(?:in\s+use|cutting|working))\b.*$/i, "$1");
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

  // Dash-separated dual type: "CM3-3 HC -1 LC"
  m = line.match(/CM\s*(\d)\s*[-=]\s*(\d+)\s+([A-Za-z][A-Za-z.]*)\s+-\s*(\d+)\s+([A-Za-z][A-Za-z.]*)\s*$/i);
  if (m) {
    const cmNum = parseInt(m[1], 10);
    const results = [];
    const c1 = parseInt(m[2], 10); if (!isNaN(c1)) results.push({ cmNum, column: mapTyreType(m[3]), count: c1 });
    const c2 = parseInt(m[4], 10); if (!isNaN(c2)) results.push({ cmNum, column: mapTyreType(m[5]), count: c2 });
    if (results.length) return results;
  }

  // Side-wall only: "CM1-51x Side Walls" / "CM1-51 Side Walls" / "CM1-51 SW"
  // Convert side walls to tyre-equivalent count by dividing by two and rounding up.
  m = line.match(/CM\s*(\d)\s*[-=]\s*(\d+)\s*(?:x\s*)?(?:side\s*walls?|sw)\b/i);
  if (m) {
    const cmNum = parseInt(m[1], 10);
    const sideWallCount = parseInt(m[2], 10);
    if (!isNaN(sideWallCount)) {
      return [{
        cmNum,
        column: null,
        count: Math.ceil(sideWallCount / 2),
        isSideWall: true,
        sideWallCount,
      }];
    }
  }

  // Tread Cut: "CM2-14 (Agri) - Tread Cut 21"
  m = line.match(/CM\s*(\d)\s*[-=]\s*(\d+)\s*\(?([A-Za-z.]+)\)?\s*-\s*Tread\s*Cut\s*(\d+)/i);
  if (m) {
    return [
      { cmNum: parseInt(m[1], 10), column: mapTyreType(m[3]), count: parseInt(m[2], 10) },
      { cmNum: parseInt(m[1], 10), column: "radialsAgriTreads", count: parseInt(m[4], 10) },
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
  m = line.match(/CM\s*(\d)\s+([A-Za-z][A-Za-z. ]*?)\s+-\s*(\d+)\s*$/i);
  if (m) {
    const count = parseInt(m[3], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[2]), count }];
  }

  // Format A: CM1 = TypeName - Count
  m = line.match(/CM\s*(\d)\s*=\s*([A-Za-z][A-Za-z. ]*?)\s*-\s*(\d+|N\/?A)\b/i);
  if (m) {
    const countStr = m[3].trim();
    const count = /n\/?a/i.test(countStr) ? null : parseInt(countStr, 10);
    if (count === null || !isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[2]), count }];
  }

  // NA: CM1 = N/A / CM1-not in use / etc.
  m = line.match(/CM\s*(\d)\s*[-=\s]+(N\/?A|not\s+in\s+use|breakdown|paused|off|offline|offloading|beltrim|maintenance|shortstaffed|short\s*staffed)\b/i);
  if (m) return [{ cmNum: parseInt(m[1], 10), column: "", count: null, isStatus: true }];

  // Zero-count with parenthesized reason: CM3-0 (Offline-shortstaffed)
  m = line.match(/CM\s*(\d)\s*[-=]\s*0\s*\(.*\)?\s*$/i);
  if (m) return [{ cmNum: parseInt(m[1], 10), column: "", count: 0, isStatus: true }];

  // Format E: CM1 - TypeName = Count
  m = line.match(/CM\s*(\d)\s*-\s*([A-Za-z][A-Za-z. ]*?)\s*=\s*(\d+)/i);
  if (m) {
    const count = parseInt(m[3], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[2]), count }];
  }

  // Format C: CM1 - Count TypeName (handles optional parens + LC qualifier)
  m = line.match(/CM\s*(\d)\s*-\s*(\d+)\s*[-\s]+\(?([A-Za-z][A-Za-z.\s()]*?)\s*-?\s*$/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[3]), count }];
  }

  // Format B: CM1 = Count TypeName
  m = line.match(/CM\s*(\d)\s*=\s*(\d+)\s+([A-Za-z][A-Za-z.\s()]*?)\s*$/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[3]), count }];
  }

  // Format F: CM1 Count TypeName (no separator)
  m = line.match(/^CM\s*(\d)\s+(\d+)\s+([A-Za-z][A-Za-z.\s()]*?)\s*$/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: mapTyreType(m[3]), count }];
  }

  // 4x4 type: CM2-30 4x4
  m = line.match(/CM\s*(\d)\s*[-=]\s*(\d+)\s+(4\s*[xX×]\s*4)\s*$/i);
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

  // Fallback: CM1-count with any trailing text (unrecognized type description)
  m = line.match(/CM\s*(\d)\s*[-=]\s*(\d+)\b/i);
  if (m) {
    const count = parseInt(m[2], 10);
    if (!isNaN(count)) return [{ cmNum: parseInt(m[1], 10), column: null, count }];
  }

  return [];
}

export function parseStandaloneTreadLine(line) {
  const m = line.match(/^\s*treads?\s+cuts?\.?\s*[-:]\s*(\d+)\s*$/i);
  if (!m) return null;
  const count = parseInt(m[1], 10);
  return Number.isNaN(count) ? null : count;
}

// ─── Machine Row Factory ──────────────────────────────────────────────────────

// Create a blank machine row — one per machine per timeframe.
export function makeMachineRow(date, cmNumber, series, startTime, finishTime, operator = "", rawMessage = "") {
  return {
    date, cmNumber, series, startTime, finishTime, operator, rawMessage,
    radialsLC: null,
    radialsHC: null,
    radialsAgri: null,
    radialsAgriTreads: null,
    nylonsLC: null,
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
  return t === "lc" || t === "hc" || t === "agri" || t === "nylon" || t === "nylons" || /^(?:s|m|l|small|medium|large)$/.test(t) || /4.?x.?4/i.test(t);
}

// ─── Summary Block Parsers ────────────────────────────────────────────────────

// Parse cutter blocks from a Cutting Summary message body.
// Returns [{cmNum, radialsLC, radialsHC, radialsAgri, radialsAgriTreads}]
export function parseSummaryBlocks(body) {
  const blocks = [];
  let cur = null;
  let section = null;
  let lastTypeField = null;
  let seenTotal = false;

  function applyTypeQty(rawType, qty, forceTread = false) {
    const t = rawType.trim().toLowerCase();
    if (forceTread || /^tread\s+lc$/.test(t) || /^tread\s+hc$/.test(t) || /^tread\s+agri/.test(t)) {
      cur.radialsAgriTreads = (cur.radialsAgriTreads ?? 0) + qty;
      return;
    }
    // Side walls → HC, qty = ceil(swQty / 2)
    if (/^(?:side\s*walls?|sw)$/i.test(t)) { cur.radialsHC = (cur.radialsHC ?? 0) + Math.ceil(qty / 2); return; }
    // Bare "radials" or "tyres" (no LC/HC qualifier) → untyped total for inference
    if (/^(?:radials?|tyres?)$/.test(t)) { cur.radialsTotal = (cur.radialsTotal ?? 0) + qty; return; }
    if (t.includes("nylon")) cur.nylonsLC = qty;
    else if (/^(?:s|m|l|small|medium|large)$/.test(t)) cur.radialsAgri = qty;
    else if (/4.?x.?4/i.test(t)) cur.radialsLC = qty;
    else if (t.includes("light") || /\blc\b/.test(t)) cur.radialsLC = qty;
    else if (t.includes("heavy") || t.includes("truck") || /\bhc\b/.test(t)) cur.radialsHC = qty;
    else if (t.includes("agri")) cur.radialsAgri = qty;
  }

  function parseInlinePayload(payload) {
    const s = payload.replace(/\*/g, "").trim();
    if (!s || /n\/?a|not\s+in\s+use|offloading|off\b/i.test(s)) return;
    // Split compound segments like "72 Agri + 213 Treads" by "+"
    if (s.includes("+")) {
      for (const part of s.split(/\s*\+\s*/)) {
        const p = part.trim();
        if (p) parseInlinePayload(p);
      }
      return;
    }
    const pairs = [...s.matchAll(/([A-Za-z][A-Za-z ]+?)\s*[-=:]\s*(\d+)/g)];
    if (pairs.length > 0) { for (const m of pairs) applyTypeQty(m[1], parseInt(m[2], 10)); return; }
    const loose = s.match(/^([A-Za-z][A-Za-z ]+)\s+(\d+)$/);
    if (loose) { applyTypeQty(loose[1], parseInt(loose[2], 10)); return; }
    // "Count Type" format: e.g. "119 Agri", "131- Light Commercial", "51x Side Walls"
    const countFirst = s.match(/^(\d+)\s*x?\s*[-–]?\s*\(?([A-Za-z][A-Za-z ]*[A-Za-z])\)?\.?$/);
    if (countFirst) { applyTypeQty(countFirst[2], parseInt(countFirst[1], 10)); return; }
    // Bare number (no type label): e.g. "168" from "Machine 1-168"
    const bareQty = s.match(/^(\d+)\s*$/);
    if (bareQty) {
      const qty = parseInt(bareQty[1], 10);
      if (!isNaN(qty)) cur._untypedQty = qty;
    }
  }

  for (const rawLine of body.split("\n")) {
    const line = rawLine.trim();
    const clean = line.replace(/\*/g, "").trim();
    if (!clean || /(?:cutting\s+)?summary\b/i.test(clean)) continue;
    // Skip total lines (e.g. "Total= 281 light commercial Tyres", "Total Tyres - 321")
    // Once we see a total, stop attributing stray lines to the last CM block
    if (/^total\b/i.test(clean)) { seenTotal = true; continue; }

    const cutterMatch = line.match(/^\*?(?:cutter|cm|machine)\s*[- ]?(\d+)\*?(?:\s*[:=\-]\s*(.*))?$/i);
    if (cutterMatch) {
      if (cur) blocks.push(cur);
      cur = { cmNum: parseInt(cutterMatch[1], 10), radialsLC: null, radialsHC: null, radialsAgri: null, radialsAgriTreads: null, nylonsLC: null, radialsTotal: null };
      section = null; lastTypeField = null; seenTotal = false;
      if (cutterMatch[2]) parseInlinePayload(cutterMatch[2]);
      continue;
    }
    if (!cur || seenTotal) continue;

    if (/^total\s+tyre/i.test(line)) { section = "tyre"; continue; }
    if (/^total\s+tread/i.test(line)) { section = "tread"; continue; }

    const inlineTread = line.match(/^tread\s+(lc|hc|agri)\s*:\s*(\d+)/i);
    if (inlineTread) {
      const qty = parseInt(inlineTread[2], 10);
      cur.radialsAgriTreads = (cur.radialsAgriTreads ?? 0) + qty;
      continue;
    }

    const typeQty = line.match(/^(lc|hc|agri)\s*[:\-–]\s*(\d+)/i);
    if (typeQty) {
      const t = typeQty[1].toLowerCase(), qty = parseInt(typeQty[2], 10);
      if (section === "tread") {
        cur.radialsAgriTreads = (cur.radialsAgriTreads ?? 0) + qty;
      } else {
        if (t === "lc") cur.radialsLC = qty; else if (t === "hc") cur.radialsHC = qty; else cur.radialsAgri = qty;
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
        if (t === "lc") cur.radialsLC = qty;
        else if (t === "hc") cur.radialsHC = qty;
        else if (t === "agri") cur.radialsAgri = qty;
        else if (t === "nylon" || t === "nylons") cur.nylonsLC = qty;
        else if (/4.?x.?4/i.test(t)) cur.radialsLC = qty;
        else if (/^(?:s|m|l|small|medium|large)$/.test(t)) cur.radialsAgri = qty;
      } else if (lastTypeField === "tread" && cur._treadType) {
        cur.radialsAgriTreads = (cur.radialsAgriTreads ?? 0) + qty;
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
    if (/\b(?:s|m|l|small|medium|large)\b/.test(t)) return "radialsAgri";
    if (t.includes("nylon")) return "nylonsLC";
    if (/4.?x.?4/i.test(t)) return "radialsLC";
    if (t.includes("light") || /\blc\b/.test(t)) return "radialsLC";
    if (t.includes("truck") || t.includes("heavy") || /\bhc\b/.test(t)) return "radialsHC";
    if (t.includes("agri")) return "radialsAgri";
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

    // Compound: "64 Tyres + 90 Treads" or "85 Tyres + 85 Treads"
    const compound = payload.match(/(\d+)\s*(?:tyres?)\s*\+\s*(\d+)\s*(?:treads?)/i);
    if (compound) {
      const qty = parseInt(compound[1], 10);
      const treadQty = parseInt(compound[2], 10);
      if (!isNaN(qty)) blocks.push({ cmNum, qty, type: parseType(payload), treadQty: isNaN(treadQty) ? null : treadQty, raw: line });
      continue;
    }

    // Single with "Tyres" label: "137 Tyres"
    const tyreOnly = payload.match(/(\d+)\s*(?:tyres?)\s*$/i);
    if (tyreOnly) {
      const qty = parseInt(tyreOnly[1], 10);
      if (!isNaN(qty)) blocks.push({ cmNum, qty, type: parseType(payload), treadQty: null, raw: line });
      continue;
    }

    const qtyMatch = payload.match(/(\d+)/);
    if (!qtyMatch) continue;
    const qty = parseInt(qtyMatch[1], 10);
    if (!isNaN(qty)) blocks.push({ cmNum, qty, type: parseType(payload), treadQty: null, raw: line });
  }
  return blocks;
}

export function parseLegacyCuttingSummaryBlocks(body) {
  const byCm = new Map();
  let currentCm = null;

  function parseCm(raw) {
    const t = raw.trim().toLowerCase();
    if (t === "1" || t === "one") return 1;
    if (t === "2" || t === "two") return 2;
    if (t === "3" || t === "three") return 3;
    return null;
  }

  function ensure(cmNum) {
    if (!byCm.has(cmNum)) {
      byCm.set(cmNum, {
        cmNum,
        radialsLC: null, radialsHC: null, radialsAgri: null,
        radialsAgriTreads: null, nylonsLC: null, radialsTotal: null,
        _hasMarker: true,
        _untypedQty: null,
      });
    }
    return byCm.get(cmNum);
  }
  function applyTypeQty(block, rawType, qty) {
    const t = rawType.trim().toLowerCase();
    if (/tread\s*(lc|hc|agri)/i.test(rawType) || /\btreads?\b/i.test(rawType)) block.radialsAgriTreads = (block.radialsAgriTreads ?? 0) + qty;
    // Side walls → HC, qty = ceil(swQty / 2)
    else if (/^(?:side\s*walls?|sw)$/i.test(t)) block.radialsHC = (block.radialsHC ?? 0) + Math.ceil(qty / 2);
    else if (/^(?:radials?|tyres?)$/.test(t)) block.radialsTotal = (block.radialsTotal ?? 0) + qty;
    else if (t.includes("nylon")) block.nylonsLC = qty;
    else if (/^(?:s|m|l|small|medium|large)$/.test(t)) block.radialsAgri = qty;
    else if (/4.?x.?4/i.test(t)) block.radialsLC = qty;
    else if (t.includes("light") || /\blc\b/.test(t)) block.radialsLC = qty;
    else if (t.includes("heavy") || t.includes("truck") || /\bhc\b/.test(t)) block.radialsHC = qty;
    else if (t.includes("agri")) block.radialsAgri = qty;
  }
  function parseSegment(segment, block) {
    const s = segment.trim();
    if (!s || /n\/?a|not\s+in\s+use|offloading|off\b/i.test(s) || /^total\b/i.test(s)) return;
    // Split compound segments like "72 Agri + 213 Treads" by "+"
    if (s.includes("+")) {
      for (const part of s.split(/\s*\+\s*/)) {
        const p = part.trim();
        if (p) parseSegment(p, block);
      }
      return;
    }
    const explicit = [...s.matchAll(/([A-Za-z][A-Za-z ]+?)\s*[-=:]\s*(\d+)/g)];
    if (explicit.length > 0) { for (const m of explicit) applyTypeQty(block, m[1].trim(), parseInt(m[2], 10)); return; }
    const loose = s.match(/^([A-Za-z][A-Za-z ]+)\s+(\d+)$/);
    if (loose) { applyTypeQty(block, loose[1].trim(), parseInt(loose[2], 10)); return; }

    const countFirst = s.match(/^(\d+)\s*x?\s*[-–]?\s*\(?([A-Za-z][A-Za-z.\s()]*)\)?$/);
    if (countFirst) {
      const qty = parseInt(countFirst[1], 10);
      const rawType = countFirst[2].trim();
      if (!isNaN(qty) && rawType) { applyTypeQty(block, rawType, qty); return; }
    }

    const bareQty = s.match(/^(\d+)\s*$/);
    if (bareQty) {
      const qty = parseInt(bareQty[1], 10);
      if (!isNaN(qty)) block._untypedQty = qty;
    }
  }
  for (const rawLine of body.split("\n")) {
    const line = rawLine.trim();
    if (!line || /(?:cutting\s+)?summary\b/i.test(line) || /^(\d{1,2}\/\d{1,2}\/\d{4}|date\s*[-:])/i.test(line) || /^total\b/i.test(line)) continue;

    const machineHeader = line.match(/^machine\s+(one|two|three|\d)\s*[-:]\s*(.*)$/i);
    if (machineHeader) {
      currentCm = parseCm(machineHeader[1]);
      if (!currentCm) continue;
      parseSegment(machineHeader[2], ensure(currentCm));
      continue;
    }

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
    if (r.radialsLC != null) cols.add("radialsLC");
    if (r.radialsHC != null) cols.add("radialsHC");
    if (r.radialsAgri != null) cols.add("radialsAgri");
    if (r.nylonsLC != null) cols.add("nylonsLC");
  }
  return cols.size === 1 ? [...cols][0] : null;
}

// ─── Old-Format Untyped Count Resolution ──────────────────────────────────────

export function resolveUntypedCounts(records, summaryRecords = [], validationLog = []) {
  const dateKey = (d) => d ? `${d.year}-${String(d.month).padStart(2,"0")}-${String(d.day).padStart(2,"0")}` : "";
  const cmNum = (cmLabel) => {
    const m = String(cmLabel ?? "").match(/(\d+)/);
    return m ? parseInt(m[1], 10) : null;
  };
  const machineDayKey = (d, cmLabel) => {
    const dk = dateKey(d);
    const cm = cmNum(cmLabel);
    return dk && cm ? `${dk}|${cm}` : "";
  };
  const typedColumnFromRecord = (r) => {
    const cols = [];
    if (r.radialsLC != null) cols.push("radialsLC");
    if (r.radialsHC != null) cols.push("radialsHC");
    if (r.radialsAgri != null) cols.push("radialsAgri");
    if (r.nylonsLC != null) cols.push("nylonsLC");
    return cols.length === 1 ? cols[0] : null;
  };

  // Build summary type lookup — typed columns + bare radials flag
  const summaryTypeByMachineDay = new Map();
  const bareSummaryByMachineDay = new Map(); // key → [summaryRecord, ...]
  for (const s of summaryRecords) {
    const key = machineDayKey(s?.date, s?.cmNumber);
    if (!key) continue;
    if (!summaryTypeByMachineDay.has(key)) summaryTypeByMachineDay.set(key, new Set());
    const bucket = summaryTypeByMachineDay.get(key);
    if (s.totalLC != null) bucket.add("radialsLC");
    if (s.totalHC != null) bucket.add("radialsHC");
    if (s.totalAgri != null) bucket.add("radialsAgri");
    // Track bare radials summaries (totalRadials set but no specific type)
    if (s.totalRadials != null && s.totalLC == null && s.totalHC == null && s.totalAgri == null) {
      if (!bareSummaryByMachineDay.has(key)) bareSummaryByMachineDay.set(key, []);
      bareSummaryByMachineDay.get(key).push(s);
    }
  }

  const rowsByMachineDay = new Map();
  for (let i = 0; i < records.length; i += 1) {
    const r = records[i];
    const key = machineDayKey(r?.date, r?.cmNumber);
    if (!key) continue;
    if (!rowsByMachineDay.has(key)) rowsByMachineDay.set(key, []);
    rowsByMachineDay.get(key).push(i);
  }

  for (let i = 0; i < records.length; i += 1) {
    const r = records[i];
    if (r._untypedCount == null || r._untypedCount === 0) continue;

    const key = machineDayKey(r?.date, r?.cmNumber);
    const bucket = key ? (rowsByMachineDay.get(key) || []) : [];
    const pos = bucket.indexOf(i);

    let target = null;
    let source = "";

    // 1. Try daily summary for same machine first (only when summary has a specific type)
    const summaryTypes = key ? summaryTypeByMachineDay.get(key) : null;
    if (summaryTypes && summaryTypes.size === 1) {
      target = [...summaryTypes][0];
      source = "daily summary for same machine";
    }

    // 2. Try previous same-day records for same machine
    if (!target && pos >= 0) {
      for (let p = pos - 1; p >= 0; p -= 1) {
        const prev = records[bucket[p]];
        const prevCol = typedColumnFromRecord(prev);
        if (prevCol) {
          target = prevCol;
          source = "previous same-day text for same machine";
          break;
        }
      }
    }

    // 3. Try following same-day records for same machine
    if (!target && pos >= 0) {
      for (let n = pos + 1; n < bucket.length; n += 1) {
        const next = records[bucket[n]];
        const nextCol = typedColumnFromRecord(next);
        if (nextCol) {
          target = nextCol;
          source = "following same-day text for same machine";
          break;
        }
      }
    }

    // If resolved from hourly texts and the daily summary was bare radials,
    // update the summary records to reflect the actual type
    if (target && key && bareSummaryByMachineDay.has(key)) {
      const summaryField = target === "radialsLC" ? "totalLC"
        : target === "radialsHC" ? "totalHC"
        : target === "radialsAgri" ? "totalAgri"
        : null;
      if (summaryField) {
        for (const sr of bareSummaryByMachineDay.get(key)) {
          if (sr[summaryField] == null) {
            sr[summaryField] = sr.totalRadials;
          }
        }
      }
    }

    // 4. Bare radials summary — fallback to HC only if hourly texts didn't resolve it
    if (!target && key && bareSummaryByMachineDay.has(key)) {
      target = "radialsHC";
      source = "bare radials daily summary for same machine (defaulted to HC)";
    }

    const timeStr = (r.startTime && r.finishTime) ? `${formatTime(r.startTime)}-${formatTime(r.finishTime)}` : "";
    if (target) {
      r[target] = (r[target] ?? 0) + r._untypedCount;
      // Side walls are definitively HC — don't mark as inferred
      if (!r._sideWallCount) r._inferredType = true;
      validationLog.push({
        date: dateToStr(r.date),
        time: timeStr,
        messageType: "Hourly",
        cutter: r.cmNumber,
        issue: r._sideWallCount != null
          ? `Side wall only count (${r._sideWallCount} side walls => ${r._untypedCount} tyres)`
          : `Untyped machine count (${r._untypedCount})`,
        action: r._sideWallCount != null
          ? `Converted side walls to tyre-equivalent count and inferred as ${target} from ${source}`
          : `Inferred as ${target} from ${source}`,
        rawText: r.rawMessage || "",
      });
    } else {
      // Mark this record so the Excel writer can highlight it in the monthly sheet
      r._unresolvedType = true;
    }

    delete r._untypedCount;
    delete r._hintColumn;
    delete r._sideWallCount;
  }
}
