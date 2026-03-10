// ─── Cutting Sheet Builder ────────────────────────────────────────────────────
//
// Column layout (0-indexed):
//   0  Date                   metadata
//   1  Cutting Machine Number  metadata
//   2  Cutting Series M/Y      metadata
//   3  Cutting Machine Operators metadata
//   4  Start Time              metadata
//   5  Finish Time             metadata
//   6  Passenger               PCR
//   7  4 X 4                   PCR
//   8  Motorcycle              PCR
//   9  Light Commercial        PCR
//  10  Light Commercial T      RADIALS
//  11  Heavy Commercial T      RADIALS   (internal: radials_heavy_commercial_t)
//  12  Agricultural T          NYLONS    (internal: nylons_agricultural_t)
//  13  Heavy Commercial T      NYLONS    (internal: nylons_heavy_commercial_t)
//  14  Total Number of Nylon T NYLONS    (= col12 + col13)
//  15  Tread Cuts              extra (not in manual workbook group structure)
//  16  Raw Text                extra (for QA/debug)
//
// Category → business column mapping:
//   passenger         → col 6  (PCR Passenger)
//   fourx4            → col 7  (PCR 4 X 4)
//   motorcycle        → col 8  (PCR Motorcycle)
//   light_commercial  → col 9  (PCR Light Commercial)
//   tread_lc          → col 10 (RADIALS Light Commercial T tread)
//   heavy_commercial_t→ col 11 (RADIALS Heavy Commercial T tyre)
//   agricultural_t    → col 12 (NYLONS Agricultural T tyre)
//   tread_hc          → col 13 (NYLONS Heavy Commercial T tread)
//   tread_agri        → col 15 (Agri Tread / old-format Tread Cuts)
//   treads            → col 15 fallback (old-format total tread cuts)
//   unknown_type / "" → all blank

import { dateToStr, dateSortKey, timeToDecimal, formatTime } from "./helpers.js";

export const CUTTING_GROUP_HEADER_ROW = [
  "", "", "", "", "", "",          // cols 0-5  metadata (no group label)
  "PCR", "", "", "",               // cols 6-9  PCR group (merge across 4)
  "RADIALS", "",                   // cols 10-11 RADIALS group (merge across 2)
  "NYLONS", "", "",                // cols 12-14 NYLONS group (merge across 3)
  "", "",                          // cols 15-16 extra (outside groups)
];

export const CUTTING_FIELD_HEADER_ROW = [
  "Date", "Cutting Machine Number", "Cutting Series M/Y", "Cutting Machine Operators",
  "Start Time", "Finish Time",
  "Passenger", "4 X 4", "Motorcycle", "Light Commercial",
  "Light Commercial T", "Heavy Commercial T",
  "Agricultural T", "Heavy Commercial T", "Total Number of Nylon T",
  "Tread Cuts", "Raw Text",
];

// Merge ranges for the group header row (row index 0, 0-based):
export const CUTTING_GROUP_MERGES = [
  { s: { r: 0, c: 6  }, e: { r: 0, c: 9  } },  // PCR
  { s: { r: 0, c: 10 }, e: { r: 0, c: 11 } },  // RADIALS
  { s: { r: 0, c: 12 }, e: { r: 0, c: 14 } },  // NYLONS
];

export function cuttingSheetRows(records) {
  const sorted = [...records].sort((a, b) => {
    const dk = dateSortKey(a.date) - dateSortKey(b.date);
    if (dk !== 0) return dk;
    const tk = timeToDecimal(a.startTime) - timeToDecimal(b.startTime);
    if (tk !== 0) return tk;
    return a.cmNumber.localeCompare(b.cmNumber);
  });

  const rows = [CUTTING_GROUP_HEADER_ROW, CUTTING_FIELD_HEADER_ROW];
  for (const r of sorted) {
    // Helper: render a numeric field — null/undefined → blank cell
    const v = (field) => (r[field] !== null && r[field] !== undefined) ? r[field] : "";

    // col 12 — Agricultural T (NYLONS): Agri tyre quantity
    const col12 = r.agricultural_t ?? null;
    // col 13 — Heavy Commercial T (NYLONS): HC tread quantity
    const col13 = r.tread_hc ?? null;
    // col 14 — Total Nylon T = col12 + col13 (only when at least one is present)
    const col14 = (col12 !== null || col13 !== null)
      ? ((col12 ?? 0) + (col13 ?? 0))
      : null;
    // col 15 — Agri tread (new format) or total tread cuts (old format)
    const col15 = r.tread_agri ?? r.treads ?? null;

    rows.push([
      dateToStr(r.date),                          // 0  Date
      r.cmNumber,                                 // 1  CM Number
      r.series,                                   // 2  Series
      r.operator || "",                           // 3  Operators
      r.startTime ? formatTime(r.startTime) : "", // 4  Start
      r.finishTime ? formatTime(r.finishTime) : "", // 5  Finish
      // PCR
      v("passenger"),                             // 6  Passenger
      v("fourx4"),                                // 7  4 X 4
      v("motorcycle"),                            // 8  Motorcycle
      v("light_commercial"),                      // 9  Light Commercial
      // RADIALS
      v("tread_lc"),                              // 10 Light Commercial T (tread)
      v("heavy_commercial_t"),                    // 11 Heavy Commercial T (radials)
      // NYLONS
      col12 !== null ? col12 : "",               // 12 Agricultural T (tyre)
      col13 !== null ? col13 : "",               // 13 Heavy Commercial T (nylons tread)
      col14 !== null ? col14 : "",               // 14 Total Number of Nylon T
      // Extra
      col15 !== null ? col15 : "",               // 15 Agri Tread (new format) / Tread Cuts (old format)
      r.rawMessage || "",                        // 16 Raw Text
    ]);
  }
  return rows;
}
