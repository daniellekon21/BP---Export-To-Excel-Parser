// ─── Cutting Sheet Builder ────────────────────────────────────────────────────
//
// Column layout (0-indexed):
//   0  Date                    metadata
//   1  Cutting Machine Number  metadata
//   2  Cutting Series M/Y      metadata
//   3  Cutting Machine Operators metadata
//   4  Start Time              metadata
//   5  Finish Time             metadata
//   6  Light Commercial        RADIALS  (internal: radialsLC)
//   7  Heavy Commercial        RADIALS  (internal: radialsHC)
//   8  Agricultural            RADIALS  (internal: radialsAgri)
//   9  Agri Treads             RADIALS  (internal: radialsAgriTreads)
//  10  Light Commercial        NYLONS   (internal: nylonsLC)
//  11  Raw Text                extra (for QA/debug)
//
// Category → business column mapping:
//   radialsLC         → col 6  (RADIALS Light Commercial)
//   radialsHC         → col 7  (RADIALS Heavy Commercial)
//   radialsAgri       → col 8  (RADIALS Agricultural)
//   radialsAgriTreads → col 9  (RADIALS Agri Treads)
//   nylonsLC          → col 10 (NYLONS Light Commercial)
//   unknown_type / "" → all blank

import { dateToStr, dateSortKey, timeToDecimal, formatTime } from "./helpers.js";

export const CUTTING_GROUP_HEADER_ROW = [
  "", "", "", "", "", "",          // cols 0-5  metadata (no group label)
  "RADIALS", "", "", "",           // cols 6-9  RADIALS group (merge across 4)
  "NYLONS",                        // col 10   NYLONS group
  "",                              // col 11   extra (outside groups)
];

export const CUTTING_FIELD_HEADER_ROW = [
  "Date", "Cutting Machine Number", "Cutting Series M/Y", "Cutting Machine Operators",
  "Start Time", "Finish Time",
  "Light Commercial", "Heavy Commercial", "Agricultural", "Agri Treads",
  "Light Commercial", "Raw Text",
];

// Merge ranges for the group header row (row index 0, 0-based):
export const CUTTING_GROUP_MERGES = [
  { s: { r: 0, c: 6 }, e: { r: 0, c: 9 } },    // RADIALS
];

export function cuttingSheetRows(records) {
  const sorted = [...records].sort((a, b) => {
    const dk = dateSortKey(a.date) - dateSortKey(b.date);
    if (dk !== 0) return dk;
    const tk = timeToDecimal(a.startTime) - timeToDecimal(b.startTime);
    if (tk !== 0) return tk;
    const aOrder = Number.isInteger(a?._parseOrder) ? a._parseOrder : null;
    const bOrder = Number.isInteger(b?._parseOrder) ? b._parseOrder : null;
    if (aOrder !== null && bOrder !== null && aOrder !== bOrder) return aOrder - bOrder;
    return a.cmNumber.localeCompare(b.cmNumber);
  });

  const rows = [CUTTING_GROUP_HEADER_ROW, CUTTING_FIELD_HEADER_ROW];
  const unresolvedIndices = new Set();
  const inferredIndices = new Set();
  for (let i = 0; i < sorted.length; i += 1) {
    const r = sorted[i];
    // Helper: render a numeric field — null/undefined → blank cell
    const v = (field) => (r[field] != null && r[field] !== 0) ? r[field] : "";

    rows.push([
      dateToStr(r.date),                          // 0  Date
      r.cmNumber,                                 // 1  CM Number
      r.series,                                   // 2  Series
      r.operator || "",                           // 3  Operators
      r.startTime ? formatTime(r.startTime) : "", // 4  Start
      r.finishTime ? formatTime(r.finishTime) : "", // 5  Finish
      // RADIALS
      v("radialsLC"),                             // 6  Light Commercial
      v("radialsHC"),                             // 7  Heavy Commercial
      v("radialsAgri"),                           // 8  Agricultural
      v("radialsAgriTreads"),                     // 9  Agri Treads
      // NYLONS
      v("nylonsLC"),                              // 10 Light Commercial
      // Extra
      r.rawMessage || "",                        // 11 Raw Text
    ]);
    if (r._unresolvedType) unresolvedIndices.add(i);
    if (r._inferredType) inferredIndices.add(i);
  }
  return { rows, unresolvedIndices, inferredIndices };
}
