// ─── Cutting Sheet Builder ────────────────────────────────────────────────────
//
// Column layout (0-indexed):
//   0  Date                       metadata
//   1  Cutting Machine Number     metadata
//   2  Cutting Series M/Y         metadata
//   3  Cutting Machine Operators  metadata
//   4  Cutting Machine Assistants metadata
//   5  Start Time                 metadata
//   6  Finish Time                metadata
//   7  Light Commercial           RADIALS  (internal: radialsLC)
//   8  Heavy Commercial           RADIALS  (internal: radialsHC)
//   9  Agricultural               RADIALS  (internal: radialsAgri)
//  10  Agri Treads                RADIALS  (internal: radialsAgriTreads)
//  11  Light Commercial           NYLONS   (internal: nylonsLC)
//  12  Raw Text                   extra (for QA/debug)
//
// Category → business column mapping:
//   radialsLC         → col 7  (RADIALS Light Commercial)
//   radialsHC         → col 8  (RADIALS Heavy Commercial)
//   radialsAgri       → col 9  (RADIALS Agricultural)
//   radialsAgriTreads → col 10 (RADIALS Agri Treads)
//   nylonsLC          → col 11 (NYLONS Light Commercial)
//   unknown_type / "" → all blank

import { dateToStr, dateSortKey, timeToDecimal, formatTime } from "./helpers.js";

export const CUTTING_GROUP_HEADER_ROW = [
  "", "", "", "", "", "", "",      // cols 0-6  metadata (no group label)
  "RADIALS", "", "", "",           // cols 7-10 RADIALS group (merge across 4)
  "NYLONS",                        // col 11   NYLONS group
  "",                              // col 12   extra (outside groups)
];

export const CUTTING_FIELD_HEADER_ROW = [
  "Date", "Cutting Machine Number", "Cutting Series M/Y", "Operator", "Assistant",
  "Start Time", "Finish Time",
  "Light Commercial", "Heavy Commercial", "Agricultural", "Agri Treads",
  "Light Commercial", "Raw Text",
];

// Merge ranges for the group header row (row index 0, 0-based):
export const CUTTING_GROUP_MERGES = [
  { s: { r: 0, c: 7 }, e: { r: 0, c: 10 } },   // RADIALS
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
  const ambiguousIndices = new Set();
  const duplicateTyreIndices = new Set();
  for (let i = 0; i < sorted.length; i += 1) {
    const r = sorted[i];
    // Helper: render a numeric field — null/undefined → blank cell
    const v = (field) => (r[field] != null && r[field] !== 0) ? r[field] : "";

    rows.push([
      dateToStr(r.date),                            // 0  Date
      r.cmNumber,                                   // 1  CM Number
      r.series,                                     // 2  Series
      r.operator || "",                             // 3  Operators
      r.cuttingAssistant || "",                     // 4  Assistants
      r.startTime ? formatTime(r.startTime) : "",   // 5  Start
      r.finishTime ? formatTime(r.finishTime) : "", // 6  Finish
      // RADIALS
      v("radialsLC"),                               // 7  Light Commercial
      v("radialsHC"),                               // 8  Heavy Commercial
      v("radialsAgri"),                             // 9  Agricultural
      v("radialsAgriTreads"),                       // 10 Agri Treads
      // NYLONS
      v("nylonsLC"),                                // 11 Light Commercial
      // Extra
      r.rawMessage || "",                           // 12 Raw Text
    ]);
    if (r._unresolvedType) unresolvedIndices.add(i);
    if (r._inferredType) inferredIndices.add(i);
    if (r._ambiguousLine) ambiguousIndices.add(i);
    if (r._duplicateTyreType) duplicateTyreIndices.add(i);
  }
  return { rows, unresolvedIndices, inferredIndices, ambiguousIndices, duplicateTyreIndices };
}
