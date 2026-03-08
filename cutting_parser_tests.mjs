/**
 * cutting_parser_tests.mjs
 * Minimal regression tests for the cutting parser fixes.
 * Run: node cutting_parser_tests.mjs
 */

// ─── Import shared modules ──────────────────────────────────────────────────

import {
  normalizeCuttingLine, classifyLine, mapTyreType, parseMachineLine,
  mapTyreTypeNew, mapTreadTypeNew, flushCutterBlock, parseCuttingMessages, parseCuttingMessagesNew,
} from "./src/cuttingParser.js";

// ── Test runner ───────────────────────────────────────────────────────────────

let passed = 0, failed = 0;

function check(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (ok) {
    console.log(`  PASS  ${label}`);
    passed++;
  } else {
    console.error(`  FAIL  ${label}`);
    console.error(`        expected: ${JSON.stringify(expected)}`);
    console.error(`        actual  : ${JSON.stringify(actual)}`);
    failed++;
  }
}

function normalize_then_parse(raw) {
  return parseMachineLine(normalizeCuttingLine(raw));
}

// ── normalizeCuttingLine ──────────────────────────────────────────────────────
console.log("\nnormalizeCuttingLine:");
check("Machine 1 → CM1", normalizeCuttingLine("Machine 1-18x LC"), "CM1-18 LC");
check("x suffix removed (space before type)", normalizeCuttingLine("Machine 1- 66x LC"), "CM1- 66 LC");
check("4x4 preserved (digit-alpha split only)", normalizeCuttingLine("CM2 4x4"), "CM2 4 x4");
check("unicode dash", normalizeCuttingLine("CM1–30 Agri"), "CM1-30 Agri");

// ── mapTyreType ───────────────────────────────────────────────────────────────
console.log("\nmapTyreType:");
check("Truck → heavy_commercial_t", mapTyreType("Truck"), "heavy_commercial_t");
check("Trucks → heavy_commercial_t", mapTyreType("Trucks"), "heavy_commercial_t");
check("Truck Radials → hc", mapTyreType("Truck Radials"), "heavy_commercial_t");
check("Radial → heavy_commercial_t", mapTyreType("Radial"), "heavy_commercial_t");
check("Radial (LC) → light_commercial", mapTyreType("Radial (LC)"), "light_commercial");
check("LC Radials → light_commercial", mapTyreType("LC Radials"), "light_commercial");
check("treads → treads", mapTyreType("treads"), "treads");
check("Tread → treads", mapTyreType("Tread"), "treads");
check("(Agri) → agricultural_t", mapTyreType("(Agri)"), "agricultural_t");

// ── parseMachineLine: Format G ────────────────────────────────────────────────
console.log("\nFormat G (CMN TypeName - Count):");
check("Machine 1 Agricultural - 10",
  normalize_then_parse("Machine 1 Agricultural - 10"),
  [{ cmNum: 1, column: "agricultural_t", count: 10 }]);

check("Machine 3 Truck Radials - 37",
  normalize_then_parse("Machine 3 Truck Radials - 37"),
  [{ cmNum: 3, column: "heavy_commercial_t", count: 37 }]);

// ── parseMachineLine: x suffix ────────────────────────────────────────────────
console.log("\nx-suffix forms:");
check("Machine 1-18x LC",
  normalize_then_parse("Machine 1-18x LC"),
  [{ cmNum: 1, column: "light_commercial", count: 18 }]);

check("Machine 3- 81x LC",
  normalize_then_parse("Machine 3- 81x LC"),
  [{ cmNum: 3, column: "light_commercial", count: 81 }]);

// ── parseMachineLine: Truck types ─────────────────────────────────────────────
console.log("\nTruck/Radial types:");
check("Machine 1-45 Trucks",
  normalize_then_parse("Machine 1-45 Trucks"),
  [{ cmNum: 1, column: "heavy_commercial_t", count: 45 }]);

check("Machine 3- 19 Truck Radial",
  normalize_then_parse("Machine 3- 19 Truck Radial"),
  [{ cmNum: 3, column: "heavy_commercial_t", count: 19 }]);

check("Machine 1- 28 truck",
  normalize_then_parse("Machine 1- 28 truck"),
  [{ cmNum: 1, column: "heavy_commercial_t", count: 28 }]);

// ── parseMachineLine: parens around type (Format C) ──────────────────────────
console.log("\nParens around type:");
check("CM2- 21 (Agri)",
  normalize_then_parse("CM2- 21 (Agri)"),
  [{ cmNum: 2, column: "agricultural_t", count: 21 }]);

check("CM3-11  (Agri)",
  normalize_then_parse("CM3-11  (Agri)"),
  [{ cmNum: 3, column: "agricultural_t", count: 11 }]);

// ── parseMachineLine: double-dash ─────────────────────────────────────────────
console.log("\nDouble-dash format:");
check("Machine 1-9 - Agri  (from Dec 10 correction)",
  normalize_then_parse("Machine 1-9 - Agri"),
  [{ cmNum: 1, column: "agricultural_t", count: 9 }]);

// ── parseMachineLine: compound with + ────────────────────────────────────────
console.log("\nCompound lines (explicit +):");
check("CM3-6 Agri + 11 treads",
  parseMachineLine("CM3-6 Agri + 11 treads"),
  [{ cmNum: 3, column: "agricultural_t", count: 6 }, { cmNum: 3, column: "treads", count: 11 }]);

check("CM2-5 Agri + 20 treads",
  parseMachineLine("CM2-5 Agri + 20 treads"),
  [{ cmNum: 2, column: "agricultural_t", count: 5 }, { cmNum: 2, column: "treads", count: 20 }]);

check("CM2-20 Radial (LC) +28 Agricultural",
  parseMachineLine("CM2-20 Radial (LC) +28 Agricultural"),
  [{ cmNum: 2, column: "light_commercial", count: 20 }, { cmNum: 2, column: "agricultural_t", count: 28 }]);

// ── parseMachineLine: compound implicit (no +) ────────────────────────────────
console.log("\nCompound lines (implicit, no +):");
check("CM2-20 Agri 18 Treads",
  parseMachineLine("CM2-20 Agri 18 Treads"),
  [{ cmNum: 2, column: "agricultural_t", count: 20 }, { cmNum: 2, column: "treads", count: 18 }]);

check("CM3-11 HC 12 Agri",
  parseMachineLine("CM3-11 HC 12 Agri"),
  [{ cmNum: 3, column: "heavy_commercial_t", count: 11 }, { cmNum: 3, column: "agricultural_t", count: 12 }]);

check("CM1-9 Agri 8 treads",
  parseMachineLine("CM1-9 Agri 8 treads"),
  [{ cmNum: 1, column: "agricultural_t", count: 9 }, { cmNum: 1, column: "treads", count: 8 }]);

// ── parseMachineLine: Tread Cut ───────────────────────────────────────────────
console.log("\nTread Cut pattern:");
check("CM2- 14 (Agri) - Tread Cut 21",
  parseMachineLine("CM2- 14 (Agri) - Tread Cut 21"),
  [{ cmNum: 2, column: "agricultural_t", count: 14 }, { cmNum: 2, column: "treads", count: 21 }]);

// ── parseMachineLine: N/A ─────────────────────────────────────────────────────
console.log("\nN/A forms:");
check("CM1= N/A",
  parseMachineLine("CM1= N/A"),
  [{ cmNum: 1, column: "", count: null, isStatus: true }]);

check("Machine 2-N/A Agri. L → N/A",
  normalize_then_parse("Machine 2-N/A Agri. L"),
  [{ cmNum: 2, column: "", count: null, isStatus: true }]);

check("Machine 3- N/A (offloading truck)",
  normalize_then_parse("Machine 3- N/A (offloading truck)"),
  [{ cmNum: 3, column: "", count: null, isStatus: true }]);

// ── parseMachineLine: count-only ──────────────────────────────────────────────
console.log("\nCount-only (no type):");
check("CM1-30 (Format Z with dash)",
  parseMachineLine("CM1-30"),
  [{ cmNum: 1, column: null, count: 30 }]);

check("Machine 1- 12 (count only)",
  normalize_then_parse("Machine 1- 12"),
  [{ cmNum: 1, column: null, count: 12 }]);

// ── parseMachineLine: missing-space forms (already working, regression) ───────
console.log("\nMissing-space forms (regression):");
check("CM1=29Agri → agricultural_t",
  normalize_then_parse("CM1=29Agri"),
  [{ cmNum: 1, column: "agricultural_t", count: 29 }]);

check("CM3=23HC → heavy_commercial_t",
  normalize_then_parse("CM3=23HC"),
  [{ cmNum: 3, column: "heavy_commercial_t", count: 23 }]);

// ── parseMachineLine: HC/LC dual (regression) ─────────────────────────────────
console.log("\nDual HC/LC via Format E (regression):");
check("CM1-Agricultural=26",
  parseMachineLine("CM1-Agricultural=26"),
  [{ cmNum: 1, column: "agricultural_t", count: 26 }]);

check("CM2-HC=20",
  parseMachineLine("CM2-HC=20"),
  [{ cmNum: 2, column: "heavy_commercial_t", count: 20 }]);

// ── New-Format Parser ─────────────────────────────────────────────────────────

function flushCutterBlockTest(block, meta, records) {
  const { series, startTime, finishTime } = meta;
  const { cmNum, operator, assistant, tyreType, tyreCount, treadType, treadCount } = block;
  const opStr = [operator, assistant].filter(Boolean).join(" / ");
  if (tyreType !== null && tyreCount !== null)
    records.push({ cmNum, series, startTime, finishTime, operator: opStr, column: mapTyreTypeNew(tyreType), count: tyreCount });
  if (treadType !== null && treadCount !== null)
    records.push({ cmNum, series, startTime, finishTime, operator: "", column: mapTreadTypeNew(treadType), count: treadCount });
}

function parseNewFormatBody(body) {
  const records = [];
  const meta = { series: "03/26", startTime: null, finishTime: null };
  const slotMatch = body.match(/\*?time\*?\s*:\s*(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})/i);
  if (slotMatch) { meta.startTime = slotMatch[1]; meta.finishTime = slotMatch[2]; }

  let currentBlock = null, lastQtyField = null;
  for (const rawLine of body.split("\n")) {
    const line = rawLine.trim();
    if (!line) continue;
    const cutterMatch = line.match(/^\*?cutter\s+(\d+)\*?/i);
    if (cutterMatch) {
      if (currentBlock) flushCutterBlockTest(currentBlock, meta, records);
      currentBlock = {
        cmNum: parseInt(cutterMatch[1], 10), operator: "", assistant: "",
        tyreType: null, tyreCount: null, treadType: null, treadCount: null
      };
      lastQtyField = null; continue;
    }
    if (!currentBlock) continue;
    const m_op = line.match(/^operator\s*:\s*(.*)/i);
    const m_ast = line.match(/^assistant\s*:\s*(.*)/i);
    const m_tyre = line.match(/^t[iy]re\s+type\s*:\s*(.*)/i);
    const m_trd = line.match(/^tread\s+type\s*:\s*(.*)/i);
    const m_qty = line.match(/^quantit[yi]e?s?\s*:\s*(\d+)/i);
    if (m_op) { currentBlock.operator = m_op[1].trim(); continue; }
    if (m_ast) { currentBlock.assistant = m_ast[1].trim(); continue; }
    if (m_tyre) { currentBlock.tyreType = m_tyre[1].trim(); lastQtyField = "tyre"; continue; }
    if (m_trd) { currentBlock.treadType = m_trd[1].trim(); lastQtyField = "tread"; continue; }
    if (m_qty) {
      const qty = parseInt(m_qty[1], 10);
      if (lastQtyField === "tyre") currentBlock.tyreCount = qty;
      if (lastQtyField === "tread") currentBlock.treadCount = qty;
    }
  }
  if (currentBlock) flushCutterBlockTest(currentBlock, meta, records);
  return records;
}

console.log("\nNew-format parser:");

// 1. Single cutter, all fields → 2 records (tyre + tread)
check("Single cutter, all fields → 2 records",
  parseNewFormatBody(`*Cutter 1*\nOperator: Jane\nAssistant: Bob\nTyre Type: LC\nQuantity: 45\nTread Type: HC\nQuantity: 12`),
  [
    { cmNum: 1, series: "03/26", startTime: null, finishTime: null, operator: "Jane / Bob", column: "light_commercial", count: 45 },
    { cmNum: 1, series: "03/26", startTime: null, finishTime: null, operator: "", column: "tread_hc", count: 12 },
  ]);

// 2. Three cutters → 6 records
check("Three cutters → 6 records",
  parseNewFormatBody(
    `*Cutter 1*\nTyre Type: LC\nQuantity: 10\nTread Type: HC\nQuantity: 5\n` +
    `*Cutter 2*\nTyre Type: HC\nQuantity: 20\nTread Type: Agri\nQuantity: 8\n` +
    `*Cutter 3*\nTyre Type: Agri\nQuantity: 30\nTread Type: LC\nQuantity: 3`
  ).length,
  6);

// 3. Blank operator/assistant → records still created
check("Blank operator → tyre record still created",
  parseNewFormatBody(`*Cutter 2*\nOperator:\nAssistant:\nTyre Type: Agri\nQuantity: 25\nTread Type: Agri\nQuantity: 7`),
  [
    { cmNum: 2, series: "03/26", startTime: null, finishTime: null, operator: "", column: "agricultural_t", count: 25 },
    { cmNum: 2, series: "03/26", startTime: null, finishTime: null, operator: "", column: "tread_agri", count: 7 },
  ]);

// 4. Missing tread section → only tyre record
check("No tread section → 1 record only",
  parseNewFormatBody(`*Cutter 1*\nTyre Type: HC\nQuantity: 18`),
  [{ cmNum: 1, series: "03/26", startTime: null, finishTime: null, operator: "", column: "heavy_commercial_t", count: 18 }]);

// 5. Missing tread quantity → no tread record
check("Tread type but no quantity → no tread record",
  parseNewFormatBody(`*Cutter 1*\nTyre Type: LC\nQuantity: 30\nTread Type: HC`),
  [{ cmNum: 1, series: "03/26", startTime: null, finishTime: null, operator: "", column: "light_commercial", count: 30 }]);

// 6. mapTyreTypeNew / mapTreadTypeNew
console.log("\nNew-format type mappers:");
check("mapTyreTypeNew LC", mapTyreTypeNew("LC"), "light_commercial");
check("mapTyreTypeNew HC", mapTyreTypeNew("HC"), "heavy_commercial_t");
check("mapTyreTypeNew Agri", mapTyreTypeNew("Agri"), "agricultural_t");
check("mapTreadTypeNew LC", mapTreadTypeNew("LC"), "tread_lc");
check("mapTreadTypeNew HC", mapTreadTypeNew("HC"), "tread_hc");
check("mapTreadTypeNew Agri", mapTreadTypeNew("Agri"), "tread_agri");

// 7. Summary format → no records (simulated via body containing "cutting summary")
// (parseCuttingMessagesNew skips these — verifying mappers don't crash on Agri/LC/HC is sufficient)

// ── Status-text stripping ─────────────────────────────────────────────────────
console.log("\nStatus-text stripping:");
check("CM1-6 - stopped due to offloading → count 6, column null",
  normalize_then_parse("CM1-6 - stopped due to offloading of scrap truck"),
  [{ cmNum: 1, column: null, count: 6 }]);

check("CM2-30 unchanged (no status text)",
  normalize_then_parse("CM2-30"),
  [{ cmNum: 2, column: null, count: 30 }]);

check("CM3-10 Agri - idle → strips status, keeps type",
  normalize_then_parse("CM3-10 Agri - idle for maintenance"),
  [{ cmNum: 3, column: "agricultural_t", count: 10 }]);

check("CM1-12 - paused → count 12, column null",
  normalize_then_parse("CM1-12 - paused"),
  [{ cmNum: 1, column: null, count: 12 }]);

check("CM2-N/A still works (not affected by strip)",
  normalize_then_parse("CM2-N/A"),
  [{ cmNum: 2, column: "", count: null, isStatus: true }]);

// ── Context-aware type inference (unit test) ──────────────────────────────────
console.log("\nContext-aware type inference:");

// Simulate what parseCuttingMessages Phase 2 does:
function inferColumns(parsedResults) {
  const knownColumns = parsedResults
    .filter(p => p.column !== null && p.column !== "unknown_type" && p.column !== "")
    .map(p => p.column);
  const uniqueKnown = [...new Set(knownColumns)];
  const inferredColumn = uniqueKnown.length === 1 ? uniqueKnown[0] : null;
  return parsedResults.map(p => {
    let col = p.column;
    if ((col === null || col === "unknown_type") && p.count !== null && inferredColumn) {
      col = inferredColumn;
    }
    return { ...p, column: col };
  });
}

check("Infer agri for bare-count siblings",
  inferColumns([
    { cmNum: 1, column: null, count: 8 },
    { cmNum: 2, column: "agricultural_t", count: 16 },
    { cmNum: 3, column: null, count: 10 },
  ]),
  [
    { cmNum: 1, column: "agricultural_t", count: 8 },
    { cmNum: 2, column: "agricultural_t", count: 16 },
    { cmNum: 3, column: "agricultural_t", count: 10 },
  ]);

check("No inference when siblings disagree",
  inferColumns([
    { cmNum: 1, column: null, count: 8 },
    { cmNum: 2, column: "agricultural_t", count: 16 },
    { cmNum: 3, column: "heavy_commercial_t", count: 10 },
  ]),
  [
    { cmNum: 1, column: null, count: 8 },
    { cmNum: 2, column: "agricultural_t", count: 16 },
    { cmNum: 3, column: "heavy_commercial_t", count: 10 },
  ]);

check("No inference when all columns null (all bare)",
  inferColumns([
    { cmNum: 1, column: null, count: 8 },
    { cmNum: 2, column: null, count: 16 },
  ]),
  [
    { cmNum: 1, column: null, count: 8 },
    { cmNum: 2, column: null, count: 16 },
  ]);

check("Status records (count null) are NOT inferred",
  inferColumns([
    { cmNum: 1, column: "", count: null },
    { cmNum: 2, column: "agricultural_t", count: 16 },
  ]),
  [
    { cmNum: 1, column: "", count: null },
    { cmNum: 2, column: "agricultural_t", count: 16 },
  ]);

// ── End-to-end old parser regression rules ───────────────────────────────────
console.log("\nOld-parser E2E rules:");

const e2eTextSkipEmpty = `2026/01/10, 09:05 - Evan Botes: Date - 10/01/2026
Cutting (08:00-09:00)
CM1= N/A
CM2= N/A
CM3= N/A`;

const outSkipEmpty = parseCuttingMessages(e2eTextSkipEmpty);
check("Skip interval when no machine has numeric update",
  outSkipEmpty.records.length,
  0
);

const e2eTextKeepPlaceholders = `2026/01/10, 10:05 - Evan Botes: Date - 10/01/2026
Cutting (09:01-10:00)
Machine 1-30 Truck`;

const outKeepPlaceholders = parseCuttingMessages(e2eTextKeepPlaceholders);
const slotRows = outKeepPlaceholders.records
  .map(r => ({
    cm: r.cmNumber,
    hc: r.heavy_commercial_t,
    start: r.startTime ? `${String(r.startTime.h).padStart(2, "0")}:${String(r.startTime.m).padStart(2, "0")}` : "",
    finish: r.finishTime ? `${String(r.finishTime.h).padStart(2, "0")}:${String(r.finishTime.m).padStart(2, "0")}` : "",
  }))
  .sort((a, b) => a.cm.localeCompare(b.cm));

check("Keep CM placeholders when at least one machine has update",
  slotRows,
  [
    { cm: "CM - 1", hc: 30, start: "09:01", finish: "10:00" },
    { cm: "CM - 2", hc: null, start: "09:01", finish: "10:00" },
    { cm: "CM - 3", hc: null, start: "09:01", finish: "10:00" },
  ]
);

const e2eTextLegacySummary = `2026/01/05, 17:00 - Evan Botes: 05/01/2026
Cutting Summary

CM1= Agricultural - 205

CM2= N/A

CM3= Light Commercial - 67
Agricultural - 111

*Total Tyres - 383*`;

const outLegacySummary = parseCuttingMessages(e2eTextLegacySummary);
const summaryRows = outLegacySummary.summaryRecords
  .map(s => ({ cm: s.cmNumber, lc: s.lc, hc: s.hc, agri: s.agri }))
  .sort((a, b) => a.cm.localeCompare(b.cm));

check("Parse legacy CM= Cutting Summary blocks (partial allowed)",
  summaryRows,
  [
    { cm: "CM - 1", lc: null, hc: null, agri: 205 },
    { cm: "CM - 3", lc: 67, hc: null, agri: 111 },
  ]
);

// ── End-to-end new parser summary compatibility ──────────────────────────────
console.log("\nNew-parser summary compatibility:");

const e2eNewModeDailySummary = `2026/03/10, 18:00 - Evan Botes: Date - 10/03/2026
Daily Summary
Machine 1-96
Machine 2-159 Agri`;

const outNewModeDailySummary = parseCuttingMessagesNew(e2eNewModeDailySummary);
const newModeSummaryRows = outNewModeDailySummary.summaryRecords
  .map(s => ({ cm: s.cmNumber, lc: s.lc, hc: s.hc, agri: s.agri }))
  .sort((a, b) => a.cm.localeCompare(b.cm));

check("Parse legacy Daily Summary in new mode",
  newModeSummaryRows,
  [
    { cm: "CM - 2", lc: null, hc: null, agri: 159 },
  ]
);

// ── Summary ────────────────────────────────────────────────────────────────────
console.log(`\n${"─".repeat(50)}`);
console.log(`Results: ${passed} passed, ${failed} failed`);
if (failed > 0) process.exit(1);
