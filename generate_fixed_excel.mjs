/**
 * generate_fixed_excel.mjs
 * Reads the WhatsApp cutting chat, parses it with the corrected logic
 * (using WhatsApp timestamp date instead of operator-typed body date),
 * and writes BPR_Cutting_Data_FIXED.xlsx with one sheet per month.
 *
 * Run: node generate_fixed_excel.mjs
 */

import { readFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";
import ExcelJS from "exceljs";

// ─── Shared Modules ────────────────────────────────────────────────────────────

import { dateToStr, monthKey, monthLabel } from "./src/helpers.js";
import { parseCuttingMessages } from "./src/cuttingParser.js";
import { cuttingSheetRows, CUTTING_GROUP_MERGES } from "./src/cuttingSheetBuilder.js";

// ─── Paths ────────────────────────────────────────────────────────────────────

const __dirname = dirname(fileURLToPath(import.meta.url));
const CHAT_PATH = resolve(__dirname, "../WhatsApp_chats/WhatsApp Chat with Blue pyramid recycling- Cut Tyres.txt");
const OUT_PATH = resolve(__dirname, "../BPR_Cutting_Data_FIXED.xlsx");

function groupRecordsByMonth(records) {
  const grouped = {};
  for (const record of records) {
    const key = monthKey(record.date);
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(record);
  }
  return grouped;
}

// ─── Main ─────────────────────────────────────────────────────────────────────

console.log("Reading chat file…");
const chatText = readFileSync(CHAT_PATH, "utf8");

console.log("Parsing messages…");
const { records } = parseCuttingMessages(chatText);
console.log(`  → ${records.length} records parsed`);

// Group by month
const byMonth = groupRecordsByMonth(records);

const sortedKeys = Object.keys(byMonth).sort();
console.log(`  → ${sortedKeys.length} months: ${sortedKeys.map(monthLabel).join(", ")}`);

// Print per-month counts (useful for validation)
console.log("\nRows per month:");
for (const key of sortedKeys) {
  console.log(`  ${monthLabel(key)}: ${byMonth[key].length} rows`);
}

// ─── Validation: compare Jan 2026 against TEST file ──────────────────────────
// Count rows per date in Jan 2026
const jan2026 = byMonth["2026-01"] || [];
const janByDate = {};
for (const r of jan2026) {
  const d = dateToStr(r.date);
  janByDate[d] = (janByDate[d] || 0) + 1;
}

console.log("\nJan 2026 rows per date (fixed):");
for (const [d, c] of Object.entries(janByDate).sort()) {
  console.log(`  ${d}: ${c}`);
}

// ─── Write Excel ──────────────────────────────────────────────────────────────
console.log("\nWriting Excel…");
const wb = new ExcelJS.Workbook();

for (const key of sortedKeys) {
  const rows = cuttingSheetRows(byMonth[key]);
  const ws = wb.addWorksheet(monthLabel(key));

  rows.forEach((row) => ws.addRow(row));

  for (const merge of CUTTING_GROUP_MERGES) {
    ws.mergeCells(
      merge.s.r + 1,
      merge.s.c + 1,
      merge.e.r + 1,
      merge.e.c + 1
    );
  }
}
await wb.xlsx.writeFile(OUT_PATH);
console.log(`\nDone → ${OUT_PATH}`);
