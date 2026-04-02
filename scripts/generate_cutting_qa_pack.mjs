// scripts/generate_cutting_qa_pack.mjs
//
// Generates a QA pack for both cutting parsers:
//   - parseCuttingMessages    (old / freeform format)
//   - parseCuttingMessagesNew (new / structured format)
//
// Outputs:
//   qa_outputs/cutting_old_parsed.json
//   qa_outputs/cutting_new_parsed.json
//   qa_outputs/cutting_old_qa_report.txt
//   qa_outputs/cutting_new_qa_report.txt
//   qa_outputs/BPR_Cutting_QA_Old.xlsx
//   qa_outputs/BPR_Cutting_QA_New.xlsx
//   sample-data/cutting_old_qa_chat.txt
//   sample-data/cutting_new_qa_chat.txt

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseCuttingMessages, parseCuttingMessagesNew } from "../src/cuttingParser.js";
import { groupByMonth, monthLabel, dateToStr } from "../src/helpers.js";
import { cuttingSheetRows, CUTTING_GROUP_MERGES } from "../src/cuttingSheetBuilder.js";
import { baseStyles, styleHeaderRow, styleBodyRows, applyColumnWidths } from "../src/excel/excelCommon.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);
const root       = path.resolve(__dirname, "..");
const outDir     = path.resolve(root, "qa_outputs");
const sampleDir  = path.resolve(root, "sample-data");
fs.mkdirSync(outDir,    { recursive: true });
fs.mkdirSync(sampleDir, { recursive: true });

// ─── Helpers ──────────────────────────────────────────────────────────────────

function msg(ts, body, sender = "QA Bot") {
  return `${ts} - ${sender}: ${body}`;
}

function block(lines) {
  return lines.join("\n");
}

// ─── OLD-FORMAT FIXTURE MESSAGES ─────────────────────────────────────────────
//
// Each case is commented with what the parser SHOULD produce so a human
// reviewer can compare the report against expectations.
//
// Date range: Feb 2025 – Mar 2025 (matches real chat files for inference context)

const oldMessages = [

  // ── Case 1: Basic Agri + Agriculture (Oct-style spelling) ──────────────────
  // Expect: CM1=N/A  CM2=25 Agri  CM3=17 Agri
  msg("2025/02/03, 09:00", block([
    "03/02/2025",
    "08:00-09:00",
    "CM1-N/A",
    "CM2-25 Agriculture",
    "CM3-17 Agriculture",
  ])),

  // ── Case 2: Space before number, parenthesized type  ──────────────────────
  // Expect: CM1=0  CM2=21 Agri  CM3=16 Agri
  msg("2025/02/04, 12:00", block([
    "04/02/2025",
    "10:30-12:00",
    "CM1-0",
    "CM2- 21 (Agri)",
    "CM3-16 (Agri)",
  ])),

  // ── Case 3: Compound "Agri + treads" ──────────────────────────────────────
  // Expect: CM1=N/A  CM2=7 Agri +24 Treads  CM3=8 Agri +31 Treads
  msg("2025/02/05, 13:48", block([
    "05/02/2025",
    "12:30-13:30",
    "CM1-N/A",
    "CM2-7 Agri + 24 treads",
    "CM3-8 Agri + 31 treads",
  ])),

  // ── Case 4: Implicit compound "tyres treads" (no + sign) ──────────────────
  // Expect: CM2=16 Agri(inferred) + 36 Treads
  msg("2025/02/07, 11:15", block([
    "07/02/2025",
    "10:00-11:00",
    "CM1-N/A",
    "CM2-16 tyres 36 treads",
    "CM3-N/A",
  ])),

  // ── Case 5: Trailing period on type ("Agri.") ─────────────────────────────
  // Expect: CM2=5 Agri
  msg("2025/02/06, 15:39", block([
    "06/02/2025",
    "14:15-15:15",
    "CM2- 5 Agri.",
  ])),

  // ── Case 6: Tread Cut inline notation ────────────────────────────────────
  // Expect: CM2=14 Agri + 21 Treads  CM3=31 Agri
  msg("2025/02/04, 16:06", block([
    "04/02/2025",
    "14:00-16:00",
    "CM2- 14 (Agri) - Tread Cut 21",
    "CM3- 31  (Agri)",
  ])),

  // ── Case 7: All-zero counts ───────────────────────────────────────────────
  // Expect: CM1=0  CM2=0  CM3=0 (placeholders/0)
  msg("2025/02/03, 10:16", block([
    "03/02/2025",
    "08:00-09:00",
    "CM1-0",
    "CM2-0",
    "CM3-0",
  ])),

  // ── Case 8: No-space between count and type ("13Agriculture") ────────────
  // Expect: CM1=13 Agri  CM2=18 Agriculture  CM3=30 Agriculture
  msg("2025/10/06, 16:17", block([
    "06/10/2025",
    "15:00-16:00",
    "CM1- 13Agriculture",
    "CM2-18 Agriculture",
    "CM3-30 Agriculture",
  ])),

  // ── Case 9: Radial(LC) qualifier ─────────────────────────────────────────
  // Expect: CM1=21 LC  CM2=20 LC  CM3=30 Agri
  msg("2025/10/09, 09:09", block([
    "09/10/2025",
    "08:00-09:00",
    "CM1- 21 Radial(LC)",
    "CM2-20 Radial (LC)",
    "CM3-30 Agriculture",
  ])),

  // ── Case 10: Mixed Radial+Agri compound in same line ─────────────────────
  // Expect: CM2=20 LC + 28 Agri
  msg("2025/10/09, 10:11", block([
    "09/10/2025",
    "09:00-10:00",
    "CM1- 27 Radial(LC)",
    "CM2-20 Radial (LC) +28 Agricultural",
    "CM3-40 Agricultural",
  ])),

  // ── Case 11: Format D – parenthesised dual-type (HC)=N / (LC)=N ──────────
  // Expect: CM1=HC:17,LC:12  CM2=HC:6,LC:14  CM3=LC:24
  msg("2025/10/10, 14:11", block([
    "10/10/2025",
    "13:00-14:00",
    "CM1-(HC) =17",
    "            (LC) =12",
    "",
    "CM2-(HC)=06",
    "            (LC)=14",
    "",
    "CM3-(LC)=24",
  ])),

  // ── Case 12: Bare "LC" / "HC" labels (no "Radial" prefix) ────────────────
  // Expect: CM1=14 LC  CM2=12 LC  CM3=16 LC
  msg("2025/10/09, 16:11", block([
    "09/10/2025",
    "15:00-16:00",
    "CM1- 14 LC",
    "CM2-12  LC",
    "CM3-16  LC",
  ])),

  // ── Case 13: Cutting Summary – legacy "Machine N-Count (Type)" style ──────
  // Expect summaryRecords: CM1=N/A  CM2=19 Agri  total=19
  msg("2025/02/06, 17:45", block([
    "06/02/2025",
    "Cutting Summary",
    "",
    "Machine 1-N/A (Awaiting spares)",
    "",
    "Machine 2-19 Agri.",
    "",
    "Total cut- 19 Agri.",
  ])),

  // ── Case 14: Cutting Summary – "CMX= Count Type" style (Oct 2025) ─────────
  // Expect summaryRecords: CM1=119 Agri  CM2=148 Agri  CM3=176 Agri
  msg("2025/10/06, 17:39", block([
    "06/10/25",
    "Cutting Summary",
    "",
    "CM1= 119 Agri",
    "",
    "CM2= 148 Agri",
    "",
    "CM3= 176 Agri",
    "",
    "Total= 443 Agricultural Tyres",
  ])),

  // ── Case 15: Cutting Summary – Agri + Radials mixed (Oct 08 style) ────────
  // Expect summaryRecords: CM1=87 Agri + 20 LC  CM2=24 Agri + 99 LC  CM3=50 Agri + 137 LC
  msg("2025/10/08, 17:16", block([
    "08/10/25",
    "Cutting Summary",
    "",
    "CM1= 87 Agri",
    "",
    "CM2= 24 Agri",
    "",
    "CM3= 50 Agri",
    "",
    " *Total= 161 Agricultural Tyres*",
    " ",
    "CM1= 20 Radials",
    "",
    "CM2= 99 Radials",
    "",
    "CM3= 137 Radials",
    "",
    " *Total= 256 Radial Tyres*",
  ])),

  // ── Case 16: Cutting Summary – Machine-style with Agri + Treads ───────────
  // Expect summaryRecords: CM1=N/A  CM2=72 Agri + 213 Treads  CM3=59 Agri + 115 Treads
  msg("2025/02/05, 18:14", block([
    "05/02/2025",
    "Cutting Summary",
    "",
    "Machine 1-N/A (Awaiting spares)",
    "",
    "Machine 2-72 Agri + 213 Treads",
    "",
    "Machine 3-59 Agri + 115 Treads",
    "",
    "Total cut- 131 Tyres and 328 Treads",
  ])),

  // ── Case 17: Multi-interval in one message with different start times ──────
  // Expect two rows for CM3 (offsets) and one for CM1, one for CM2
  msg("2025/02/10, 10:18", block([
    "10/02/2025",
    "08:00-09:00",
    "",
    "CM1-9 Agri",
    "",
    "CM2-49 Treads",
    "",
    "08:30-09:30",
    "",
    "CM3-17 Agri",
  ])),

  // ── Case 18: Deleted / noise messages (should be ignored) ─────────────────
  msg("2025/02/10, 10:33", "This message was deleted"),
  msg("2025/02/10, 09:35", "This message was deleted"),
  msg("2025/02/10, 10:00", "<Media omitted>"),
  msg("2025/02/07, 09:00", "Good morning team!"),

  // ── Case 19: Operator roster message (should be ignored) ──────────────────
  msg("2025/10/07, 08:15", block([
    "07/10/2025",
    "",
    "CM1-Sanele and Lindani",
    "",
    "CM2-Ntando and Athulile",
    "",
    "CM3-Qiniso and Yamkela",
  ])),

  // ── Case 20: Standalone Treads Cut line (after CM line) ──────────────────
  // Expect: CM3=11 Agri + 8 Treads
  msg("2025/02/04, 14:07", block([
    "04/02/2025",
    "12:30 -14:00 (30 mins lunch break)",
    "CM1-0",
    "CM2- 21 (Agri)",
    "CM3-11  (Agri) -",
    "Treads Cut-8",
  ])),

  // ── Case 21: Message-edited marker should be stripped ────────────────────
  // Expect: CM2=14 Agri Tread:21  CM3=31 Agri
  msg("2025/02/04, 16:07", block([
    "04/02/2025",
    "12:30 -14:00 (30 mins lunch break)",
    "CM2- 14 (Agri) - Tread Cut 21 <This message was edited>",
    "CM3- 31  (Agri) <This message was edited>",
  ])),

  // ── Case 22: Future date in body – should fall back to timestamp date ──────
  msg("2025/02/11, 09:00", block([
    "11/02/2036",   // <- future, should be rejected; falls back to msg ts
    "08:00-09:00",
    "CM1-10 Agri",
    "CM2-12 Agri",
  ])),

];

// ─── NEW-FORMAT FIXTURE MESSAGES ─────────────────────────────────────────────
//
// Structured per-cutter-per-hour format introduced late 2025.

const newMessages = [

  // ── Case 1: Standard clean entry (Cutter 1/2/3, all fields present) ───────
  // Expect: CM-1 LC=18+tread12  CM-2 HC=10+tread8  CM-3 Agri=6+tread4
  msg("2025/11/03, 08:07", block([
    "03/11/2025",
    "",
    "Time: 07:00-08:00",
    "",
    "Cutter 1",
    "Operator: Sipho",
    "Assistant: Andile",
    "Tyre Type: LC",
    "Quantity: 18",
    "Tread Type: LC",
    "Quantity: 12",
    "",
    "Cutter 2",
    "Operator: Menzi",
    "Assistant: Philani",
    "Tyre Type: HC",
    "Quantity: 10",
    "Tread Type: HC",
    "Quantity: 8",
    "",
    "Cutter 3",
    "Operator: Bhekinkosi",
    "Assistant: Sanele",
    "Tyre Type: Agri",
    "Quantity: 6",
    "Tread Type: Agri",
    "Quantity: 4",
  ])),

  // ── Case 2: All-caps CUTTER and lowercase cutter ──────────────────────────
  // Expect: same columns as case 1 but different quantities
  msg("2026/04/03, 08:02", block([
    "03/04/2026",
    "",
    "Time:08:00-09:00",
    "",
    "Cutter 1",
    "Operator: Sipho",
    "Assistant: Andile",
    "Tyre Type:LC",
    "Quantity:17",
    "Tread Type:LC",
    "Quantity:11",
    "",
    "CUTTER 2",
    "Operator:Menzi",
    "Assistant: Philani",
    "Tyre Type: HC",
    "Quantity:9",
    "Tread Type: HC",
    "Quantity:6",
    "",
    "cutter 3",
    "Operator: Bhekinkosi",
    "Assistant: Sanele",
    "Tyre Type: Agri",
    "Quantity:5",
    "Tread Type:Agri",
    "Quantity:3",
  ])),

  // ── Case 3: Loose time formatting ("Time : 09:00 - 10:00") ───────────────
  // Expect: CM-1 LC=18+tread12  CM-2 HC=10  CM-3 Agri=4+tread2
  msg("2026/04/03, 09:01", block([
    "03/04/2026",
    "",
    "Time : 09:00 - 10:00",
    "",
    "Cutter 1",
    "Operator: Sipho",
    "Assistant: Andile",
    "Tyre Type: LC",
    "Quantity: 18",
    "Tread Type: LC",
    "Quantity: 12",
    "",
    "Cutter 2",
    "Operator: Menzi",
    "Assistant: Philani",
    "Tyre Type: HC",
    "Quantity: 10",
    "",
    "Cutter 3",
    "Operator: Bhekinkosi",
    "Assistant: Sanele",
    "Tyre Type: Agri",
    "Quantity: 4",
    "Tread Type: Agri",
    "Quantity: 2",
  ])),

  // ── Case 4: Missing Tread Type for one cutter ─────────────────────────────
  // Expect: CM-2 HC=10 only (no tread row)
  msg("2026/04/04, 08:01", block([
    "04/04/2026",
    "",
    "Time: 08:00-09:00",
    "",
    "Cutter 1",
    "Operator: Sipho",
    "Assistant: Andile",
    "Tyre Type: LC",
    "Quantity: 16",
    "Tread Type: LC",
    "Quantity: 10",
    "",
    "Cutter 2",
    "Operator: Menzi",
    "Assistant: Philani",
    "Tyre Type: HC",
    "Quantity: 8",
    "",
    "Cutter 3",
    "Operator: Bhekinkosi",
    "Assistant: Sanele",
    "Tyre Type: Agri",
    "Quantity: 4",
    "Tread Type: Agri",
    "Quantity: 3",
  ])),

  // ── Case 5: Duplicate cutter header in one message (operator submitted ────
  //            Cutter 2 twice — parser should handle gracefully)
  // Expect: two CM-2 HC rows (or one merged — verify in report)
  msg("2026/04/04, 09:00", block([
    "04/04/2026",
    "",
    "Time: 09:00-10:00",
    "",
    "Cutter 2",
    "Operator: Menzi",
    "Assistant: Philani",
    "Tyre Type: HC",
    "Quantity: 8",
    "Tread Type: HC",
    "Quantity: 5",
    "",
    "Cutter 2",
    "Operator: Menzi",
    "Assistant: Philani",
    "Tyre Type: HC",
    "Quantity: 9",
    "Tread Type: HC",
    "Quantity: 6",
    "",
    "Cutter 3",
    "Operator: Bhekinkosi",
    "Assistant: Sanele",
    "Tyre Type: Agri",
    "Quantity: 4",
    "Tread Type: Agri",
    "Quantity: 3",
  ])),

  // ── Case 6: Non-standard tyre type ("Passenger") ─────────────────────────
  // Expect: CM-1 type=unknown_type (Passenger maps to unknown)
  msg("2026/04/05, 08:05", block([
    "05/04/2026",
    "",
    "Time:08:00-09:00",
    "",
    "Cutter 1",
    "Operator: Sipho",
    "Assistant: Andile",
    "Tyre Type: Passenger",
    "Quantity: 20",
    "Tread Type: Passenger",
    "Quantity: 10",
    "",
    "Cutter 2",
    "Operator: Menzi",
    "Assistant: Philani",
    "Tyre Type: HC",
    "Quantity: 7",
    "Tread Type: HC",
    "Quantity: 4",
    "",
    "Cutter 3",
    "Operator: Bhekinkosi",
    "Assistant: Sanele",
    "Tyre Type: Agri",
    "Quantity: 5",
    "Tread Type: Agri",
    "Quantity: 3",
  ])),

  // ── Case 7: No colon spacing ("Tyre Type:LC" / "Quantity:17") ─────────────
  // Expect: parses correctly — same as case 2 values
  msg("2026/05/02, 08:00", block([
    "02/05/2026",
    "",
    "Time: 08:00-09:00",
    "",
    "Cutter 1",
    "Operator: Sipho",
    "Assistant: Andile",
    "Tyre Type: LC",
    "Quantity: 19",
    "Tread Type: LC",
    "Quantity: 13",
    "",
    "Cutter 2",
    "Operator: Menzi",
    "Assistant: Philani",
    "Tyre Type: HC",
    "Quantity: 9",
    "Tread Type: HC",
    "Quantity: 6",
    "",
    "Cutter 3",
    "Operator: Bhekinkosi",
    "Assistant: Sanele",
    "Tyre Type: Agri",
    "Quantity: 6",
    "Tread Type: Agri",
    "Quantity: 4",
  ])),

  // ── Case 8: New-format Cutting Summary (Cutter N / Total Tyres / LC-N) ────
  // Expect summaryRecords: CM-1 LC=36  CM-2 HC=17  CM-3 Agri=11
  msg("2026/05/02, 17:05", block([
    "02/05/2026",
    "",
    "Cutting Summary",
    "",
    "Cutter 1",
    "Total Tyres",
    "LC - 36",
    "HC - 0",
    "Agri - 0",
    "Total Treads",
    "LC - 24",
    "HC - 0",
    "Agri - 0",
    "",
    "Cutter 2",
    "Total Tyres",
    "LC - 0",
    "HC - 17",
    "Agri - 0",
    "Total Treads",
    "LC - 0",
    "HC - 11",
    "Agri - 0",
    "",
    "Cutter 3",
    "Total Tyres",
    "LC - 0",
    "HC - 0",
    "Agri - 11",
    "Total Treads",
    "LC - 0",
    "HC - 0",
    "Agri - 7",
  ])),

  // ── Case 9: New-format Cutting Summary – partial (some zero fields omitted)
  msg("2026/06/01, 17:10", block([
    "01/06/2026",
    "",
    "Cutting Summary",
    "",
    "Cutter 1",
    "Total Tyres",
    "LC - 21",
    "HC - 0",
    "Total Treads",
    "LC - 15",
    "HC - 0",
    "Agri - 0",
    "",
    "Cutter 2",
    "Total Tyres",
    "HC - 8",
    "Total Treads",
    "HC - 5",
    "",
    "Cutter 3",
    "Total Tyres",
    "Agri - 7",
    "Total Treads",
    "Agri - 4",
  ])),

  // ── Case 10: Ignored / noise messages ─────────────────────────────────────
  msg("2026/04/06, 09:00", "Good morning"),
  msg("2026/04/06, 09:01", "This message was deleted"),
  msg("2026/04/06, 09:02", "<Media omitted>"),

];

// ─── ALSO PARSE REAL CHAT FILES ───────────────────────────────────────────────

const realChatsDir = path.resolve(root, "..", "WhatsApp_chats", "Cutting");
const realChatFiles = {
  "Cutting_Feb_2025.txt": "old",
  "Cutting_Oct_2025.txt": "old",
  "BABE generated/fake_new_format_cutting_tests.txt": "new",
  "BABE generated/messy_cutting_parser_test.txt": "new",
};

// ─── BUILD + WRITE WORKBOOK (Node-compatible) ────────────────────────────────

async function buildCuttingWorkbook(records, summaryRecords, validationLog) {
  const { default: ExcelJS } = await import("exceljs");

  function normalizeCmNumber(cmNumber) {
    const m = String(cmNumber ?? "").match(/(\d+)/);
    if (!m) return null;
    const n = parseInt(m[1], 10);
    return Number.isNaN(n) ? null : n;
  }

  function withSummaryPlaceholders(summaryRows) {
    const byDate = new Map();
    for (const s of summaryRows) {
      if (!s?.date) continue;
      const key = `${s.date.year}-${String(s.date.month).padStart(2, "0")}-${String(s.date.day).padStart(2, "0")}`;
      if (!byDate.has(key)) byDate.set(key, { date: s.date, byCm: new Map() });
      const bucket = byDate.get(key).byCm;
      const cm = normalizeCmNumber(s.cmNumber);
      if (!cm) continue;
      if (!bucket.has(cm)) {
        bucket.set(cm, { date: s.date, cmNumber: `CM - ${cm}`, totalLC: null, totalHC: null, totalRadials: null, totalAgri: null, totalAgriTreads: null, totalUnknown: null });
      }
      const cur = bucket.get(cm);
      for (const k of ["totalLC", "totalHC", "totalRadials", "totalAgri", "totalAgriTreads", "totalUnknown"]) {
        if (s[k] !== null && s[k] !== undefined) cur[k] = s[k];
      }
      if (s._unresolved) cur._unresolved = true;
    }
    const out = [];
    const sorted = [...byDate.values()].sort((a, b) => {
      const ak = a.date.year * 10000 + a.date.month * 100 + a.date.day;
      const bk = b.date.year * 10000 + b.date.month * 100 + b.date.day;
      return ak - bk;
    });
    for (const d of sorted) {
      for (const cm of [1, 2, 3]) {
        const existing = d.byCm.get(cm);
        out.push(existing || { date: d.date, cmNumber: `CM - ${cm}`, totalLC: null, totalHC: null, totalRadials: null, totalAgri: null, totalAgriTreads: null, totalUnknown: null });
      }
    }
    return out;
  }

  const wb = new ExcelJS.Workbook();
  const styles = baseStyles();
  const byMonth = groupByMonth(records);
  const summaryByMonth = groupByMonth(summaryRecords);
  const allKeys = [...new Set([...Object.keys(byMonth), ...Object.keys(summaryByMonth)])].sort();

  for (const key of allKeys) {
    const sheetName = monthLabel(key);
    const monthRecords = byMonth[key] || [];
    const monthSummary = summaryByMonth[key] || [];
    const { rows, unresolvedIndices } = cuttingSheetRows(monthRecords);
    let totalsRowAdded = false;
    if (rows.length >= 2) { rows.push(["TOTALS", "", "", "", "", "", "", "", "", "", "", ""]); totalsRowAdded = true; }

    const unresolvedSummaryRows = [];
    if (monthSummary.length > 0) {
      const sorted = withSummaryPlaceholders(monthSummary);
      rows.push([]);
      rows.push(["Cutting Summary", "", "", "", "", "", ""]);
      rows.push(["Date", "CM Number", "LC", "HC", "Radials Total", "Agri", "Unknown"]);
      for (const s of sorted) {
        const nz = (v) => (v != null && v !== 0) ? v : "";
        const rt = s.totalRadials != null && s.totalRadials !== 0
          ? s.totalRadials
          : (s.totalLC || s.totalHC) ? (s.totalLC ?? 0) + (s.totalHC ?? 0) : "";
        rows.push([dateToStr(s.date), s.cmNumber, nz(s.totalLC), nz(s.totalHC), rt, nz(s.totalAgri), nz(s.totalUnknown)]);
        if (s._unresolved) unresolvedSummaryRows.push(rows.length);
      }
    }

    const ws = wb.addWorksheet(sheetName);
    rows.forEach((row) => ws.addRow(row));

    for (const merge of CUTTING_GROUP_MERGES) {
      ws.mergeCells(merge.s.r + 1, merge.s.c + 1, merge.e.r + 1, merge.e.c + 1);
    }

    // Style header rows 1-2
    for (const rowIdx of [1, 2]) {
      const row = ws.getRow(rowIdx);
      for (let c = 1; c <= 12; c++) {
        const cell = row.getCell(c);
        cell.font = { bold: true, color: { argb: styles.textDark } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC2C8D6" } };
        cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
        cell.border = styles.baseBorder;
      }
    }

    // Highlight unresolved rows pink
    const pink = "FFFADBD8";
    const dataStart = 3;
    for (const idx of unresolvedIndices) {
      const row = ws.getRow(dataStart + idx);
      for (let c = 1; c <= 12; c++) {
        row.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: pink } };
      }
    }
    for (const rowIdx of unresolvedSummaryRows) {
      const row = ws.getRow(rowIdx);
      for (let c = 1; c <= 7; c++) {
        row.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: pink } };
      }
    }
  }

  if (validationLog.length > 0) {
    const ws = wb.addWorksheet("Validation_Log");
    ws.addRow(["Body Date", "Time", "Message Type", "Cutter", "Issue", "Action Taken", "Raw Text"]);
    for (const e of validationLog) {
      ws.addRow([e.date, e.time, e.messageType, e.cutter, e.issue, e.action, e.rawText || ""]);
    }
    styleHeaderRow(ws.getRow(1), styles, false);
    styleBodyRows(ws, 2, ws.rowCount, styles.baseBorder, 7);
    ws.views = [{ state: "frozen", ySplit: 1 }];
    applyColumnWidths(ws, [12, 12, 14, 12, 44, 44, 60]);
  }

  return wb;
}

// ─── REPORT BUILDER ───────────────────────────────────────────────────────────

function buildReport(label, chatText, parsed, sourceFiles = []) {
  const { records, summaryRecords, validationLog } = parsed;

  // Count unknown-type records
  const unknownRecords = records.filter(r => r.radialsLC == null && r.radialsHC == null && r.radialsAgri == null && r.radialsAgriTreads == null && r.nylonsLC == null && !r._syntheticPlaceholder);
  const unresolvedSummary = summaryRecords.filter(s => s._unresolved);

  // Tally by tyre column
  const colTotals = {};
  for (const r of records) {
    for (const col of ["radialsLC", "radialsHC", "radialsAgri", "radialsAgriTreads", "nylonsLC"]) {
      if (r[col] != null) colTotals[col] = (colTotals[col] || 0) + r[col];
    }
  }

  // Count per date
  const byDate = {};
  for (const r of records) {
    const key = dateToStr(r.date);
    byDate[key] = (byDate[key] || 0) + 1;
  }

  const lines = [];
  lines.push(`Cutting ${label} QA Report`);
  lines.push(`Generated: ${new Date().toISOString()}`);
  lines.push("");

  if (sourceFiles.length > 0) {
    lines.push("Source Files");
    for (const f of sourceFiles) lines.push(`  - ${f}`);
    lines.push("");
  }

  lines.push("Parse Totals");
  lines.push(`  Hourly records:         ${records.length}`);
  lines.push(`  Summary records:        ${summaryRecords.length}`);
  lines.push(`  Validation log entries: ${validationLog.length}`);
  lines.push(`  Unknown-type records:   ${unknownRecords.length}`);
  lines.push(`  Unresolved summaries:   ${unresolvedSummary.length}`);
  lines.push("");

  lines.push("Tyre Column Totals (from hourly records)");
  for (const [col, n] of Object.entries(colTotals)) {
    lines.push(`  ${col.padEnd(22)} ${n}`);
  }
  if (Object.keys(colTotals).length === 0) lines.push("  (none)");
  lines.push("");

  lines.push("Records per Date");
  for (const [date, n] of Object.entries(byDate).sort()) {
    lines.push(`  ${date}  →  ${n} row(s)`);
  }
  if (Object.keys(byDate).length === 0) lines.push("  (none)");
  lines.push("");

  if (validationLog.length > 0) {
    const issueCounts = {};
    for (const v of validationLog) {
      const k = v.issue || "PARSER_WARNING";
      issueCounts[k] = (issueCounts[k] || 0) + 1;
    }
    lines.push("Validation Issue Types");
    for (const [k, n] of Object.entries(issueCounts).sort()) {
      lines.push(`  ${k}: ${n}`);
    }
    lines.push("");

    lines.push("All Validation Log Entries");
    for (const e of validationLog) {
      lines.push(`  [${e.date || ""}] ${e.cutter || ""} — ${e.issue || ""}: ${e.action || ""}`);
      if (e.rawText) lines.push(`    raw: ${String(e.rawText).replace(/\s+/g, " ").slice(0, 200)}`);
    }
    lines.push("");
  }

  if (unknownRecords.length > 0) {
    lines.push("Unknown-Type Hourly Records (check these!)");
    for (const r of unknownRecords) {
      lines.push(`  ${dateToStr(r.date)}  ${r.cmNumber}  ${r.startTime ? r.startTime.h + ":" + String(r.startTime.m).padStart(2,"0") : ""}-${r.finishTime ? r.finishTime.h + ":" + String(r.finishTime.m).padStart(2,"0") : ""}`);
    }
    lines.push("");
  }

  if (unresolvedSummary.length > 0) {
    lines.push("Unresolved Summary Records (check these!)");
    for (const s of unresolvedSummary) {
      lines.push(`  ${dateToStr(s.date)}  ${s.cmNumber}  unknown=${s.totalUnknown}`);
    }
    lines.push("");
  }

  return lines.join("\n");
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────

async function main() {

  // ── Old format: fixtures ──────────────────────────────────────────────────
  const oldChatText = oldMessages.join("\n");
  const oldChatPath = path.resolve(sampleDir, "cutting_old_qa_chat.txt");
  fs.writeFileSync(oldChatPath, oldChatText, "utf8");
  const oldParsed = parseCuttingMessages(oldChatText);

  // ── Old format: real chat files ───────────────────────────────────────────
  const realOldSources = [];
  for (const [file, format] of Object.entries(realChatFiles)) {
    if (format !== "old") continue;
    const p = path.join(realChatsDir, file);
    if (!fs.existsSync(p)) { console.warn(`Skipping missing file: ${p}`); continue; }
    const text = fs.readFileSync(p, "utf8");
    const parsed = parseCuttingMessages(text);
    oldParsed.records.push(...parsed.records);
    oldParsed.summaryRecords.push(...parsed.summaryRecords);
    oldParsed.validationLog.push(...parsed.validationLog);
    realOldSources.push(p);
  }

  const oldJsonPath   = path.resolve(outDir, "cutting_old_parsed.json");
  const oldReportPath = path.resolve(outDir, "cutting_old_qa_report.txt");
  const oldXlsxPath   = path.resolve(outDir, "BPR_Cutting_QA_Old.xlsx");
  fs.writeFileSync(oldJsonPath, JSON.stringify(oldParsed, null, 2), "utf8");
  fs.writeFileSync(oldReportPath, buildReport("Old-Format", oldChatText, oldParsed, realOldSources), "utf8");
  const oldWb = await buildCuttingWorkbook(oldParsed.records, oldParsed.summaryRecords, oldParsed.validationLog);
  await oldWb.xlsx.writeFile(oldXlsxPath);

  // ── New format: fixtures ──────────────────────────────────────────────────
  const newChatText = newMessages.join("\n");
  const newChatPath = path.resolve(sampleDir, "cutting_new_qa_chat.txt");
  fs.writeFileSync(newChatPath, newChatText, "utf8");
  const newParsed = parseCuttingMessagesNew(newChatText);

  // ── New format: real chat files ───────────────────────────────────────────
  const realNewSources = [];
  for (const [file, format] of Object.entries(realChatFiles)) {
    if (format !== "new") continue;
    const p = path.join(realChatsDir, file);
    if (!fs.existsSync(p)) { console.warn(`Skipping missing file: ${p}`); continue; }
    const text = fs.readFileSync(p, "utf8");
    const parsed = parseCuttingMessagesNew(text);
    newParsed.records.push(...parsed.records);
    newParsed.summaryRecords.push(...parsed.summaryRecords);
    newParsed.validationLog.push(...parsed.validationLog);
    realNewSources.push(p);
  }

  const newJsonPath   = path.resolve(outDir, "cutting_new_parsed.json");
  const newReportPath = path.resolve(outDir, "cutting_new_qa_report.txt");
  const newXlsxPath   = path.resolve(outDir, "BPR_Cutting_QA_New.xlsx");
  fs.writeFileSync(newJsonPath, JSON.stringify(newParsed, null, 2), "utf8");
  fs.writeFileSync(newReportPath, buildReport("New-Format", newChatText, newParsed, realNewSources), "utf8");
  const newWb = await buildCuttingWorkbook(newParsed.records, newParsed.summaryRecords, newParsed.validationLog);
  await newWb.xlsx.writeFile(newXlsxPath);

  // ── Summary ───────────────────────────────────────────────────────────────
  console.log("\nCutting QA pack generated:\n");
  console.log("Old-Format");
  console.log(`  Fixture chat:  ${oldChatPath}`);
  console.log(`  Parsed JSON:   ${oldJsonPath}`);
  console.log(`  Report:        ${oldReportPath}`);
  console.log(`  Excel:         ${oldXlsxPath}`);
  console.log(`  Records: ${oldParsed.records.length}  Summary: ${oldParsed.summaryRecords.length}  Warnings: ${oldParsed.validationLog.length}`);
  console.log("");
  console.log("New-Format");
  console.log(`  Fixture chat:  ${newChatPath}`);
  console.log(`  Parsed JSON:   ${newJsonPath}`);
  console.log(`  Report:        ${newReportPath}`);
  console.log(`  Excel:         ${newXlsxPath}`);
  console.log(`  Records: ${newParsed.records.length}  Summary: ${newParsed.summaryRecords.length}  Warnings: ${newParsed.validationLog.length}`);
}

main().catch((err) => { console.error(err); process.exit(1); });
