import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseBalingMessagesNew } from "../src/parsing/balingParserNew.js";
import { createBalingWorkbook } from "../src/excel/createBalingWorkbook.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const root = path.resolve(__dirname, "..");

const outDir = path.resolve(root, "qa_outputs");
const sampleDir = path.resolve(root, "sample-data");
fs.mkdirSync(outDir, { recursive: true });
fs.mkdirSync(sampleDir, { recursive: true });

const chatOut = path.resolve(sampleDir, "generated_baling_newformat_qa_chat.txt");
const xlsxOut = path.resolve(outDir, "BPR_Baling_Data_QA_NewFormat.xlsx");
const jsonOut = path.resolve(outDir, "baling_newformat_qa_parsed.json");
const reportOut = path.resolve(outDir, "baling_newformat_qa_report.txt");

function msg(ts, body, sender = "QA Bot") {
  return `${ts} - ${sender}: ${body}`;
}

function block(lines) {
  return lines.join("\n");
}

const messages = [
  // --- Valid new-format production rows ---
  msg("2026/01/09, 08:20", block([
    "09/01/2026",
    "BM #: 1",
    "",
    "PCR003-01/2026",
    "Operator: Bekinkosi",
    "Ass: Bhekisisa",
    "Start: 18:45",
    "Finish: 18:59",
    "Total Time: 14 minutes",
    "PCR",
    "Lc: 80",
    "Weight: 650 kg",
  ])),
  msg("2026/01/09, 09:15", block([
    "09/01/2026",
    "bm#=2",
    "",
    "PCR004-01/2026",
    "Operator: Donna",
    "ASS: Andile",
    "Start: 09:05",
    "Finish: 09:28",
    "Passenger: 93",
    "4x4: 7",
    "Weight: 902 KG",
  ])),
  msg("2026/01/10, 10:02", block([
    "10/01/2026",
    "BM#:1",
    "",
    "CN003-01/2026",
    "Operator: Donna",
    "Ass: Andile",
    "Start: 19:31",
    "Finish: 19:49",
    "LC: 38",
    "SW: 19",
    "Weight: 620 kg",
  ])),
  msg("2026/01/10, 11:40", block([
    "10/01/2026",
    "bm - 1",
    "",
    "CRC001-01/2026",
    "Operator: Menzi",
    "Ass: Langa",
    "Start: 10:45",
    "Finish: 11:04",
    "LC Full: 12",
    "LC Cut",
    "T: 24",
    "SW: 28",
    "Weight: 955 kg",
  ])),
  msg("2026/01/10, 12:20", block([
    "10/01/2026",
    "BM: 2",
    "",
    "CRS001-01/2026",
    "Operator: Bekinkosi",
    "Ass: Bhekisisa",
    "Start: 10:45",
    "Finish: 11:04",
    "HC Full: 15",
    "HC Cut",
    "T: 28",
    "SW: 41",
    "Weight: 955 kg",
  ])),
  msg("2026/01/10, 13:02", block([
    "10/01/2026",
    "BM #: 2",
    "",
    "CRS002-01/2026",
    "Operator: Bekinkosi",
    "Ass: Bhekisisa",
    "Start: 13:10",
    "Finish: 13:33",
    "LC Full: 10",
    "LC Cut",
    "LC T: 20",
    "LC SW: 30",
    "HC Full: 5",
    "HC Cut",
    "HC T: 12",
    "HC SW: 8",
    "Weight: 980 kg",
  ])),
  msg("2026/01/11, 08:55", block([
    "11/01/2026",
    "BM #: 1",
    "",
    "CA012-01/2026",
    "Operator: Menzi",
    "Ass: Langa",
    "Start: 08:20",
    "Finish: 08:46",
    "T: 44",
    "SW: 26",
    "Weight: 910 kg",
  ])),
  msg("2026/01/11, 10:14", block([
    "11/01/2026",
    "BM #: 1",
    "",
    "TB001-01/2026",
    "Operator: Menzi",
    "Ass: Langa",
    "Start: 13:31",
    "Finish: 13:57",
    "TB: 14",
    "Weight: 734 kg",
  ])),
  msg("2026/01/11, 10:18", block([
    "11/01/2026",
    "BM #:1",
    "",
    "tb002-01/2026",
    "Operator: Menzi",
    "Ass: Langa",
    "Start: 14:05",
    "Finish: 14:26",
    "tb-11",
    "Weight: 711 kg",
  ])),
  msg("2026/01/11, 10:25", block([
    "11/01/2026",
    "BM#=1",
    "",
    "TB003-01/2026",
    "Operator: Menzi",
    "Ass: Langa",
    "Start: 14:35",
    "Finish: 14:52",
    "TB=9",
    "Weight: 698 kg",
  ])),
  msg("2026/01/12, 09:00", block([
    "12/01/2026",
    "BM #: 2",
    "",
    "ConV001-01/2026",
    "Operator: Donna",
    "Ass: Andile",
    "Start: 11:10",
    "Finish: 11:38",
    "Process: 28 minutes",
    "Weight: 810 kg",
  ])),
  msg("2026/01/12, 10:05", block([
    "12/01/2026",
    "BM #: 1",
    "",
    "PShrB001-01/2026",
    "Operator: Bekinkosi",
    "Ass: Bhekisisa",
    "Start: 12:05",
    "Finish: 12:22",
    "Weight: 1020 kg",
  ])),

  // --- Old-format/summary/fallback variants ---
  msg("2026/01/12, 15:10", "Daily summary Date - 12/01/2026 Bales - 12 Weight - 9500 kg Tons - 9.5 Passenger - 500 4x4 - 60 LC - 210"),
  msg("2026/01/12, 15:20", "Machine 1 Date - 12/01/2026 B123 - Production Operator - Donna Assistant - Menzi Start time - 14:08 Finish time -14:22 Item - Passenger Qty - 71 Total Qty - 71 Weight - 1024kg"),

  // --- Intentionally problematic / should warn ---
  msg("2026/01/13, 09:03", block([
    "09/01/2034",
    "BM #: 1",
    "",
    "PCR999-01/2026",
    "Operator: Donna",
    "Ass: Andile",
    "Start: 09:00",
    "Finish: 09:20",
    "Passenger: 100",
    "Weight: 905 kg",
  ])),
  msg("2026/01/13, 09:20", block([
    "13/01/2026",
    "BM #: 2",
    "",
    "CRS003-01/2026",
    "Operator: Bekinkosi",
    "Ass: Bhekisisa",
    "Start: 11:30",
    "Finish: 11:12",
    "HC Full: 9",
    "HC T: 18",
    "HC SW: 20",
    "Weight: 890 kg",
  ])),
  msg("2026/01/13, 09:24", block([
    "13/01/2026",
    "BM #: 2",
    "",
    "PCR007-01/2026",
    "Operator: Donna",
    "Ass: Andile",
    "Start: 10:12",
    "Finish: 10:32",
    "Passenger: 50",
    "MysteryCompound: 19",
    "Weight: 730 kg",
  ])),
  msg("2026/01/13, 09:30", block([
    "13/01/2026",
    "BM #: 2",
    "",
    "PCR007-01/2026",
    "Operator: Donna",
    "Ass: Andile",
    "Start: 10:35",
    "Finish: 10:54",
    "Passenger: 52",
    "Weight: 744 kg",
  ])),
  msg("2026/01/13, 09:40", "Machine Two Date - 13/01/2026 Failed Bale - wire snapped Operator - Donna Assistant - Menzi Start time - 09:00 Finish time - 09:20"),

  // --- February spread ---
  msg("2026/02/03, 08:10", block([
    "03/02/2026",
    "BM #: 1",
    "",
    "CN018-02/2026",
    "Operator: Donna",
    "Ass: Andile",
    "Start: 08:01",
    "Finish: 08:19",
    "LC: 42",
    "SW: 16",
    "Weight: 640 kg",
  ])),
  msg("2026/02/03, 09:10", block([
    "03/02/2026",
    "BM #: 2",
    "",
    "CRC010-02/2026",
    "Operator: Menzi",
    "Ass: Langa",
    "Start: 09:05",
    "Finish: 09:33",
    "LC Full: 6",
    "LC T: 14",
    "LC SW: 18",
    "HC Full: 4",
    "HC T: 11",
    "HC SW: 12",
    "Weight: 900 kg",
  ])),
  msg("2026/02/03, 09:45", "Daily summary Date - 03/02/2026 Bales - 8 Weight - 6400 kg Tons - 6.4 Passenger - 330 LC - 120"),

  // --- Soft-noise / ignored chat lines ---
  msg("2026/02/03, 10:00", "Noted, thanks team"),
  msg("2026/02/03, 10:03", "This message was deleted"),
  msg("2026/02/03, 10:06", "<Media omitted>"),
];

const chatText = messages.join("\n");
fs.writeFileSync(chatOut, chatText, "utf8");

const parsed = parseBalingMessagesNew(chatText);
fs.writeFileSync(jsonOut, JSON.stringify(parsed, null, 2), "utf8");

const wb = await createBalingWorkbook(parsed);
await wb.xlsx.writeFile(xlsxOut);

const issueCounts = new Map();
for (const v of parsed.validationLog) {
  const key = v.issueType || "PARSER_WARNING";
  issueCounts.set(key, (issueCounts.get(key) || 0) + 1);
}

const lines = [];
lines.push("Baling New-Format QA Report");
lines.push(`Generated on: ${new Date().toISOString()}`);
lines.push("");
lines.push("Artifacts");
lines.push(`- Chat input: ${chatOut}`);
lines.push(`- Parsed JSON: ${jsonOut}`);
lines.push(`- Excel output: ${xlsxOut}`);
lines.push("");
lines.push("Coverage Overview");
lines.push(`- Total WhatsApp messages generated: ${messages.length}`);
lines.push(`- Standard records: ${parsed.standardRecords.length}`);
lines.push(`- Failed records: ${parsed.failedRecords.length}`);
lines.push(`- Scrap records: ${parsed.scrapRecords.length}`);
lines.push(`- CR/CA test records: ${parsed.crcaRecords.length}`);
lines.push(`- Summary records: ${parsed.summaryRecords.length}`);
lines.push(`- Ignored messages: ${parsed.ignoredMessages}`);
lines.push(`- Validation entries: ${parsed.validationLog.length}`);
lines.push("");
lines.push("Issue Types");
for (const [k, v] of [...issueCounts.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
  lines.push(`- ${k}: ${v}`);
}

const topProblems = parsed.validationLog.slice(0, 20);
if (topProblems.length > 0) {
  lines.push("");
  lines.push("Sample Problem Rows (first 20 validation entries)");
  for (const p of topProblems) {
    lines.push(`- [${p.severity}] ${p.issueType}: ${p.problemDescription}`);
    lines.push(`  Date=${p.chatDateParsed || ""} Bale=${p.baleNumberCode || ""} Machine=${p.machine || ""}`);
    lines.push(`  Raw=${String(p.rawMessage || "").replace(/\s+/g, " ").slice(0, 220)}`);
  }
}

fs.writeFileSync(reportOut, lines.join("\n"), "utf8");

console.log("Generated QA artifacts:");
console.log(`- ${chatOut}`);
console.log(`- ${jsonOut}`);
console.log(`- ${xlsxOut}`);
console.log(`- ${reportOut}`);
