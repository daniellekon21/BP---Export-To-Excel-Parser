import fs from "node:fs";
import path from "node:path";
import { parseBalingMessages } from "../src/parsing/balingParser.js";
import { splitWhatsAppMessages } from "../src/helpers.js";

function arg(name, fallback = "") {
  const found = process.argv.find((a) => a.startsWith(`--${name}=`));
  return found ? found.split("=").slice(1).join("=") : fallback;
}

function shortLine(text, max = 140) {
  const one = String(text || "").replace(/\s+/g, " ").trim();
  return one.length > max ? `${one.slice(0, max - 1)}…` : one;
}

function safePct(n, d) {
  if (!d) return 0;
  return (n / d) * 100;
}

function clamp(v, lo, hi) {
  return Math.max(lo, Math.min(hi, v));
}

const input = arg("input", "./sample-data/WhatsApp_Chat_with_Blue_pyramid_recycling-_Baled_Finished_Goods.txt");
const out = arg("out", "");

const inputPath = path.resolve(process.cwd(), input);
if (!fs.existsSync(inputPath)) {
  console.error(`Input not found: ${inputPath}`);
  process.exit(1);
}

const text = fs.readFileSync(inputPath, "utf8");
const messages = splitWhatsAppMessages(text);
const parsed = parseBalingMessages(text);

const keywordCandidateRegex = /\b(date\s*[:\-]\s*\d{1,2}\/\d{1,2}\/\d{4}|\bB\s?\d{1,4}\b|failed?\s*bale|wire\s*snapped|scrap|side\s*wall|\bcr\s*[- ]?\d+\b|\bca\s*[- ]?\d+\b|daily\s*summary|\bsummary\b|weight\s*[:\-]\s*\d+)\b/i;
const candidateMessages = messages.filter((m) => keywordCandidateRegex.test(m.body));

const allRows = [
  ...parsed.standardRecords,
  ...parsed.failedRecords,
  ...parsed.scrapRecords,
  ...parsed.crcaRecords,
  ...parsed.summaryRecords,
];

const warnedRaw = new Set(parsed.validationLog.map((v) => String(v.rawMessage || "")));
const cleanRows = allRows.filter((r) => !warnedRaw.has(String(r.rawMessage || "")));

const byIssue = new Map();
for (const v of parsed.validationLog) {
  const k = v.issueType || "PARSER_WARNING";
  if (!byIssue.has(k)) byIssue.set(k, []);
  byIssue.get(k).push(v);
}

const sortedIssues = [...byIssue.entries()].sort((a, b) => b[1].length - a[1].length);

const coverage = safePct(allRows.length, candidateMessages.length || 1);
const cleanliness = safePct(cleanRows.length, allRows.length || 1);
const errorCount = parsed.validationLog.filter((v) => String(v.severity).toUpperCase() === "ERROR").length;
const warningCount = parsed.validationLog.filter((v) => String(v.severity).toUpperCase() === "WARNING").length;
const infoCount = parsed.validationLog.filter((v) => String(v.severity).toUpperCase() === "INFO").length;

let score = 100;
score -= clamp((100 - Math.min(coverage, 100)) * 0.35, 0, 35);
score -= clamp((100 - cleanliness) * 0.45, 0, 45);
score -= clamp(errorCount * 3, 0, 15);
score -= clamp((warningCount / Math.max(1, allRows.length)) * 100 * 0.25, 0, 20);
score = clamp(Math.round(score), 0, 100);

const lines = [];
lines.push("Baling Parser Performance Report");
lines.push(`Input: ${inputPath}`);
lines.push("");
lines.push("Overview");
lines.push(`- Total WhatsApp messages: ${messages.length}`);
lines.push(`- Candidate production/summary messages: ${candidateMessages.length}`);
lines.push(`- Parsed rows emitted (all sheets): ${allRows.length}`);
lines.push(`- Parsed clean rows (no warnings): ${cleanRows.length}`);
lines.push(`- Validation log entries: ${parsed.validationLog.length} (INFO ${infoCount}, WARNING ${warningCount}, ERROR ${errorCount})`);
lines.push("");
lines.push("What Works Well");
lines.push(`- Standard bales parsed: ${parsed.standardRecords.length}`);
lines.push(`- Failed bales parsed: ${parsed.failedRecords.length}`);
lines.push(`- Scrap sidewall records parsed: ${parsed.scrapRecords.length}`);
lines.push(`- CR/CA or test records parsed: ${parsed.crcaRecords.length}`);
lines.push(`- Daily summaries parsed: ${parsed.summaryRecords.length}`);
lines.push(`- Candidate coverage: ${coverage.toFixed(1)}%`);
lines.push(`- Clean row ratio: ${cleanliness.toFixed(1)}%`);
lines.push("");
lines.push("Score");
lines.push(`- Overall score: ${score}/100`);
lines.push(`- Formula: coverage 35% + cleanliness 45% + error/warning penalties`);
lines.push("");
lines.push("Problematic Texts / Rows");

if (sortedIssues.length === 0) {
  lines.push("- No validation issues found.");
} else {
  for (const [issueType, entries] of sortedIssues.slice(0, 12)) {
    lines.push(`- ${issueType}: ${entries.length}`);
    for (const ex of entries.slice(0, 3)) {
      lines.push(`  • ${shortLine(ex.problemDescription || ex.issue, 120)}`);
      lines.push(`    Msg: ${shortLine(ex.rawMessage, 170)}`);
    }
  }
}

lines.push("");
lines.push("High-Risk Gaps");
if (errorCount === 0) {
  lines.push("- No ERROR-severity items logged.");
} else {
  for (const e of parsed.validationLog.filter((v) => String(v.severity).toUpperCase() === "ERROR").slice(0, 8)) {
    lines.push(`- ${e.issueType}: ${shortLine(e.problemDescription || e.issue, 140)}`);
    lines.push(`  Msg: ${shortLine(e.rawMessage, 170)}`);
  }
}

const report = lines.join("\n");
console.log(report);

if (out) {
  const outPath = path.resolve(process.cwd(), out);
  fs.writeFileSync(outPath, report, "utf8");
  console.log(`\nSaved report: ${outPath}`);
}
