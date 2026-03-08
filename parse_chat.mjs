import fs from "node:fs";
import path from "node:path";
import { parseCuttingMessages, parseCuttingMessagesNew } from "./src/parsing/cuttingParser.js";
import { parseBalingMessages } from "./src/parsing/balingParser.js";

function arg(name, fallback = "") {
  const found = process.argv.find((a) => a.startsWith(`--${name}=`));
  return found ? found.split("=").slice(1).join("=") : fallback;
}

const type = (arg("type", "cutting") || "cutting").toLowerCase();
const format = (arg("format", "old") || "old").toLowerCase();
const input = arg("input", "");
const output = arg("output", "");

if (!input) {
  console.error("Missing --input=<path-to-whatsapp-chat.txt>");
  process.exit(1);
}

const inputPath = path.resolve(process.cwd(), input);
if (!fs.existsSync(inputPath)) {
  console.error(`Input file not found: ${inputPath}`);
  process.exit(1);
}

const text = fs.readFileSync(inputPath, "utf8");
let parsed;

if (type === "baling") {
  parsed = parseBalingMessages(text);
  console.log(`Baling parsed: ${parsed.allRecords.length} records (${parsed.standardRecords.length} standard, ${parsed.failedRecords.length} failed, ${parsed.scrapRecords.length} scrap, ${parsed.crcaRecords.length} CR/CA, ${parsed.summaryRecords.length} summaries)`);
} else {
  parsed = format === "new" ? parseCuttingMessagesNew(text) : parseCuttingMessages(text);
  console.log(`Cutting parsed: ${parsed.records.length} records, ${parsed.summaryRecords.length} summaries, ${parsed.validationLog.length} validation entries`);
}

if (output) {
  const outPath = path.resolve(process.cwd(), output);
  fs.writeFileSync(outPath, JSON.stringify(parsed, null, 2));
  console.log(`Wrote parsed output to ${outPath}`);
}
