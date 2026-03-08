import assert from "node:assert/strict";
import { parseBalingMessages } from "../src/parsing/balingParser.js";

function run(name, fn) {
  try {
    fn();
    console.log(`PASS ${name}`);
  } catch (err) {
    console.error(`FAIL ${name}`);
    console.error(err.message);
    process.exitCode = 1;
  }
}

function parseOne(body) {
  const chat = `2024/10/18, 11:30 - Tester: ${body}`;
  const parsed = parseBalingMessages(chat);
  if (parsed.standardRecords.length) return parsed.standardRecords[0];
  if (parsed.failedRecords.length) return parsed.failedRecords[0];
  throw new Error("No parsed record");
}

run("operator keeps only text before Assistant", () => {
  const r = parseOne("Machine Two Date - 18/10/2024 B329 - Production Operator - Bhekinkosi Assistant - Bhekisisa Start time - 11:20 Finish time - 11:36 Item - Light Commercial Qty - 36 Passenger Qty - 10 Total Qty - 46 Weight - 980kg");
  assert.equal(r.operator, "Bhekinkosi");
});

run("operator N/A becomes blank", () => {
  const r = parseOne("Machine Two Date - 18/10/2024 B330 - Production Operator - N/A Assistant - Bhekisisa Start time - 11:20 Finish time - 11:36 Item - Passenger Qty - 10 Total Qty - 10 Weight - 900kg");
  assert.equal(r.operator, "");
});

run("assistant remains clean without trailing process fields", () => {
  const r = parseOne("Machine One Date - 18/10/2024 B331 - Production Operator - Donna Assistant - Menzi Start time - 09:00 Finish time - 09:20 Item - Passenger Qty - 10 Total Qty - 10 Weight - 900kg");
  assert.equal(r.assistant, "Menzi");
});
