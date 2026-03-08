import assert from "node:assert/strict";
import { parseBalingMessages, isFailedBaleMessage, extractFailedBaleReason } from "../src/parsing/balingParser.js";

function wrapWhatsApp(body, ts = "2024/09/23, 09:30", sender = "Tester") {
  return `${ts} - ${sender}: ${body}`;
}

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

run("1) Explicit failed bale is classified/routed and not emitted to standard", () => {
  const body = "Machine Two Date - 23/09/2024 Failed Bale - Wire's snapped Operator - Donna Assistant - Menzi Start time - 09:00 Finish time - 09:20";
  const parsed = parseBalingMessages(wrapWhatsApp(body));

  assert.equal(parsed.failedRecords.length, 1);
  assert.equal(parsed.standardRecords.length, 0);
  assert.match(parsed.failedRecords[0].failureReason.toLowerCase(), /wire\s+snapped|wires\s+snapped/);

  const rerouteLog = parsed.validationLog.find((v) => v.issueType === "FAILED_BALE_REROUTED");
  assert.ok(rerouteLog);
  assert.equal(rerouteLog.severity, "INFO");
});

run("2) Wording variation (due to wire snapped) routes to failed with partial fields", () => {
  const body = "Machine Two Failed Bale due to wire snapped Date - 30/09/2024 Operator - Lange Assistant - Philani Start time - 11:23 Finish time -11:40 Item 4x4 - 72 Total Qty - 72 Weight - 1020kg";
  const parsed = parseBalingMessages(wrapWhatsApp(body, "2024/09/30, 11:45"));

  assert.equal(parsed.failedRecords.length, 1);
  assert.equal(parsed.standardRecords.length, 0);
  const row = parsed.failedRecords[0];
  assert.equal(row.machine, "BM - 2");
  assert.equal(row.fourx4Qty, 72);
  assert.equal(row.totalQty, 72);
  assert.equal(row.weightKg, 1020);
  assert.match((row.failureReason || "").toLowerCase(), /wire\s+snapped/);
});

run("3) Admin instruction about failed bales does not pollute standard production", () => {
  const body = "We must report failed bails separately - Change B43 to B42 and B44 to B43 A failed bale is not a bale We will only report to account for lost time and consumables etc";
  const parsed = parseBalingMessages(wrapWhatsApp(body, "2024/09/23, 17:00"));

  assert.equal(parsed.standardRecords.length, 0);
  assert.ok(parsed.failedRecords.length >= 1);

  const uncertain = parsed.validationLog.find((v) => v.issueType === "FAILED_BALE_UNCERTAIN");
  assert.ok(uncertain, "Expected uncertain failed-bale warning for admin text");
  assert.equal(uncertain.severity, "WARNING");
});

run("4) Normal production row is not classified as failed_bale", () => {
  const body = "Machine One Date - 30/09/2024 B143 - Production Operator - Donna Assistant - Sanele Start time - 14:08 Finish time -14:22 Item - 4x4 Qty - 71 Total Qty - 71 Weight - 1024kg";
  const parsed = parseBalingMessages(wrapWhatsApp(body, "2024/09/30, 14:25"));

  assert.equal(parsed.failedRecords.length, 0);
  assert.equal(parsed.standardRecords.length, 1);
});

run("5) Avoid false positive on unrelated 'failed' admin message", () => {
  const sample = "System failed to upload attendance report. Will retry in 10 mins.";
  const cls = isFailedBaleMessage(sample);
  assert.equal(cls.isFailed, false);
});

run("6) Failed reason extractor handles rope/strap variants", () => {
  assert.equal(extractFailedBaleReason("Fail bale - rope snapped"), "rope snapped");
  assert.equal(extractFailedBaleReason("Bale popped during baling"), "bale popped");
});
