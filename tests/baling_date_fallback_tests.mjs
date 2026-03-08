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

run("future body date falls back to WhatsApp timestamp date", () => {
  const chat = "2024/10/25, 11:10 - Tester: Machine One Date - 25/10/2034 B389 - Production Operator - Donna Assistant - Bongumusa Start time - 13:08 Finish time -13:29 Item - Passengers - 125 Total Qty -125 Weight - 995kg";
  const parsed = parseBalingMessages(chat);

  assert.equal(parsed.standardRecords.length, 1);
  const r = parsed.standardRecords[0];
  assert.equal(r.chatDateParsed.year, 2024);
  assert.equal(r.chatDateParsed.month, 10);
  assert.equal(r.chatDateParsed.day, 25);
});

run("time parsing supports optional seconds and defaults to :00", () => {
  const chat = "2024/10/25, 11:10 - Tester: Machine One Date - 25/10/2024 B389 - Production Operator - Donna Assistant - Bongumusa Start time - 13:08 Finish time -13:29:15 Item - Passengers - 125 Total Qty -125 Weight - 995kg";
  const parsed = parseBalingMessages(chat);

  assert.equal(parsed.standardRecords.length, 1);
  const r = parsed.standardRecords[0];
  assert.equal(r.startTime.h, 13);
  assert.equal(r.startTime.m, 8);
  assert.equal(r.startTime.s, 0);
  assert.equal(r.finishTime.h, 13);
  assert.equal(r.finishTime.m, 29);
  assert.equal(r.finishTime.s, 15);
});
