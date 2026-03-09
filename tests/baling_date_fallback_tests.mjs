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
  assert.equal(r.bodyDateText, "25/10/2034");
  assert.equal(r.usedTimestampFallbackForFutureBodyDate, true);
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

run("standalone body date line is used over WhatsApp timestamp date", () => {
  const chat = "2025/07/03, 09:59 - Tester: Machine 1\n02/07/2025\nCA015-07/25\nOperator-Bhekisisa\nAssistant-Bhekinosi\nStart time-09:38\nFinish time-09:58\nProcess-20 mins\nT-43\nWeight-925kg";
  const parsed = parseBalingMessages(chat);

  assert.ok(parsed.allRecords.length >= 1);
  const r = parsed.allRecords[0];
  assert.equal(r.chatDateParsed.year, 2025);
  assert.equal(r.chatDateParsed.month, 7);
  assert.equal(r.chatDateParsed.day, 2);
});

run("machine BM variants and Ass label parse case-insensitively", () => {
  const chat = "2026/01/09, 19:50 - Tester: 09/01/2026\nBM#:2\nCN003-01/2026\nOperator: Donna\nAss=Andile\nStart: 19:31\nFinish: 19:49\nSW: 19\nWeight: 620 kg";
  const parsed = parseBalingMessages(chat);

  assert.equal(parsed.standardRecords.length, 1);
  const r = parsed.standardRecords[0];
  assert.equal(r.machine, "BM - 2");
  assert.equal(r.assistant, "Andile");
});

run("TB bale parses quantity from TB label with :, -, = and any case", () => {
  const chat1 = "2026/01/09, 19:50 - Tester: 09/01/2026\nBM #: 1\nTB001-01/2026\nOperator: Menzi\nAss: Langa\nTB: 12\nWeight: 734 kg";
  const chat2 = "2026/01/09, 19:50 - Tester: 09/01/2026\nbm#=1\ntb001-01/2026\nOperator: Menzi\nAss: Langa\ntB-7\nWeight: 734 kg";
  const chat3 = "2026/01/09, 19:50 - Tester: 09/01/2026\nBM#-1\nTB001-01/2026\nOperator: Menzi\nAss: Langa\nTB=9\nWeight: 734 kg";

  const p1 = parseBalingMessages(chat1).standardRecords[0];
  const p2 = parseBalingMessages(chat2).standardRecords[0];
  const p3 = parseBalingMessages(chat3).standardRecords[0];

  assert.equal(p1.tubeQty, 12);
  assert.equal(p2.tubeQty, 7);
  assert.equal(p3.tubeQty, 9);
});

run("TB bale with LC/T/SW text is treated as malformed for TB qty and non-TB fields stay null", () => {
  const chat = "2026/01/09, 19:50 - Tester: 09/01/2026\nBM #: 1\nTB001-01/2026\nOperator: Menzi\nAss: Langa\nStart: 13:31\nFinish: 13:57\nLC\nT: 7\nSW: 0\nWeight: 734 kg";
  const parsed = parseBalingMessages(chat);
  assert.equal(parsed.standardRecords.length, 1);
  const r = parsed.standardRecords[0];
  assert.equal(r.tubeQty, null);
  assert.equal(r.lcQty, null);
  assert.equal(r.treadQty, null);
  assert.equal(r.sideWallQty, null);
  const warn = parsed.validationLog.find((v) => v.issueType === "MALFORMED_TB");
  assert.ok(warn);
});

run("Start/Finish with '=' are parsed and duration is computed from end-start when total time label is absent", () => {
  const chat = "2026/01/09, 19:50 - Tester: 09/01/2026\nBM #: 1\nPCR010-01/2026\nOperator: Donna\nAss: Andile\nStart = 19:31\nFinish = 19:49\nLC: 80\nWeight: 620 kg";
  const parsed = parseBalingMessages(chat);
  assert.equal(parsed.standardRecords.length, 1);
  const r = parsed.standardRecords[0];
  assert.equal(r.startTime?.h, 19);
  assert.equal(r.startTime?.m, 31);
  assert.equal(r.finishTime?.h, 19);
  assert.equal(r.finishTime?.m, 49);
  assert.equal(r.durationMinutes, 18);
});
