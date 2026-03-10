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

run("daily summary category headers parse in both orders, separators, spacing, and case-insensitive", () => {
  const chat = [
    "2026/02/02, 18:40 - Tester: *Daily summary:* 02/02/2026",
    "",
    "  20    -    ca (Agricultural)",
    "Agricultural - 719",
    "Sidewalls - 727",
    "17644 kg",
    "",
    "CRC = 4",
    "LC T: 130",
    "Sidewalls: 86",
    "3949 kg",
    "",
    "pShRb : 2",
    "1020 kg",
    "",
    "ConV - 1",
    "600 kg",
  ].join("\n");

  const parsed = parseBalingMessages(chat);
  assert.equal(parsed.summaryRecords.length, 1);
  const s = parsed.summaryRecords[0];
  assert.equal(s.chatDateParsed.year, 2026);
  assert.equal(s.chatDateParsed.month, 2);
  assert.equal(s.chatDateParsed.day, 2);
  assert.equal(s.summaryCategories.CA.baleCount, 20);
  assert.equal(s.summaryCategories.CRC.baleCount, 4);
  assert.equal(s.summaryCategories.PSHRB.baleCount, 2);
  assert.equal(s.summaryCategories.CONV.baleCount, 1);
});

run("daily summary block parsing keeps CA fields in CA block and maps PB block to PCR", () => {
  const chat = [
    "2026/03/08, 18:40 - Tester: *Daily summary:* 17/02/2026",
    "",
    "12 - CA (Agricultural)",
    "Agricultural - 413",
    "Sidewalls - 463",
    "10839 kg",
    "10.839 tons",
    "",
    "06 - PB (Passenger)",
    "Motorcycle - 1122",
    "5420 kg",
    "5.420 tons",
    "",
    "Baler Machine One - 11 Bales",
    "Baler Machine Two - 07 Bales",
    "Total Bales - 18",
  ].join("\n");

  const parsed = parseBalingMessages(chat);
  assert.equal(parsed.summaryRecords.length, 1);
  const s = parsed.summaryRecords[0];

  // Date should come from body summary line, not WhatsApp timestamp fallback.
  assert.equal(s.chatDateParsed.year, 2026);
  assert.equal(s.chatDateParsed.month, 2);
  assert.equal(s.chatDateParsed.day, 17);

  // CA block
  assert.equal(s.summaryCategories.CA.baleCount, 12);
  assert.equal(s.summaryCategories.CA.agriT, 413);
  assert.equal(s.summaryCategories.CA.agriSW, 463);
  assert.equal(s.summaryCategories.CA.weightKg, 10839);
  assert.equal(s.summaryCategories.CA.tons, 10.839);

  // PB block should map to PCR and keep block-local values.
  assert.equal(s.summaryCategories.PCR.baleCount, 6);
  assert.equal(s.summaryCategories.PCR.motorcycleQty, 1122);
  assert.equal(s.summaryCategories.PCR.weightKg, 5420);
  assert.equal(s.summaryCategories.PCR.tons, 5.42);
});

run("daily summary CN block maps Light Commercial + Sidewalls into Nylon T/SW", () => {
  const chat = [
    "2026/01/05, 20:00 - Tester: *Daily summary:* 05/01/2026",
    "",
    "06 - CN (Cut Nylons)",
    "Light Commercial - 365",
    "Sidewalls - 193",
    "5467 kg",
    "5.467 tons",
  ].join("\n");

  const parsed = parseBalingMessages(chat);
  assert.equal(parsed.summaryRecords.length, 1);
  const s = parsed.summaryRecords[0];
  assert.equal(s.summaryCategories.CN.baleCount, 6);
  assert.equal(s.summaryCategories.CN.nylonT, 365);
  assert.equal(s.summaryCategories.CN.nylonSW, 193);
  assert.equal(s.summaryCategories.CN.lcQty, null);
  assert.equal(s.summaryCategories.CN.lcSW, null);
  assert.equal(s.summaryCategories.CN.weightKg, 5467);
  assert.equal(s.summaryCategories.CN.tons, 5.467);
});
