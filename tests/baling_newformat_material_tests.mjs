import assert from "node:assert/strict";
import { parseBalingMessagesNew } from "../src/parsing/balingParserNew.js";
import { createBalingWorkbook } from "../src/excel/createBalingWorkbook.js";

function run(name, fn) {
  Promise.resolve()
    .then(fn)
    .then(() => console.log(`PASS ${name}`))
    .catch((err) => {
      console.error(`FAIL ${name}`);
      console.error(err.message);
      process.exitCode = 1;
    });
}

run("PCR + LC: N maps to Light Commercial whole (not null)", async () => {
  const chat = [
    "2026/01/09, 08:20 - QA Bot: 09/01/2026",
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
  ].join("\n");

  const parsed = parseBalingMessagesNew(chat);
  assert.equal(parsed.standardRecords.length, 1);
  const row = parsed.standardRecords[0];
  assert.equal(row.baleNumber, "PCR003");
  assert.equal(row.lcQty, 80);
  assert.equal(row.lcWholeQty, null);

  const wb = await createBalingWorkbook(parsed);
  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);

  // O column = Light Commercial (whole under PCR block)
  assert.equal(ws.getRow(4).getCell(15).value, 80);
  // R column = Light Commercial (Whole) under RADIALS should stay blank for PCR.
  assert.equal(ws.getRow(4).getCell(18).value, "");
  // Total Number of Tyres should also reflect the 80 whole tyres
  assert.equal(ws.getRow(4).getCell(29).value, 80);
});

run("Explicit RADIALS component lines (LC T/SW, HC T/SW) populate RADIALS columns", async () => {
  const chat = [
    "2026/01/10, 13:02 - QA Bot: 10/01/2026",
    "BM #: 2",
    "",
    "CRS002-01/2026",
    "Operator: Bekinkosi",
    "Ass: Bhekisisa",
    "Start: 13:10",
    "Finish: 13:33",
    "LC Full: 10",
    "LC T: 20",
    "LC SW: 30",
    "HC Full: 5",
    "HC T: 12",
    "HC SW: 8",
    "Weight: 980 kg",
  ].join("\n");

  const parsed = parseBalingMessagesNew(chat);
  assert.equal(parsed.standardRecords.length, 1);
  const row = parsed.standardRecords[0];
  assert.equal(row.lcTread, 20);
  assert.equal(row.lcSideWall, 30);
  assert.equal(row.lcWholeQty, 10);
  assert.equal(row.hcTread, 12);
  assert.equal(row.hcSideWall, 8);
  assert.equal(row.hcWholeQty, 5);
  assert.equal(row.otherItemRaw, "");

  const wb = await createBalingWorkbook(parsed);
  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  assert.equal(ws.getRow(4).getCell(16).value, 20); // LC T
  assert.equal(ws.getRow(4).getCell(17).value, 30); // LC SW
  assert.equal(ws.getRow(4).getCell(18).value, 10); // LC Whole
  assert.equal(ws.getRow(4).getCell(19).value, 12); // HC T
  assert.equal(ws.getRow(4).getCell(20).value, 8);  // HC SW
  assert.equal(ws.getRow(4).getCell(21).value, 5);  // HC Whole
});

run("CN with LC whole + SW maps to Nylon T and Nylon SW", async () => {
  const chat = [
    "2026/01/10, 11:40 - QA Bot: 10/01/2026",
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
  ].join("\n");

  const parsed = parseBalingMessagesNew(chat);
  assert.equal(parsed.standardRecords.length, 1);
  const wb = await createBalingWorkbook(parsed);
  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  // CN should push into NYLONS subcategory.
  assert.equal(ws.getRow(4).getCell(26).value, 38); // Nylon T
  assert.equal(ws.getRow(4).getCell(27).value, 19); // Nylon SW
});

run("CA with bare T/SW maps to Agricultural T/SW", async () => {
  const chat = [
    "2026/01/11, 08:55 - QA Bot: 11/01/2026",
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
  ].join("\n");

  const parsed = parseBalingMessagesNew(chat);
  assert.equal(parsed.standardRecords.length, 1);
  const row = parsed.standardRecords[0];
  assert.equal(row.agriTread, 44);
  assert.equal(row.agriSideWall, 26);

  const wb = await createBalingWorkbook(parsed);
  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  assert.equal(ws.getRow(4).getCell(22).value, 44); // Agricultural T
  assert.equal(ws.getRow(4).getCell(23).value, 26); // Agricultural SW
});

run("TB with TB: N maps to Tubes and total", async () => {
  const chat = [
    "2026/01/11, 14:05 - QA Bot: 11/01/2026",
    "BM #: 1",
    "",
    "TB001-01/2026",
    "Operator: Menzi",
    "Ass: Langa",
    "Start: 13:31",
    "Finish: 13:57",
    "TB: 14",
    "Weight: 734 kg",
  ].join("\n");

  const parsed = parseBalingMessagesNew(chat);
  assert.equal(parsed.standardRecords.length, 1);
  const row = parsed.standardRecords[0];
  assert.equal(row.tubeQty, 14);
  assert.equal(row.newFormatTotalQty, 14);
  assert.equal(row.totalQty, 14);

  const wb = await createBalingWorkbook(parsed);
  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  assert.equal(ws.getRow(4).getCell(28).value, 14); // Tubes
  assert.equal(ws.getRow(4).getCell(29).value, 14); // Total Number of Tyres
});

run("PShrB maps machine/type/number/series/weight with other tyre fields blank (case-insensitive)", async () => {
  const chat = [
    "2026/01/12, 10:05 - QA Bot: 12/01/2026",
    "BM #: 1",
    "",
    "pshrb001-01/2026",
    "Weight: 1020 kg",
  ].join("\n");

  const parsed = parseBalingMessagesNew(chat);
  assert.equal(parsed.standardRecords.length, 1);
  const row = parsed.standardRecords[0];
  assert.equal(row.machine, "BM - 1");
  assert.equal(row.baleNumber, "PSHRB001");
  assert.equal(row.baleSeries, "01/2026");
  assert.equal(row.weightKg, 1020);
  assert.equal(row.passengerQty, null);
  assert.equal(row.fourx4Qty, null);
  assert.equal(row.motorcycleQty, null);
  assert.equal(row.lcQty, null);
  assert.equal(row.hcQty, null);
  assert.equal(row.agriQty, null);
  assert.equal(row.treadQty, null);
  assert.equal(row.sideWallQty, null);
  assert.equal(row.tubeQty, null);

  const wb = await createBalingWorkbook(parsed);
  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  assert.equal(ws.getRow(4).getCell(3).value, "BM - 1");   // Machine
  assert.equal(ws.getRow(4).getCell(4).value, "PShrB");    // Bale Type
  assert.equal(ws.getRow(4).getCell(5).value, "001");      // Bale Number
  assert.equal(ws.getRow(4).getCell(6).value, "01/2026");  // Bale Series
  assert.equal(ws.getRow(4).getCell(30).value, 1020);      // Bale Weight KG
  assert.equal(ws.getRow(4).getCell(31).value, 1.02);      // Bale Weight TONS
  assert.equal(ws.getRow(4).getCell(12).value, "");        // Passenger blank
  assert.equal(ws.getRow(4).getCell(13).value, "");        // 4x4 blank
  assert.equal(ws.getRow(4).getCell(14).value, "");        // Motorcycle blank
});

run("new-format labels parse with mixed case and '=' separators", async () => {
  const chat = [
    "2026/01/13, 14:40 - QA Bot: 13/01/2026",
    "bm#=2",
    "",
    "pShRb007-01/2026",
    "oPeRaToR= Menzi",
    "aSsIsTaNt= Langa",
    "sTaRt= 13:31",
    "fInIsH= 13:57",
    "wEiGhT: 734 kg",
  ].join("\n");

  const parsed = parseBalingMessagesNew(chat);
  assert.equal(parsed.standardRecords.length, 1);
  const row = parsed.standardRecords[0];
  assert.equal(row.machine, "BM - 2");
  assert.equal(row.baleNumber, "PSHRB007");
  assert.equal(row.baleSeries, "01/2026");
  assert.equal(row.operator, "Menzi");
  assert.equal(row.assistant, "Langa");
  assert.equal(row.durationMinutes, 26);
  assert.equal(row.weightKg, 734);
});
