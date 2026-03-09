import assert from "node:assert/strict";
import { createBalingWorkbook } from "../src/excel/createBalingWorkbook.js";

function mkRow({ y, m, d, bale, raw = "raw" } = {}) {
  const date = (y && m && d) ? { year: y, month: m, day: d } : null;
  return {
    sourceTimestamp: "2024/09/23",
    chatDateParsed: date,
    date,
    machine: "Machine One",
    baleNumber: bale || "B1",
    recordType: "STANDARD",
    productionType: "Production",
    rawMessage: raw,
  };
}

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

run("1) different months land in different month sheets", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      mkRow({ y: 2024, m: 9, d: 23, bale: "B100" }),
      mkRow({ y: 2024, m: 10, d: 1, bale: "B101" }),
    ],
  });

  const names = wb.worksheets.map((w) => w.name);
  assert.ok(names.includes("Sep 2024"));
  assert.ok(names.includes("Oct 2024"));
});

run("2) same-month rows are grouped into one sheet", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      mkRow({ y: 2025, m: 1, d: 2, bale: "B200" }),
      mkRow({ y: 2025, m: 1, d: 10, bale: "B201" }),
    ],
  });

  const jan = wb.getWorksheet("Jan 2025");
  assert.ok(jan);
  // 3 header rows + 2 data rows + 1 totals row
  assert.equal(jan.rowCount, 6);
});

run("3) month sheets are in chronological order", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      mkRow({ y: 2025, m: 2, d: 1, bale: "B300" }),
      mkRow({ y: 2024, m: 9, d: 1, bale: "B301" }),
      mkRow({ y: 2024, m: 12, d: 1, bale: "B302" }),
    ],
  });

  const names = wb.worksheets.map((w) => w.name);
  const monthNames = names.filter((n) => ["Sep 2024", "Dec 2024", "Feb 2025"].includes(n));
  assert.deepEqual(monthNames, ["Sep 2024", "Dec 2024", "Feb 2025"]);
});

run("4) missing/invalid dates go to review sheet and are logged", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      mkRow({ y: 2025, m: 1, d: 5, bale: "B400" }),
      mkRow({ bale: "B401", raw: "Missing date row" }),
    ],
  });

  const review = wb.getWorksheet("Review_Missing_Date");
  assert.ok(review);
  assert.equal(review.rowCount, 2); // header + 1 invalid row

  const log = wb.getWorksheet("Validation_Log");
  const issues = [];
  for (let i = 2; i <= log.rowCount; i += 1) {
    issues.push(String(log.getRow(i).getCell(2).value || ""));
  }
  assert.ok(issues.includes("MISSING_OR_INVALID_DATE"));
});

run("5) no production row is silently lost", async () => {
  const rows = [
    mkRow({ y: 2024, m: 9, d: 2, bale: "B500" }),
    mkRow({ y: 2024, m: 9, d: 3, bale: "B501" }),
    mkRow({ y: 2024, m: 10, d: 4, bale: "B502" }),
    mkRow({ bale: "B503", raw: "No date" }),
  ];

  const wb = await createBalingWorkbook({ standardRecords: rows });
  const prodCount = wb.worksheets
    .filter((w) => /^([A-Z][a-z]{2} \d{4})$/.test(w.name))
    .reduce((sum, w) => sum + Math.max(0, w.rowCount - 4), 0);
  const review = wb.getWorksheet("Review_Missing_Date");
  const reviewCount = review ? Math.max(0, review.rowCount - 1) : 0;

  assert.equal(prodCount + reviewCount, rows.length);
});

run("6) month tabs are created from non-standard production families too", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [mkRow({ y: 2024, m: 10, d: 10, bale: "B600" })],
    crcaRecords: [{
      sourceTimestamp: "2026/03/04",
      chatDateParsed: { year: 2026, month: 3, day: 4 },
      date: { year: 2026, month: 3, day: 4 },
      machine: "Machine Two",
      baleTestCode: "CR-01",
      recordType: "CR_CA",
      testType: "Test",
      rawMessage: "CR test row",
    }],
  });

  const names = wb.worksheets.map((w) => w.name);
  assert.ok(names.includes("Oct 2024"));
  assert.ok(names.includes("Mar 2026"));
});

run("7) total qty is auto-corrected to computed category sum", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [{
      ...mkRow({ y: 2024, m: 10, d: 18, bale: "B700" }),
      passengerQty: 10,
      lcQty: 36,
      totalQty: 43, // user-declared wrong total
    }],
  });

  const ws = wb.getWorksheet("Oct 2024");
  assert.ok(ws);
  // column 29 = Total Number of Tyres
  assert.equal(ws.getRow(4).getCell(29).value, 46);
});

run("7b) passenger + 4x4 only keeps other tyre component columns blank", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [{
      ...mkRow({ y: 2024, m: 10, d: 19, bale: "B701" }),
      passengerQty: 110,
      fourx4Qty: 2,
    }],
  });

  const ws = wb.getWorksheet("Oct 2024");
  assert.ok(ws);
  const row = ws.getRow(4);
  assert.equal(row.getCell(12).value, 110); // Passenger
  assert.equal(row.getCell(13).value, 2);   // 4x4
  assert.equal(row.getCell(29).value, 112); // Total Number of Tyres

  // All other tyre-material component columns remain blank.
  assert.equal(row.getCell(14).value, ""); // Motorcycle
  assert.equal(row.getCell(15).value, ""); // LC
  assert.equal(row.getCell(16).value, ""); // LC T
  assert.equal(row.getCell(17).value, ""); // LC SW
  assert.equal(row.getCell(18).value, ""); // LC Whole
  assert.equal(row.getCell(19).value, ""); // HC T
  assert.equal(row.getCell(20).value, ""); // HC SW
  assert.equal(row.getCell(21).value, ""); // HC Whole
  assert.equal(row.getCell(22).value, ""); // Agricultural T
  assert.equal(row.getCell(23).value, ""); // Agricultural SW
  assert.equal(row.getCell(24).value, ""); // Total Number of T
  assert.equal(row.getCell(25).value, ""); // Total Number of SW
  assert.equal(row.getCell(26).value, ""); // Nylon T
  assert.equal(row.getCell(27).value, ""); // Nylon SW
  assert.equal(row.getCell(28).value, ""); // Tubes
});

run("8) rows are sorted by date then bale number (numeric)", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      { ...mkRow({ y: 2024, m: 9, d: 25, bale: "B20" }) },
      { ...mkRow({ y: 2024, m: 9, d: 25, bale: "B3" }) },
      { ...mkRow({ y: 2024, m: 9, d: 24, bale: "B100" }) },
      { ...mkRow({ y: 2024, m: 9, d: 25, bale: "B100" }) },
    ],
  });

  const ws = wb.getWorksheet("Sep 2024");
  assert.ok(ws);
  // col 2 = date, col 5 = bale number
  assert.equal(ws.getRow(4).getCell(2).value, "24/09/2024");
  assert.equal(ws.getRow(4).getCell(5).value, "100");
  assert.equal(ws.getRow(5).getCell(5).value, "3");
  assert.equal(ws.getRow(6).getCell(5).value, "20");
  assert.equal(ws.getRow(7).getCell(5).value, "100");
});

run("9) start/finish times are exported as HH:MM:SS (default :00)", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2024, m: 9, d: 25, bale: "B800" }),
        startTime: { h: 9, m: 5 },
        finishTime: { h: 9, m: 28, s: 13 },
      },
    ],
  });

  const ws = wb.getWorksheet("Sep 2024");
  assert.ok(ws);
  // col 9 = Start Time, col 10 = Finish Time
  assert.equal(ws.getRow(4).getCell(9).value, "09:05:00");
  assert.equal(ws.getRow(4).getCell(10).value, "09:28:13");
  // col 11 = Total Time to Bale
  assert.equal(ws.getRow(2).getCell(11).value, "Total Time to Bale");
  assert.equal(ws.getRow(4).getCell(11).value, "00:23:13");
});

run("10) Total Time to Bale color traffic-light is applied", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2024, m: 9, d: 25, bale: "B801" }),
        startTime: { h: 9, m: 0, s: 0 },
        finishTime: { h: 9, m: 10, s: 0 }, // green
      },
      {
        ...mkRow({ y: 2024, m: 9, d: 25, bale: "B802" }),
        startTime: { h: 10, m: 0, s: 0 },
        finishTime: { h: 10, m: 16, s: 0 }, // orange
      },
      {
        ...mkRow({ y: 2024, m: 9, d: 25, bale: "B803" }),
        startTime: { h: 11, m: 0, s: 0 },
        finishTime: { h: 11, m: 22, s: 0 }, // red
      },
    ],
  });

  const ws = wb.getWorksheet("Sep 2024");
  assert.ok(ws);
  assert.equal(ws.getRow(4).getCell(11).fill.fgColor.argb, "FFC6EFCE");
  assert.equal(ws.getRow(5).getCell(11).fill.fgColor.argb, "FFFFE699");
  assert.equal(ws.getRow(6).getCell(11).fill.fgColor.argb, "FFF4CCCC");
});

run("11) header follows example input-material wording and sublabels", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [mkRow({ y: 2024, m: 9, d: 25, bale: "B900" })],
  });

  const ws = wb.getWorksheet("Sep 2024");
  assert.ok(ws);
  assert.equal(ws.getRow(2).getCell(12).value, "Passenger");
  assert.equal(ws.getRow(2).getCell(13).value, "4 X 4");
  assert.equal(ws.getRow(2).getCell(15).value, "Light Commercial");
  assert.equal(ws.getRow(2).getCell(18).value, "Light Commercial (Whole)");
  assert.equal(ws.getRow(2).getCell(21).value, "Heavy Commercial (Whole)");
  assert.equal(ws.getRow(2).getCell(22).value, "Agricultural T");
  assert.equal(ws.getRow(2).getCell(23).value, "Agricultural SW");
  assert.equal(ws.getRow(2).getCell(24).value, "Total Number of T");
  assert.equal(ws.getRow(2).getCell(25).value, "Total Number of SW");
  assert.equal(ws.getRow(2).getCell(26).value, "Nylon T");
  assert.equal(ws.getRow(2).getCell(27).value, "Nylon SW");
  assert.equal(ws.getRow(2).getCell(28).value, "Tubes");
  assert.equal(ws.getRow(2).getCell(29).value, "Total Number of Tyres");
  assert.equal(ws.getRow(3).getCell(12).value, "PCR");
  assert.equal(ws.getRow(3).getCell(16).value, "RADIALS");
  assert.equal(ws.getRow(3).getCell(26).value, "NYLONS");
});

run("12) production headers renamed and Bale Weight TONS is calculated from KG", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2024, m: 10, d: 18, bale: "B901" }),
        weightKg: 1020,
      },
    ],
  });

  const ws = wb.getWorksheet("Oct 2024");
  assert.ok(ws);
  assert.equal(ws.getRow(2).getCell(29).value, "Total Number of Tyres");
  assert.equal(ws.getRow(2).getCell(30).value, "Bale Weight\nKG");
  assert.equal(ws.getRow(2).getCell(31).value, "Bale Weight\nTONS");
  assert.equal(ws.getRow(4).getCell(31).value, 1.02);
});

run("12b) body date text is exported in the last production column", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2024, m: 10, d: 18, bale: "B902" }),
        bodyDateText: "17/10/2024",
      },
    ],
  });

  const ws = wb.getWorksheet("Oct 2024");
  assert.ok(ws);
  assert.equal(ws.getRow(2).getCell(33).value, "Body Date Text");
  assert.equal(ws.getRow(4).getCell(33).value, "17/10/2024");
});

run("12c) daily summaries are appended to month sheet and no separate summary sheet exists", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      { ...mkRow({ y: 2024, m: 10, d: 18, bale: "B903" }) },
    ],
    summaryRecords: [
      {
        sourceTimestamp: "2024/10/18",
        chatDateParsed: { year: 2024, month: 10, day: 18 },
        date: { year: 2024, month: 10, day: 18 },
        summaryType: "BALING_DAILY_SUMMARY",
        baleCount: 1,
        rawMessage: "Daily summary text",
      },
    ],
  });

  const ws = wb.getWorksheet("Oct 2024");
  assert.ok(ws);
  let foundSummaryTable = false;
  for (let i = 1; i <= ws.rowCount; i += 1) {
    if (ws.getRow(i).getCell(2).value === "Daily Summaries") {
      foundSummaryTable = true;
      break;
    }
  }
  assert.ok(foundSummaryTable);
  assert.equal(wb.getWorksheet("Daily_Summaries"), undefined);
});

run("12d) future body-date fallback highlights Date cell in pastel yellow", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2024, m: 10, d: 25, bale: "B904" }),
        bodyDateText: "25/10/2034",
        usedTimestampFallbackForFutureBodyDate: true,
      },
    ],
  });

  const ws = wb.getWorksheet("Oct 2024");
  assert.ok(ws);
  const dateCell = ws.getRow(4).getCell(2);
  assert.equal(dateCell.fill.fgColor.argb, "FFFFF2CC");
});

run("13) totals row computes tyre/weight sums and average baling time", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2024, m: 10, d: 18, bale: "B910" }),
        startTime: { h: 10, m: 0, s: 0 },
        finishTime: { h: 10, m: 10, s: 0 },
        passengerQty: 10,
        weightKg: 1000,
      },
      {
        ...mkRow({ y: 2024, m: 10, d: 18, bale: "B911" }),
        startTime: { h: 10, m: 20, s: 0 },
        finishTime: { h: 10, m: 40, s: 0 },
        passengerQty: 20,
        weightKg: 2000,
      },
    ],
  });

  const ws = wb.getWorksheet("Oct 2024");
  assert.ok(ws);
  const totalsRow = ws.getRow(ws.rowCount);
  assert.equal(totalsRow.getCell(2).value, "TOTALS");
  assert.equal(totalsRow.getCell(11).value, "00:15:00");
  assert.equal(totalsRow.getCell(12).value, 30);
  assert.equal(totalsRow.getCell(29).value, 30);
  assert.equal(totalsRow.getCell(30).value, 3000);
  assert.equal(totalsRow.getCell(31).value, 3);
  // 15:00 boundary => orange according to current rule.
  assert.equal(totalsRow.getCell(11).fill.fgColor.argb, "FFFFE699");
});

run("14) when SW exists, LC is mapped to LC T and LC whole is 0", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2026, m: 1, d: 5, bale: "CRC001" }),
        lcQty: 56,
        sideWallQty: 18,
        weightKg: 934,
      },
    ],
  });

  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  // O = Light Commercial, P = Light Commercial T, Q = Light Commercial SW
  assert.equal(ws.getRow(4).getCell(15).value, "");
  assert.equal(ws.getRow(4).getCell(16).value, 56);
  assert.equal(ws.getRow(4).getCell(17).value, 18);
  // LC maps to treads in this case; total uses max(SW/2, T).
  assert.equal(ws.getRow(4).getCell(29).value, 56);
});

run("15) bale type prefix mapping uses expected business labels and Bale Number shows numeric sequence only", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      { ...mkRow({ y: 2026, m: 1, d: 5, bale: "CA015" }) },
      { ...mkRow({ y: 2026, m: 1, d: 5, bale: "B200" }) },
      { ...mkRow({ y: 2026, m: 1, d: 5, bale: "PB300" }) },
      { ...mkRow({ y: 2026, m: 1, d: 5, bale: "CR100" }) },
      { ...mkRow({ y: 2026, m: 1, d: 5, bale: "SR900" }) },
      { ...mkRow({ y: 2026, m: 1, d: 5, bale: "CN010" }) },
      { ...mkRow({ y: 2026, m: 1, d: 5, bale: "ConV777" }) },
      { ...mkRow({ y: 2026, m: 1, d: 5, bale: "PShrB123" }) },
    ],
  });

  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  const typeByNumber = {};
  for (let r = 4; r < ws.rowCount; r += 1) {
    const code = String(ws.getRow(r).getCell(5).value || "");
    const type = String(ws.getRow(r).getCell(4).value || "");
    if (code) typeByNumber[code] = type;
  }
  assert.equal(typeByNumber["015"], "CA");
  assert.equal(typeByNumber["200"], "PCR");
  assert.equal(typeByNumber["300"], "PCR");
  assert.equal(typeByNumber["100"], "CRC");
  assert.equal(typeByNumber["900"], "CRS");
  assert.equal(typeByNumber["010"], "CN");
  assert.equal(typeByNumber["777"], "ConV");
  assert.equal(typeByNumber["123"], "PShrB");
});

run("16) CN + LC maps T/SW into Nylon subcategory, not RADIALS totals", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2026, m: 1, d: 5, bale: "CN001" }),
        lcQty: 56,
        treadQty: 43,     // T
        sideWallQty: 18,  // SW
      },
    ],
  });

  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  // P/Q (16/17) = RADIALS LC T/SW should be zero for CN.
  assert.equal(ws.getRow(4).getCell(16).value, "");
  assert.equal(ws.getRow(4).getCell(17).value, "");
  // V/W (22/23) = RADIALS Total T/SW should be zero for CN.
  assert.equal(ws.getRow(4).getCell(24).value, "");
  assert.equal(ws.getRow(4).getCell(25).value, "");
  // X/Y (24/25) = Nylon T/Nylon SW should hold T/SW.
  assert.equal(ws.getRow(4).getCell(26).value, 43);
  assert.equal(ws.getRow(4).getCell(27).value, 18);
});

run("17) total tyres from T and SW uses max(SW/2, T)", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2026, m: 1, d: 6, bale: "CRC100" }),
        treadQty: 20,
        sideWallQty: 70, // SW/2 = 35 > T(20)
      },
    ],
  });

  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  // AB column in current layout = Total Number of Tyres
  assert.equal(ws.getRow(4).getCell(29).value, 35);
});

run("18) CR/CRC keeps RADIALS LC T/SW and forces NYLONS T/SW to zero", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2026, m: 1, d: 5, bale: "CR001" }),
        lcQty: 71,
        sideWallQty: 23,
      },
    ],
  });

  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  // RADIALS LC T/SW
  assert.equal(ws.getRow(4).getCell(16).value, 71);
  assert.equal(ws.getRow(4).getCell(17).value, 23);
  // NYLONS T/SW should be zero
  assert.equal(ws.getRow(4).getCell(26).value, "");
  assert.equal(ws.getRow(4).getCell(27).value, "");
});

run("19) CA maps T/SW into Agricultural T/SW", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2026, m: 1, d: 7, bale: "CA015" }),
        treadQty: 45,
        sideWallQty: 32,
      },
    ],
  });
  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  // T/U = Agricultural T/SW
  assert.equal(ws.getRow(4).getCell(22).value, 45);
  assert.equal(ws.getRow(4).getCell(23).value, 32);
  // V/W total T/SW should be sum of RADIALS T/SW columns.
  assert.equal(ws.getRow(4).getCell(24).value, 45);
  assert.equal(ws.getRow(4).getCell(25).value, 32);
});

run("20) CRS with LC maps LC T/SW and HC T/SW as available", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2026, m: 1, d: 8, bale: "CRS011" }),
        lcQty: 30,
        hcQty: 12,
        sideWallQty: 14,
      },
    ],
  });
  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  // LC T/SW
  assert.equal(ws.getRow(4).getCell(16).value, 30);
  assert.equal(ws.getRow(4).getCell(17).value, 14);
  // HC T/SW
  assert.equal(ws.getRow(4).getCell(19).value, 12);
  assert.equal(ws.getRow(4).getCell(20).value, 14);
  // RADIALS totals = LC + HC + Agri per column type
  assert.equal(ws.getRow(4).getCell(24).value, 42);
  assert.equal(ws.getRow(4).getCell(25).value, 28);
});

run("21) when LC and HC both exist, total tyres uses max(SW/2, LC+HC)", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2026, m: 1, d: 10, bale: "CRS022" }),
        lcQty: 56,
        hcQty: 12,
        sideWallQty: 18, // SW/2 = 9, LC+HC = 68 -> expect 68
      },
    ],
  });

  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  assert.equal(ws.getRow(4).getCell(29).value, 68);
});

run("22) when Agri also exists, total tyres uses agri part + (LC/HC part)", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2026, m: 1, d: 11, bale: "CRS023" }),
        lcQty: 56,
        hcQty: 12,
        sideWallQty: 18, // HC/LC SW context => max(18/2, 56+12) = 68
        agriQty: 30,     // max(30, 18/2) = 30
      },
    ],
  });

  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  assert.equal(ws.getRow(4).getCell(29).value, 98);
});

run("23) ConV makes the whole production row red font", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2026, m: 1, d: 9, bale: "ConV001" }),
        weightKg: 800,
        rawMessage: "Conveyor bale test text",
      },
    ],
  });
  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  for (let c = 1; c <= 33; c += 1) {
    assert.equal(ws.getRow(4).getCell(c).font.color.argb, "FFFF0000");
  }
});

run("24) CRS with HC Full maps HC whole + HC T/SW and total uses whole + max(SW/2, T)", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2026, m: 1, d: 9, bale: "CRS001" }),
        hcWholeQty: 15,
        hcQty: 28,
        sideWallQty: 41,
      },
    ],
  });

  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  // HC T / SW / Whole
  assert.equal(ws.getRow(4).getCell(19).value, 28);
  assert.equal(ws.getRow(4).getCell(20).value, 41);
  assert.equal(ws.getRow(4).getCell(21).value, 15);
  // Total tyres = HC whole + max(41/2, 28) = 15 + 28 = 43
  assert.equal(ws.getRow(4).getCell(29).value, 43);
});

run("25) when both LC and HC have whole+cut, total tyres sums LC total and HC total", async () => {
  const wb = await createBalingWorkbook({
    standardRecords: [
      {
        ...mkRow({ y: 2026, m: 1, d: 10, bale: "CRS002" }),
        lcWholeQty: 10,
        lcQty: 20,
        lcSideWall: 30, // LC total = 10 + max(15, 20) = 30
        hcWholeQty: 5,
        hcQty: 12,
        hcSideWall: 8, // HC total = 5 + max(4, 12) = 17
      },
    ],
  });

  const ws = wb.getWorksheet("Jan 2026");
  assert.ok(ws);
  assert.equal(ws.getRow(4).getCell(18).value, 10); // LC Whole
  assert.equal(ws.getRow(4).getCell(21).value, 5);  // HC Whole
  assert.equal(ws.getRow(4).getCell(29).value, 47); // 30 + 17
});
