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
  // column 28 = Total Number of Tyres
  assert.equal(ws.getRow(4).getCell(28).value, 46);
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
  assert.equal(ws.getRow(4).getCell(5).value, "B100");
  assert.equal(ws.getRow(5).getCell(5).value, "B3");
  assert.equal(ws.getRow(6).getCell(5).value, "B20");
  assert.equal(ws.getRow(7).getCell(5).value, "B100");
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
  assert.equal(ws.getRow(3).getCell(12).value, "PCR");
  assert.equal(ws.getRow(3).getCell(16).value, "RADIALS");
  assert.equal(ws.getRow(3).getCell(22).value, "NYLONS");
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
  assert.equal(ws.getRow(2).getCell(28).value, "Total Number of Tyres");
  assert.equal(ws.getRow(2).getCell(29).value, "Bale Weight\nKG");
  assert.equal(ws.getRow(2).getCell(30).value, "Bale Weight\nTONS");
  assert.equal(ws.getRow(4).getCell(30).value, 1.02);
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
  assert.equal(totalsRow.getCell(28).value, 30);
  assert.equal(totalsRow.getCell(29).value, 3000);
  assert.equal(totalsRow.getCell(30).value, 3);
  // 15:00 boundary => orange according to current rule.
  assert.equal(totalsRow.getCell(11).fill.fgColor.argb, "FFFFE699");
});
