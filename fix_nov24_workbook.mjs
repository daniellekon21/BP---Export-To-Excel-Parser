import ExcelJS from "exceljs";
import { writeFileSync } from "node:fs";

const INPUT = "../Excel_files/nov24_old_format.xlsx";
const OUTPUT = "../Excel_files/nov24_old_format_corrected.xlsx";
const REPORT = "../qa_report.md";
const SHEET = "Sheet1";

function cv(cell) {
  const val = cell.value;
  if (val && typeof val === "object" && "text" in val) return String(val.text || "").trim();
  return String(val ?? "").trim();
}

function rowObj(ws, r) {
  const row = ws.getRow(r);
  return {
    row: r,
    date: cv(row.getCell(1)),
    cm: cv(row.getCell(2)),
    start: cv(row.getCell(5)),
    finish: cv(row.getCell(6)),
    lc: cv(row.getCell(10)),
    tread_lc: cv(row.getCell(11)),
    hc: cv(row.getCell(12)),
    agri: cv(row.getCell(13)),
    tread_hc: cv(row.getCell(14)),
    nylon_total: cv(row.getCell(15)),
    tread_cuts: cv(row.getCell(16)),
  };
}

function rowValuesString(o) {
  return `date=${o.date || ""}, cm=${o.cm || ""}, time=${o.start || ""}-${o.finish || ""}, LC=${o.lc || ""}, HC=${o.hc || ""}, Agri=${o.agri || ""}, TreadCuts=${o.tread_cuts || ""}`;
}

function keyOf(o) {
  return `${o.date}|${o.start}|${o.finish}|${o.cm}`;
}

function clearQty(row) {
  for (const c of [10, 11, 12, 13, 14, 15, 16]) row.getCell(c).value = null;
}

function findRowByKey(ws, key) {
  for (let r = 3; r <= ws.rowCount; r += 1) {
    const o = rowObj(ws, r);
    if (keyOf(o) === key) return r;
  }
  return null;
}

const expected = new Map();
const put = (date, start, finish, cm, mode, qty, status = "fixed", note = "") => {
  expected.set(`${date}|${start}|${finish}|${cm}`, { mode, qty, status, note });
};

// 20/11 LC blocks
for (const [s, e, a, b, c] of [
  ["08:00", "09:00", 18, 23, 20],
  ["09:01", "10:00", 30, 34, 41],
  ["10:01", "11:00", 41, 48, 60],
  ["11:01", "12:00", 55, 55, 67],
  ["14:30", "15:30", 66, 76, 81],
  ["15:31", "16:30", 89, 83, 86],
]) {
  put("20/11/2024", s, e, "CM - 1", "lc", a, "unchanged");
  put("20/11/2024", s, e, "CM - 2", "lc", b, "unchanged");
  put("20/11/2024", s, e, "CM - 3", "lc", c, "unchanged");
}

// 21/11 trucks
for (const [s, e, a, b, c] of [
  ["08:00", "09:00", 45, 27, 27],
  ["09:01", "10:00", 51, 40, 40],
  ["10:01", "11:00", 65, 50, 73],
]) {
  put("21/11/2024", s, e, "CM - 1", "hc", a, "unchanged");
  put("21/11/2024", s, e, "CM - 2", "hc", b, "unchanged");
  put("21/11/2024", s, e, "CM - 3", "hc", c, "unchanged");
}

// 22/11
put("22/11/2024", "08:00", "09:00", "CM - 1", "hc", 30, "fixed", "from chat: Machine 1-30");
put("22/11/2024", "08:00", "09:00", "CM - 2", "hc", 4, "fixed", "from chat: Machine 2-4");
put("22/11/2024", "08:00", "09:00", "CM - 3", "na", null, "unchanged", "chat says N/A (offloading truck)");

// 28/11
put("28/11/2024", "08:00", "09:00", "CM - 1", "hc", 12, "ambiguous", "chat quantity given; type not explicit");
put("28/11/2024", "09:01", "10:00", "CM - 1", "hc", 28, "ambiguous", "chat quantity given; type not explicit");
put("28/11/2024", "10:01", "11:00", "CM - 1", "hc", 28, "fixed", "chat says truck");
put("28/11/2024", "10:01", "11:00", "CM - 2", "hc", 12, "fixed", "chat says truck");

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(INPUT);
const ws = wb.getWorksheet(SHEET);
if (!ws) throw new Error(`Worksheet ${SHEET} not found`);

const originalRows = [];
for (let r = 3; r <= ws.rowCount; r += 1) {
  const o = rowObj(ws, r);
  if (o.date) originalRows.push(o);
}

const qa = [];

// Deletions first: rows without expected chat entry, including blank placeholders
const toDelete = [];
for (const o of originalRows) {
  const key = keyOf(o);
  if (!expected.has(key)) {
    toDelete.push(o);
    qa.push({
      sheet: SHEET,
      row: o.row,
      original: rowValuesString(o),
      expected: "No row (no source chat entry)",
      action: "Deleted placeholder/unsupported row",
      status: "fixed",
    });
  }
}

// Apply deletes from bottom
for (const o of [...toDelete].sort((a, b) => b.row - a.row)) ws.spliceRows(o.row, 1);

// Apply updates/validations for expected rows and produce QA entries for all remaining original rows
for (const o of originalRows) {
  const key = keyOf(o);
  const exp = expected.get(key);
  if (!exp) continue;

  const curRowNum = findRowByKey(ws, key);
  if (curRowNum == null) {
    qa.push({
      sheet: SHEET,
      row: o.row,
      original: rowValuesString(o),
      expected: `Missing row for ${key}`,
      action: "Could not update; row not found after deletions",
      status: "ambiguous",
    });
    continue;
  }

  const row = ws.getRow(curRowNum);
  const before = rowObj(ws, curRowNum);

  // Desired values
  let desiredLC = "";
  let desiredHC = "";
  if (exp.mode === "lc") desiredLC = String(exp.qty);
  if (exp.mode === "hc") desiredHC = String(exp.qty);

  const isAlready = before.lc === desiredLC && before.hc === desiredHC && before.agri === "" && before.tread_cuts === "";

  // enforce expected structure
  clearQty(row);
  if (exp.mode === "lc") row.getCell(10).value = exp.qty;
  if (exp.mode === "hc") row.getCell(12).value = exp.qty;
  row.commit();

  qa.push({
    sheet: SHEET,
    row: o.row,
    original: rowValuesString(o),
    expected: `date=${o.date}, cm=${o.cm}, time=${o.start}-${o.finish}, LC=${desiredLC}, HC=${desiredHC}${exp.note ? ` (${exp.note})` : ""}`,
    action: exp.mode === "na"
      ? "Kept N/A row with blank qty"
      : isAlready
        ? "No change"
        : `Updated qty placement/value to match chat (${exp.mode.toUpperCase()})`,
    status: exp.mode === "na" ? exp.status : (isAlready ? "unchanged" : exp.status),
  });
}

qa.push({
  sheet: SHEET,
  row: "scope",
  original: "Workbook date range is 20/11/2024..28/11/2024",
  expected: "02/12/2024 entries excluded",
  action: "No rows added for 02/12/2024 because workbook is Nov-2024 scoped",
  status: "unchanged",
});

await wb.xlsx.writeFile(OUTPUT);

qa.sort((a, b) => {
  const ar = Number.isFinite(Number(a.row)) ? Number(a.row) : 999999;
  const br = Number.isFinite(Number(b.row)) ? Number(b.row) : 999999;
  if (ar !== br) return ar - br;
  return String(a.status).localeCompare(String(b.status));
});

const reportLines = [];
reportLines.push("# QA Report - nov24_old_format.xlsx");
reportLines.push("");
reportLines.push("Source of truth: `WhatsApp_chats/nov_2024_whatsapp.txt`");
reportLines.push("Input workbook: `Excel_files/nov24_old_format.xlsx`");
reportLines.push("Corrected workbook: `Excel_files/nov24_old_format_corrected.xlsx`");
reportLines.push("");
reportLines.push("| sheet | row | original values | expected values from chat | action taken | status |");
reportLines.push("|---|---:|---|---|---|---|");
for (const q of qa) {
  const esc = (s) => String(s).replaceAll("|", "\\|");
  reportLines.push(`| ${esc(q.sheet)} | ${esc(q.row)} | ${esc(q.original)} | ${esc(q.expected)} | ${esc(q.action)} | ${esc(q.status)} |`);
}

writeFileSync(REPORT, reportLines.join("\n"), "utf8");
console.log(`Wrote: ${OUTPUT}`);
console.log(`Wrote: ${REPORT}`);
