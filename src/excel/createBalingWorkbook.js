import { dateSortKey, dateToStr, monthLabel } from "../helpers.js";
import {
  BALING_SHEET_NAMES,
  BALING_FAILED_HEADERS,
  BALING_SCRAP_HEADERS,
  BALING_CRCA_HEADERS,
  BALING_SUMMARY_HEADERS,
} from "../config/balingSchemas.js";
import { baseStyles, styleHeaderRow, styleBodyRows, applyColumnWidths } from "./excelCommon.js";

const EXAMPLE_PROD_HEADER_ROW2 = [
  "",
  "Date",
  "Baling Machine \nNumber",
  "Bale Type",
  "Bale Number",
  "Bale Series\nM/Y",
  "Baler Operator",
  "Assistant Baler Operator",
  "Start Time",
  "Finish Time",
  "Total Time to Bale",
  "Passenger",
  "4 X 4",
  "Motorcylce",
  "Light Commercial",
  "Light Commercial T",
  "Light Commercial SW",
  "Heavy Commercial T",
  "Heavy Commercial SW",
  "Total Number of T",
  "Total Number of SW",
  "Agricultural T",
  "Agricultural SW",
  "Heavy Commercial T",
  "Heavy Commercial SW",
  "Total Number of Nylon T",
  "Total Number of Nylon SW",
  "Total Number of Tyres",
  "Bale Weight\nKG",
  "Bale Weight\nTONS",
  "Raw Text",
];

const EXAMPLE_PROD_HEADER_ROW3 = [
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "PCR",
  "",
  "",
  "",
  "RADIALS",
  "",
  "",
  "",
  "",
  "",
  "NYLONS",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
];

function sortByDateAndTime(records) {
  const baleSortKey = (row) => {
    const candidates = [row?.baleNumber, row?.baleTestCode];
    for (const c of candidates) {
      const s = String(c || "").trim();
      if (!s) continue;
      const m = s.match(/\bB\s*[- ]?(\d+)\b/i);
      if (m) return parseInt(m[1], 10);
    }
    return Number.MAX_SAFE_INTEGER;
  };

  return [...records].sort((a, b) => {
    const d = dateSortKey(a.chatDateParsed || a.date) - dateSortKey(b.chatDateParsed || b.date);
    if (d !== 0) return d;

    const bk = baleSortKey(a) - baleSortKey(b);
    if (bk !== 0) return bk;

    const aTime = a.startTime ? (a.startTime.h * 60 + a.startTime.m) : 0;
    const bTime = b.startTime ? (b.startTime.h * 60 + b.startTime.m) : 0;
    if (aTime !== bTime) return aTime - bTime;

    return String(a.machine || "").localeCompare(String(b.machine || ""));
  });
}

function num(v) {
  return v === null || v === undefined ? "" : v;
}

function formatTimeWithSeconds(t) {
  if (!t) return "";
  const hours = Number.isFinite(t.h) ? t.h : 0;
  const minutes = Number.isFinite(t.m) ? t.m : 0;
  const seconds = Number.isFinite(t.s) ? t.s : 0;
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
}

function toSeconds(timeObj) {
  if (!timeObj || !Number.isFinite(timeObj.h) || !Number.isFinite(timeObj.m)) return null;
  const s = Number.isFinite(timeObj.s) ? timeObj.s : 0;
  return (timeObj.h * 3600) + (timeObj.m * 60) + s;
}

function formatSecondsAsHHMMSS(totalSeconds) {
  if (!Number.isFinite(totalSeconds) || totalSeconds < 0) return "";
  const hrs = Math.floor(totalSeconds / 3600);
  const mins = Math.floor((totalSeconds % 3600) / 60);
  const secs = Math.floor(totalSeconds % 60);
  return `${String(hrs).padStart(2, "0")}:${String(mins).padStart(2, "0")}:${String(secs).padStart(2, "0")}`;
}

function durationSecondsFromRow(r) {
  const startSec = toSeconds(r?.startTime);
  const finishSec = toSeconds(r?.finishTime);
  if (Number.isFinite(startSec) && Number.isFinite(finishSec) && finishSec >= startSec) {
    return finishSec - startSec;
  }

  if (Number.isFinite(r?.durationMinutes) && r.durationMinutes >= 0) {
    return Math.round(r.durationMinutes * 60);
  }

  return null;
}

function formatDurationForRow(r) {
  const seconds = durationSecondsFromRow(r);
  return Number.isFinite(seconds) ? formatSecondsAsHHMMSS(seconds) : "";
}

function parseHHMMSSToSeconds(value) {
  const m = String(value || "").trim().match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
  if (!m) return null;
  const h = Number(m[1]);
  const min = Number(m[2]);
  const s = Number(m[3]);
  if (!Number.isFinite(h) || !Number.isFinite(min) || !Number.isFinite(s)) return null;
  return (h * 3600) + (min * 60) + s;
}

function applyDurationTrafficLight(ws, durationColumnIndex, fromRow = 2) {
  const fills = {
    green: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFC6EFCE" },
    },
    orange: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFE699" },
    },
    red: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF4CCCC" },
    },
  };

  for (let rowIdx = fromRow; rowIdx <= ws.rowCount; rowIdx += 1) {
    const cell = ws.getRow(rowIdx).getCell(durationColumnIndex);
    const seconds = parseHHMMSSToSeconds(cell.value);
    if (!Number.isFinite(seconds)) continue;
    if (seconds < 15 * 60) {
      cell.fill = fills.green;
    } else if (seconds < 20 * 60) {
      cell.fill = fills.orange;
    } else {
      cell.fill = fills.red;
    }
  }
}

function addProductionGroupedHeader(ws) {
  ws.addRow(new Array(EXAMPLE_PROD_HEADER_ROW2.length).fill(""));
  ws.addRow(EXAMPLE_PROD_HEADER_ROW2);
  ws.addRow(EXAMPLE_PROD_HEADER_ROW3);
  // Match template behavior: sublabels are centered across grouped blocks.
  ws.mergeCells(3, 12, 3, 15); // L3:O3 -> PCR
  ws.mergeCells(3, 16, 3, 21); // P3:U3 -> RADIALS
  ws.mergeCells(3, 22, 3, 27); // V3:AA3 -> NYLONS
}

function styleProductionHeader(ws, styles) {
  const thin = styles.thinBlack;
  const medium = styles.mediumBlack;
  const groupStarts = new Set([2, 12, 16, 22, 28, 31]); // B, L, P, V, AB, AE
  const groupEnds = new Set([11, 15, 21, 27, 30, 31]); // K, O, U, AA, AD, AE

  const fills = [
    { from: 2, to: 15, fgColor: { theme: 4, tint: 0.7999511703848384 } }, // B:O light blue
    { from: 16, to: 21, fgColor: { argb: "FFE97132" } }, // P:U RADIALS
    { from: 22, to: 27, fgColor: { argb: "FF0D8ABA" } }, // V:AA NYLONS
    { from: 28, to: 30, fgColor: { theme: 4, tint: 0.7999511703848384 } }, // AB:AD same as PCR
    { from: 31, to: 31, fgColor: { theme: 4, tint: 0.7999511703848384 } }, // AE same as PCR
  ];

  const fillForCol = (col) => fills.find((f) => col >= f.from && col <= f.to)?.fgColor || { theme: 4, tint: 0.7999511703848384 };

  ws.getRow(2).height = 48;
  ws.getRow(3).height = 17;

  for (let row = 2; row <= 3; row += 1) {
    for (let col = 2; col <= 31; col += 1) {
      const cell = ws.getRow(row).getCell(col);
      const isRadials = col >= 16 && col <= 21;
      const isNylons = col >= 22 && col <= 27;
      cell.font = {
        bold: true,
        size: 11,
        color: (isRadials || isNylons) ? { argb: "FF000000" } : { theme: 1 },
        name: "Aptos Narrow",
        family: 2,
        scheme: "minor",
      };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: row === 2 };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: fillForCol(col), bgColor: { indexed: 64 } };
      cell.border = {
        left: groupStarts.has(col) ? medium : thin,
        right: groupEnds.has(col) ? medium : thin,
        top: row === 2 ? medium : thin,
        bottom: row === 3 ? medium : thin,
      };
    }
  }
}

function applyNumericFormatting(ws, numericColumns) {
  for (const col of numericColumns) {
    ws.getColumn(col).numFmt = "0";
  }
}

function componentQtySum(r) {
  const values = [
    r.passengerQty,
    r.fourx4Qty,
    r.lcQty,
    r.motorcycleQty,
    r.srQty,
    r.agriQty,
    r.treadQty,
    r.sideWallQty,
  ].map((v) => Number(v));

  const present = values.filter((v) => Number.isFinite(v));
  if (present.length === 0) return null;
  return present.reduce((a, b) => a + b, 0);
}

function resolvedTotalQty(r) {
  const sum = componentQtySum(r);
  if (sum !== null) return sum;
  return r.totalQty ?? null;
}

function resolvedWeightTons(r) {
  const kg = Number(r?.weightKg);
  if (!Number.isFinite(kg)) return null;
  return Math.round((kg / 1000) * 1000) / 1000;
}

function dateToMonthYear(d) {
  if (!d || !Number.isInteger(d.month) || !Number.isInteger(d.year)) return "";
  return `${String(d.month).padStart(2, "0")}/${String(d.year).slice(-2)}`;
}

function isUsableDate(d) {
  if (!d) return false;
  if (!Number.isInteger(d.year) || !Number.isInteger(d.month) || !Number.isInteger(d.day)) return false;
  if (d.month < 1 || d.month > 12) return false;
  if (d.day < 1 || d.day > 31) return false;
  return true;
}

function monthKeyFromDate(d) {
  return `${d.year}-${String(d.month).padStart(2, "0")}`;
}

function groupRowsByMonth(records) {
  const byMonth = new Map();
  const invalid = [];
  for (const r of records) {
    const d = r.chatDateParsed || r.date || null;
    if (!isUsableDate(d)) {
      invalid.push(r);
      continue;
    }
    const key = monthKeyFromDate(d);
    if (!byMonth.has(key)) byMonth.set(key, []);
    byMonth.get(key).push(r);
  }
  return { byMonth, invalid };
}

function appendProductionRow(ws, r) {
  const effectiveRecordType =
    r.recordType ||
    (r.failureType ? "FAILED" : (r.baleTestCode ? "CR_CA" : (r.scrapType ? "SCRAP" : "STANDARD")));

  ws.addRow([
    "",
    dateToStr(r.chatDateParsed),
    r.machine || "",
    r.productionType || r.testType || effectiveRecordType,
    r.baleNumber || r.baleTestCode || "",
    dateToMonthYear(r.chatDateParsed || r.date),
    r.operator || "",
    r.assistant || "",
    formatTimeWithSeconds(r.startTime),
    formatTimeWithSeconds(r.finishTime),
    formatDurationForRow(r),
    num(r.passengerQty),
    num(r.fourx4Qty),
    num(r.motorcycleQty),
    num(r.lcQty),
    "",
    "",
    "",
    "",
    num(r.treadQty),
    num(r.sideWallQty),
    num(r.agriQty),
    "",
    "",
    "",
    "",
    "",
    num(resolvedTotalQty(r)),
    num(r.weightKg),
    num(resolvedWeightTons(r)),
    r.rawMessage || "",
  ]);
}

function sumNumeric(rows, getter) {
  return rows.reduce((acc, row) => {
    const v = Number(getter(row));
    return Number.isFinite(v) ? acc + v : acc;
  }, 0);
}

function appendProductionTotalsRow(ws, rows) {
  const durationValues = rows
    .map((r) => durationSecondsFromRow(r))
    .filter((v) => Number.isFinite(v) && v >= 0);
  const avgDurationSec = durationValues.length
    ? Math.round(durationValues.reduce((a, b) => a + b, 0) / durationValues.length)
    : null;

  const totalPassenger = sumNumeric(rows, (r) => r.passengerQty);
  const total4x4 = sumNumeric(rows, (r) => r.fourx4Qty);
  const totalMotorcycle = sumNumeric(rows, (r) => r.motorcycleQty);
  const totalLc = sumNumeric(rows, (r) => r.lcQty);
  const totalTread = sumNumeric(rows, (r) => r.treadQty);
  const totalSideWall = sumNumeric(rows, (r) => r.sideWallQty);
  const totalAgri = sumNumeric(rows, (r) => r.agriQty);
  const totalTyres = sumNumeric(rows, (r) => resolvedTotalQty(r));
  const totalWeightKg = sumNumeric(rows, (r) => r.weightKg);
  const totalWeightTons = Math.round((totalWeightKg / 1000) * 1000) / 1000;

  ws.addRow([
    "",
    "TOTALS",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    avgDurationSec !== null ? formatSecondsAsHHMMSS(avgDurationSec) : "",
    totalPassenger || "",
    total4x4 || "",
    totalMotorcycle || "",
    totalLc || "",
    "",
    "",
    "",
    "",
    totalTread || "",
    totalSideWall || "",
    totalAgri || "",
    "",
    "",
    "",
    "",
    "",
    totalTyres || "",
    totalWeightKg || "",
    totalWeightTons || "",
    "",
  ]);
}

export async function createBalingWorkbook(data = {}) {
  const {
    standardRecords = [],
    failedRecords = [],
    scrapRecords = [],
    crcaRecords = [],
    summaryRecords = [],
    validationLog = [],
  } = data;

  const { default: ExcelJS } = await import("exceljs");
  const wb = new ExcelJS.Workbook();
  const styles = baseStyles();

  const monthSheets = [];
  const logEntries = [...validationLog];
  const allProductionRows = [
    ...standardRecords,
    ...failedRecords,
    ...scrapRecords,
    ...crcaRecords,
  ];
  const { byMonth, invalid: invalidDatedProductionRows } = groupRowsByMonth(allProductionRows);
  const sortedMonthKeys = [...byMonth.keys()].sort();

  for (const key of sortedMonthKeys) {
    const ws = wb.addWorksheet(monthLabel(key));
    addProductionGroupedHeader(ws);
    const monthRows = sortByDateAndTime(byMonth.get(key));
    for (const r of monthRows) appendProductionRow(ws, r);
    appendProductionTotalsRow(ws, monthRows);
    monthSheets.push(ws);
  }

  let reviewWs = null;
  if (invalidDatedProductionRows.length > 0) {
    reviewWs = wb.addWorksheet("Review_Missing_Date");
    reviewWs.addRow([...EXAMPLE_PROD_HEADER_ROW2, "Review Reason"]);
    for (const r of invalidDatedProductionRows) {
      reviewWs.addRow([
        "",
        "",
        r.machine || "",
        r.productionType || "Production",
        r.baleNumber || "",
        "",
        r.operator || "",
        r.assistant || "",
        formatTimeWithSeconds(r.startTime),
        formatTimeWithSeconds(r.finishTime),
        formatDurationForRow(r),
        num(r.passengerQty),
        num(r.fourx4Qty),
        num(r.motorcycleQty),
        num(r.lcQty),
        "",
        "",
        "",
        "",
        num(r.treadQty),
        num(r.sideWallQty),
        num(r.agriQty),
        "",
        "",
        "",
        "",
        "",
        num(resolvedTotalQty(r)),
        num(r.weightKg),
        num(resolvedWeightTons(r)),
        r.rawMessage || "",
        "Missing or invalid parsed date - month routing skipped",
      ]);

      logEntries.push({
        severity: "WARNING",
        issueType: "MISSING_OR_INVALID_DATE",
        chatDateParsed: "",
        sourceMessageTimestamp: r.sourceTimestamp || "",
        machine: r.machine || "",
        baleNumberCode: r.baleNumber || r.baleTestCode || "",
        sheetTargetAttempted: "Monthly Production Sheet",
        problemDescription: "Row could not be routed to a month sheet due to missing/invalid parsed date",
        rawMessage: r.rawMessage || "",
        date: "",
      });
    }
  }

  const failedWs = wb.addWorksheet(BALING_SHEET_NAMES.failed);
  failedWs.addRow(BALING_FAILED_HEADERS);
  for (const r of sortByDateAndTime(failedRecords)) {
    failedWs.addRow([
      dateToStr(r.chatDateParsed),
      r.machine || "",
      r.baleNumber || "",
      r.failureType || "FAILED_BALE",
      r.failureReason || "",
      r.operator || "",
      r.assistant || "",
      formatTimeWithSeconds(r.startTime),
      formatTimeWithSeconds(r.finishTime),
      num(r.passengerQty),
      num(r.fourx4Qty),
      num(r.lcQty),
      num(r.motorcycleQty),
      num(r.srQty),
      num(r.agriQty),
      num(r.treadQty),
      num(r.sideWallQty),
      num(resolvedTotalQty(r)),
      num(r.weightKg),
      r.notesFlags || "",
      r.rawMessage || "",
    ]);
  }

  const scrapWs = wb.addWorksheet(BALING_SHEET_NAMES.scrap);
  scrapWs.addRow(BALING_SCRAP_HEADERS);
  for (const r of sortByDateAndTime(scrapRecords)) {
    scrapWs.addRow([
      dateToStr(r.chatDateParsed),
      r.machine || "",
      r.productionLabel || "",
      r.baleNumber || "",
      r.scrapType || "Scrap",
      num(r.scrapQty),
      num(r.weightKg),
      r.operator || "",
      r.assistant || "",
      formatTimeWithSeconds(r.startTime),
      formatTimeWithSeconds(r.finishTime),
      r.notesFlags || "",
      r.rawMessage || "",
    ]);
  }

  const crcaWs = wb.addWorksheet(BALING_SHEET_NAMES.crca);
  crcaWs.addRow(BALING_CRCA_HEADERS);
  for (const r of sortByDateAndTime(crcaRecords)) {
    crcaWs.addRow([
      dateToStr(r.chatDateParsed),
      r.machine || "",
      r.baleTestCode || r.baleNumber || "",
      r.testType || "Test",
      r.recordType || "CR_CA",
      r.operator || "",
      r.assistant || "",
      formatTimeWithSeconds(r.startTime),
      formatTimeWithSeconds(r.finishTime),
      formatDurationForRow(r),
      num(r.treadQty),
      num(r.sideWallQty),
      num(r.passengerQty),
      num(r.fourx4Qty),
      num(r.lcQty),
      num(r.motorcycleQty),
      num(r.srQty),
      num(r.agriQty),
      num(r.weightKg),
      r.notesFlags || "",
      r.rawMessage || "",
    ]);
  }

  const summaryWs = wb.addWorksheet(BALING_SHEET_NAMES.summaries);
  summaryWs.addRow(BALING_SUMMARY_HEADERS);
  for (const r of sortByDateAndTime(summaryRecords)) {
    summaryWs.addRow([
      dateToStr(r.chatDateParsed),
      r.summaryType || "BALING_DAILY_SUMMARY",
      r.machine || "",
      num(r.baleCount),
      num(r.weightKg),
      num(r.tons),
      num(r.passengerQty),
      num(r.fourx4Qty),
      num(r.lcQty),
      num(r.motorcycleQty),
      num(r.srQty),
      num(r.agriQty),
      num(r.treadQty),
      num(r.sideWallQty),
      num(r.totalTyres),
      r.machine1StartHour || "",
      r.machine1FinishHour || "",
      r.machine2StartHour || "",
      r.machine2FinishHour || "",
      r.notesFlags || "",
      r.rawMessage || "",
    ]);
  }

  const logWs = wb.addWorksheet(BALING_SHEET_NAMES.validation);
  logWs.addRow([
    "Severity",
    "Issue Type",
    "Chat Date Parsed",
    "Baling Machine Number",
    "Bale Number / Code",
    "Sheet Target Attempted",
    "Problem Description",
    "Raw Message",
  ]);

  for (const e of logEntries) {
    logWs.addRow([
      e.severity || "WARNING",
      e.issueType || "PARSER_WARNING",
      e.chatDateParsed || e.date || "",
      e.machine || "",
      e.baleNumberCode || "",
      e.sheetTargetAttempted || "",
      e.problemDescription || e.issue || "",
      e.rawMessage || "",
    ]);
  }

  const sheets = [...monthSheets, failedWs, scrapWs, crcaWs, summaryWs, logWs];
  if (reviewWs) sheets.push(reviewWs);
  for (const ws of sheets) {
    if (monthSheets.includes(ws)) {
      styleProductionHeader(ws, styles);
      styleBodyRows(ws, 4, ws.rowCount, styles.baseBorder);
      ws.views = [{ state: "frozen", xSplit: 1, ySplit: 3, topLeftCell: "B4" }];
      ws.autoFilter = { from: { row: 2, column: 2 }, to: { row: 2, column: 31 } };
    } else {
      styleHeaderRow(ws.getRow(1), styles, false);
      styleBodyRows(ws, 2, ws.rowCount, styles.baseBorder);
      ws.views = [{ state: "frozen", ySplit: 1 }];
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: ws.columnCount } };
    }
  }

  for (const ws of monthSheets) {
    applyColumnWidths(ws, [
      13,
      10.6640625,
      13,
      10.1640625,
      12.5,
      13,
      14.5,
      16.1640625,
      11.5,
      12.5,
      16.83203125,
      10.83203125,
      13,
      11.5,
      17.33203125,
      17.5,
      19.1640625,
      18.6640625,
      21,
      17.5,
      18.6640625,
      14.1640625,
      15.1640625,
      18.6640625,
      20.33203125,
      23.1640625,
      25,
      20.83203125,
      13,
      13,
      65,
    ]);
    applyNumericFormatting(ws, [12, 13, 14, 15, 20, 21, 22, 28, 29, 30]);
    ws.getColumn(30).numFmt = "0.000";
  }
  applyColumnWidths(failedWs, [14, 14, 10, 14, 22, 18, 18, 10, 10, 10, 10, 10, 10, 12, 10, 10, 12, 10, 10, 28, 80]);
  applyColumnWidths(scrapWs, [14, 14, 24, 10, 20, 10, 10, 18, 18, 10, 10, 28, 80]);
  applyColumnWidths(crcaWs, [14, 14, 18, 10, 10, 18, 18, 10, 10, 10, 10, 12, 10, 10, 10, 12, 10, 10, 10, 28, 80]);
  applyColumnWidths(summaryWs, [14, 24, 14, 10, 10, 10, 10, 10, 10, 12, 10, 10, 10, 12, 10, 16, 16, 16, 16, 28, 80]);
  applyColumnWidths(logWs, [10, 22, 14, 14, 18, 20, 44, 90]);
  if (reviewWs) applyColumnWidths(reviewWs, [13, 10.6640625, 13, 10.1640625, 12.5, 13, 14.5, 16.1640625, 11.5, 12.5, 16.83203125, 10.83203125, 13, 11.5, 17.33203125, 17.5, 19.1640625, 18.6640625, 21, 17.5, 18.6640625, 14.1640625, 15.1640625, 18.6640625, 20.33203125, 23.1640625, 25, 20.83203125, 13, 13, 65, 42]);

  // Numeric formatting for qty/weight/duration columns.
  applyNumericFormatting(failedWs, [10, 11, 12, 13, 14, 15, 16, 17, 18, 19]);
  applyNumericFormatting(scrapWs, [6, 7]);
  applyNumericFormatting(crcaWs, [11, 12, 13, 14, 15, 16, 17, 18, 19]);
  applyNumericFormatting(summaryWs, [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]);
  if (reviewWs) {
    applyNumericFormatting(reviewWs, [12, 13, 14, 15, 20, 21, 22, 28, 29, 30]);
    reviewWs.getColumn(30).numFmt = "0.000";
  }

  for (const ws of monthSheets) applyDurationTrafficLight(ws, 11, 4);
  applyDurationTrafficLight(crcaWs, 10);
  if (reviewWs) applyDurationTrafficLight(reviewWs, 11);

  return wb;
}
