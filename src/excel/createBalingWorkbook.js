import { dateSortKey, dateToStr, monthLabel } from "../helpers.js";
import {
  BALING_SHEET_NAMES,
  BALING_PRODUCTION_HEADERS,
  BALING_FAILED_HEADERS,
  BALING_SCRAP_HEADERS,
  BALING_CRCA_HEADERS,
  BALING_SUMMARY_HEADERS,
} from "../config/balingSchemas.js";
import { baseStyles, styleHeaderRow, styleBodyRows, applyColumnWidths } from "./excelCommon.js";

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
  const headerLen = BALING_PRODUCTION_HEADERS.length;
  const top = new Array(headerLen).fill("");
  const sub = new Array(headerLen).fill("");

  for (let i = 0; i < headerLen; i += 1) {
    const col = i + 1;
    const label = BALING_PRODUCTION_HEADERS[i];
    if (col >= 11 && col <= 13) {
      sub[i] = label;
    } else {
      top[i] = label;
    }
  }

  top[10] = "Input Materials - No. of Tyres";
  ws.addRow(top);
  ws.addRow(sub);
  ws.mergeCells(1, 11, 1, 13);
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
    dateToStr(r.chatDateParsed),
    r.machine || "",
    r.baleNumber || r.baleTestCode || "",
    effectiveRecordType,
    r.productionType || r.testType || (effectiveRecordType === "STANDARD" ? "Production" : ""),
    r.operator || "",
    r.assistant || "",
    formatTimeWithSeconds(r.startTime),
    formatTimeWithSeconds(r.finishTime),
    formatDurationForRow(r),
    num(r.passengerQty),
    num(r.fourx4Qty),
    num(r.lcQty),
    num(r.motorcycleQty),
    num(r.srQty),
    num(r.agriQty),
    num(r.treadQty),
    num(r.sideWallQty),
    r.otherItemRaw || "",
    num(resolvedTotalQty(r)),
    num(r.weightKg),
    r.notesFlags || "",
    r.rawMessage || "",
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
    monthSheets.push(ws);
  }

  let reviewWs = null;
  if (invalidDatedProductionRows.length > 0) {
    reviewWs = wb.addWorksheet("Review_Missing_Date");
    reviewWs.addRow([...BALING_PRODUCTION_HEADERS, "Review Reason"]);
    for (const r of invalidDatedProductionRows) {
      reviewWs.addRow([
        "",
        r.machine || "",
        r.baleNumber || "",
        r.recordType || "STANDARD",
        r.productionType || "Production",
        r.operator || "",
        r.assistant || "",
        formatTimeWithSeconds(r.startTime),
        formatTimeWithSeconds(r.finishTime),
        formatDurationForRow(r),
        num(r.passengerQty),
        num(r.fourx4Qty),
        num(r.lcQty),
        num(r.motorcycleQty),
        num(r.srQty),
        num(r.agriQty),
        num(r.treadQty),
        num(r.sideWallQty),
        r.otherItemRaw || "",
        num(resolvedTotalQty(r)),
        num(r.weightKg),
        r.notesFlags || "",
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
      styleHeaderRow(ws.getRow(1), styles, false);
      styleHeaderRow(ws.getRow(2), styles, false);
      styleBodyRows(ws, 3, ws.rowCount, styles.baseBorder);
      ws.views = [{ state: "frozen", ySplit: 2 }];
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: ws.columnCount } };
    } else {
      styleHeaderRow(ws.getRow(1), styles, false);
      styleBodyRows(ws, 2, ws.rowCount, styles.baseBorder);
      ws.views = [{ state: "frozen", ySplit: 1 }];
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: ws.columnCount } };
    }
  }

  for (const ws of monthSheets) {
    applyColumnWidths(ws, [14, 14, 10, 12, 12, 18, 18, 10, 10, 10, 10, 10, 10, 12, 10, 10, 10, 12, 28, 10, 10, 28, 80]);
    applyNumericFormatting(ws, [11, 12, 13, 14, 15, 16, 17, 18, 20, 21]);
  }
  applyColumnWidths(failedWs, [14, 14, 10, 14, 22, 18, 18, 10, 10, 10, 10, 10, 10, 12, 10, 10, 12, 10, 10, 28, 80]);
  applyColumnWidths(scrapWs, [14, 14, 24, 10, 20, 10, 10, 18, 18, 10, 10, 28, 80]);
  applyColumnWidths(crcaWs, [14, 14, 18, 10, 10, 18, 18, 10, 10, 10, 10, 12, 10, 10, 10, 12, 10, 10, 10, 28, 80]);
  applyColumnWidths(summaryWs, [14, 24, 14, 10, 10, 10, 10, 10, 10, 12, 10, 10, 10, 12, 10, 16, 16, 16, 16, 28, 80]);
  applyColumnWidths(logWs, [10, 22, 14, 14, 18, 20, 44, 90]);
  if (reviewWs) applyColumnWidths(reviewWs, [14, 14, 10, 12, 12, 18, 18, 10, 10, 10, 10, 10, 10, 12, 10, 10, 10, 12, 28, 10, 10, 28, 80, 42]);

  // Numeric formatting for qty/weight/duration columns.
  applyNumericFormatting(failedWs, [10, 11, 12, 13, 14, 15, 16, 17, 18, 19]);
  applyNumericFormatting(scrapWs, [6, 7]);
  applyNumericFormatting(crcaWs, [11, 12, 13, 14, 15, 16, 17, 18, 19]);
  applyNumericFormatting(summaryWs, [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]);
  if (reviewWs) applyNumericFormatting(reviewWs, [11, 12, 13, 14, 15, 16, 17, 18, 20, 21]);

  for (const ws of monthSheets) applyDurationTrafficLight(ws, 10, 3);
  applyDurationTrafficLight(crcaWs, 10);
  if (reviewWs) applyDurationTrafficLight(reviewWs, 10);

  return wb;
}
