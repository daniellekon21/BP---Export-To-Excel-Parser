import { dateSortKey, dateToStr, monthLabel, groupByMonth } from "../helpers.js";
import { CUTTING_GROUP_MERGES, cuttingSheetRows } from "../cuttingSheetBuilder.js";
import { baseStyles, styleHeaderRow, styleBodyRows, applyColumnWidths, workbookToBrowserDownload } from "./excelCommon.js";

function normalizeCmNumber(cmNumber) {
  const m = String(cmNumber ?? "").match(/(\d+)/);
  if (!m) return null;
  const n = parseInt(m[1], 10);
  return Number.isNaN(n) ? null : n;
}

function withSummaryPlaceholders(summaryRows) {
  const byDate = new Map();

  for (const s of summaryRows) {
    if (!s?.date) continue;
    const key = `${s.date.year}-${String(s.date.month).padStart(2, "0")}-${String(s.date.day).padStart(2, "0")}`;
    if (!byDate.has(key)) byDate.set(key, { date: s.date, byCm: new Map() });
    const bucket = byDate.get(key).byCm;
    const cm = normalizeCmNumber(s.cmNumber);
    if (!cm) continue;

    if (!bucket.has(cm)) {
      bucket.set(cm, {
        date: s.date,
        cmNumber: `CM - ${cm}`,
        lc: null, hc: null, agri: null,
        tread_lc: null, tread_hc: null, tread_agri: null,
      });
    }
    const cur = bucket.get(cm);
    for (const k of ["lc", "hc", "agri", "tread_lc", "tread_hc", "tread_agri"]) {
      if (s[k] !== null && s[k] !== undefined) cur[k] = s[k];
    }
  }

  const out = [];
  const sortedDates = [...byDate.values()].sort((a, b) => dateSortKey(a.date) - dateSortKey(b.date));
  for (const d of sortedDates) {
    for (const cm of [1, 2, 3]) {
      const existing = d.byCm.get(cm);
      out.push(existing || {
        date: d.date,
        cmNumber: `CM - ${cm}`,
        lc: null, hc: null, agri: null,
        tread_lc: null, tread_hc: null, tread_agri: null,
      });
    }
  }
  return out;
}

function styleCuttingHeaderRows(ws, styles) {
  const cuttingHeaderGray = "FFC2C8D6";
  const cuttingHeaderOrange = "FFF08A2B";
  const cuttingHeaderYellow = "FFF7D447";

  const colorForColumn = (col) => {
    if (col >= 11 && col <= 12) return cuttingHeaderOrange;
    if (col >= 13 && col <= 15) return cuttingHeaderYellow;
    return cuttingHeaderGray;
  };

  for (const rowIndex of [1, 2]) {
    const row = ws.getRow(rowIndex);
    for (let c = 1; c <= 17; c += 1) {
      const cell = row.getCell(c);
      cell.font = { ...(cell.font || {}), bold: true, color: { argb: styles.textDark } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colorForColumn(c) } };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = {
        top: styles.thinBlack,
        left: styles.thinBlack,
        bottom: styles.thinBlack,
        right: styles.thinBlack,
      };
    }
  }

  const groups = [
    { start: 1, end: 6 },
    { start: 7, end: 10 },
    { start: 11, end: 12 },
    { start: 13, end: 15 },
    { start: 16, end: 16 },
    { start: 17, end: 17 },
  ];
  const topRow = ws.getRow(1);
  const bottomRow = ws.getRow(2);

  for (const g of groups) {
    for (let c = g.start; c <= g.end; c += 1) {
      topRow.getCell(c).border = { ...(topRow.getCell(c).border || {}), top: styles.mediumBlack };
      bottomRow.getCell(c).border = { ...(bottomRow.getCell(c).border || {}), bottom: styles.mediumBlack };
    }

    topRow.getCell(g.start).border = { ...(topRow.getCell(g.start).border || {}), left: styles.mediumBlack };
    bottomRow.getCell(g.start).border = { ...(bottomRow.getCell(g.start).border || {}), left: styles.mediumBlack };
    topRow.getCell(g.end).border = { ...(topRow.getCell(g.end).border || {}), right: styles.mediumBlack };
    bottomRow.getCell(g.end).border = { ...(bottomRow.getCell(g.end).border || {}), right: styles.mediumBlack };
  }
}

function styleTotalsRow(row, styles) {
  const cuttingHeaderGray = "FFC2C8D6";
  for (let c = 1; c <= 17; c += 1) {
    const cell = row.getCell(c);
    cell.font = { ...(cell.font || {}), bold: true, color: { argb: styles.textDark } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: cuttingHeaderGray } };
    cell.border = styles.baseBorder;
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  }

  for (let c = 1; c <= 17; c += 1) {
    const cell = row.getCell(c);
    const border = { ...(cell.border || {}) };
    border.top = styles.mediumBlack;
    border.bottom = styles.mediumBlack;
    if (c === 1) border.left = styles.mediumBlack;
    if (c === 17) border.right = styles.mediumBlack;
    cell.border = border;
  }
}

function styleDailySummaryHeaderRows(titleRow, fieldRow, styles) {
  const summaryGreenDark = "FF2F7D4A";
  const summaryGreenLight = "FFE6F4EA";

  titleRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: styles.textWhite } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: summaryGreenDark } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = styles.baseBorder;
  });
  fieldRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: styles.textDark } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: summaryGreenLight } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = styles.baseBorder;
  });

  for (let c = 1; c <= 8; c += 1) {
    const topCell = titleRow.getCell(c);
    const bottomCell = fieldRow.getCell(c);
    topCell.border = { ...(topCell.border || {}), top: styles.mediumBlack };
    bottomCell.border = { ...(bottomCell.border || {}), bottom: styles.mediumBlack };
    if (c === 1) {
      topCell.border = { ...(topCell.border || {}), left: styles.mediumBlack };
      bottomCell.border = { ...(bottomCell.border || {}), left: styles.mediumBlack };
    }
    if (c === 8) {
      topCell.border = { ...(topCell.border || {}), right: styles.mediumBlack };
      bottomCell.border = { ...(bottomCell.border || {}), right: styles.mediumBlack };
    }
  }
}

function styleThickFrame(ws, fromRow, toRow, fromCol, toCol, styles) {
  if (fromRow > toRow || fromCol > toCol) return;
  for (let c = fromCol; c <= toCol; c += 1) {
    const topCell = ws.getCell(fromRow, c);
    const bottomCell = ws.getCell(toRow, c);
    topCell.border = { ...(topCell.border || {}), top: styles.mediumBlack };
    bottomCell.border = { ...(bottomCell.border || {}), bottom: styles.mediumBlack };
  }
  for (let r = fromRow; r <= toRow; r += 1) {
    const leftCell = ws.getCell(r, fromCol);
    const rightCell = ws.getCell(r, toCol);
    leftCell.border = { ...(leftCell.border || {}), left: styles.mediumBlack };
    rightCell.border = { ...(rightCell.border || {}), right: styles.mediumBlack };
  }
}

function styleFirstDateRows(ws, fromRow, toRow) {
  let prevDate = "";
  for (let r = fromRow; r <= toRow; r += 1) {
    const cell = ws.getRow(r).getCell(1);
    const raw = cell.value;
    const dateText = String(raw && typeof raw === "object" && "text" in raw ? raw.text : (raw ?? "")).trim();
    if (!/^\d{2}\/\d{2}\/\d{4}$/.test(dateText)) continue;
    if (dateText !== prevDate) {
      cell.font = { ...(cell.font || {}), bold: true };
      prevDate = dateText;
    }
  }
}

function collapseRepeatedDates(ws, fromRow, toRow) {
  let prevDate = "";
  for (let r = fromRow; r <= toRow; r += 1) {
    const cell = ws.getRow(r).getCell(1);
    const raw = cell.value;
    const dateText = String(raw && typeof raw === "object" && "text" in raw ? raw.text : (raw ?? "")).trim();
    if (!/^\d{2}\/\d{2}\/\d{4}$/.test(dateText)) continue;
    if (dateText === prevDate) {
      cell.value = "";
    } else {
      prevDate = dateText;
    }
  }
}

export async function downloadCuttingWorkbook(records, filename, extras = {}) {
  const { summaryRecords = [], validationLog = [] } = extras;
  const { default: ExcelJS } = await import("exceljs");
  const wb = new ExcelJS.Workbook();
  const styles = baseStyles();

  const byMonth = groupByMonth(records);
  const summaryByMonth = groupByMonth(summaryRecords);
  const allKeys = [...new Set([...Object.keys(byMonth), ...Object.keys(summaryByMonth)])].sort();

  for (const key of allKeys) {
    const sheetName = monthLabel(key);
    const monthRecords = byMonth[key] || [];
    const monthSummary = summaryByMonth[key] || [];

    const rows = cuttingSheetRows(monthRecords);
    let hasCuttingTotalsRow = false;

    if (rows.length >= 2) {
      rows.push(["TOTALS", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
      hasCuttingTotalsRow = true;
    }

    if (monthSummary.length > 0) {
      const sorted = withSummaryPlaceholders(monthSummary);
      rows.push([]);
      rows.push(["Daily Summary", "", "LC Tyres", "HC Tyres", "Agri Tyres", "LC Treads", "HC Treads", "Agri Treads"]);
      rows.push(["Date", "CM Number", "LC", "HC", "Agri", "Tread LC", "Tread HC", "Tread Agri"]);
      for (const s of sorted) {
        rows.push([
          dateToStr(s.date),
          s.cmNumber,
          s.lc ?? "",
          s.hc ?? "",
          s.agri ?? "",
          s.tread_lc ?? "",
          s.tread_hc ?? "",
          s.tread_agri ?? "",
        ]);
      }
    }

    const ws = wb.addWorksheet(sheetName);
    rows.forEach((row) => ws.addRow(row));

    const dataStartRow = 3;
    const dataEndRow = 2 + monthRecords.length;
    const totalsRowIndex = hasCuttingTotalsRow ? dataEndRow + 1 : null;

    for (const merge of CUTTING_GROUP_MERGES) {
      ws.mergeCells(merge.s.r + 1, merge.s.c + 1, merge.e.r + 1, merge.e.c + 1);
    }
    styleCuttingHeaderRows(ws, styles);
    ws.views = [{ state: "frozen", ySplit: 2 }];
    ws.autoFilter = {
      from: { row: 2, column: 1 },
      to: { row: 2, column: 17 },
    };
    applyColumnWidths(ws, [13, 16, 14, 26, 10, 10, 10, 10, 10, 14, 15, 15, 14, 14, 18, 12, 65]);

    if (totalsRowIndex) {
      for (let col = 7; col <= 16; col += 1) {
        const totalCell = ws.getCell(totalsRowIndex, col);
        if (monthRecords.length > 0) {
          totalCell.value = {
            formula: `SUM(${totalCell.address.replace(/\d+$/, dataStartRow)}:${totalCell.address.replace(/\d+$/, dataEndRow)})`,
          };
        } else {
          totalCell.value = 0;
        }
      }
      styleTotalsRow(ws.getRow(totalsRowIndex), styles);
    }

    if (monthSummary.length > 0) {
      const summaryTitleRow = rows.findIndex((r) => Array.isArray(r) && r[0] === "Daily Summary") + 1;
      if (summaryTitleRow > 0) {
        styleDailySummaryHeaderRows(ws.getRow(summaryTitleRow), ws.getRow(summaryTitleRow + 1), styles);
        styleBodyRows(ws, summaryTitleRow + 2, ws.rowCount, styles.baseBorder);
        collapseRepeatedDates(ws, 3, summaryTitleRow - 2);
        styleFirstDateRows(ws, 3, summaryTitleRow - 2);
        collapseRepeatedDates(ws, summaryTitleRow + 2, ws.rowCount);
        styleFirstDateRows(ws, summaryTitleRow + 2, ws.rowCount);
        styleThickFrame(ws, summaryTitleRow, ws.rowCount, 1, 8, styles);
      } else {
        styleBodyRows(ws, 3, ws.rowCount, styles.baseBorder);
        collapseRepeatedDates(ws, 3, ws.rowCount);
        styleFirstDateRows(ws, 3, ws.rowCount);
      }
    } else {
      styleBodyRows(ws, 3, ws.rowCount, styles.baseBorder);
      collapseRepeatedDates(ws, 3, ws.rowCount);
      styleFirstDateRows(ws, 3, ws.rowCount);
    }
  }

  if (validationLog.length > 0) {
    const logRows = [
      ["Body Date", "Time", "Message Type", "Cutter", "Issue", "Action Taken"],
      ...validationLog.map((e) => [e.date, e.time, e.messageType, e.cutter, e.issue, e.action]),
    ];
    const ws = wb.addWorksheet("Validation_Log");
    logRows.forEach((row) => ws.addRow(row));
    styleHeaderRow(ws.getRow(1), styles, false);
    styleBodyRows(ws, 2, ws.rowCount, styles.baseBorder);
    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: 6 },
    };
    applyColumnWidths(ws, [12, 12, 14, 12, 44, 44]);
  }

  await workbookToBrowserDownload(wb, filename);
}
