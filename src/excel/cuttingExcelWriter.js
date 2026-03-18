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
        totalLC: null, totalHC: null, totalRadials: null,
        totalAgri: null, totalAgriTreads: null, totalUnknown: null,
      });
    }
    const cur = bucket.get(cm);
    for (const k of ["totalLC", "totalHC", "totalRadials", "totalAgri", "totalAgriTreads", "totalUnknown"]) {
      if (s[k] !== null && s[k] !== undefined) cur[k] = s[k];
    }
    if (s._unresolved) cur._unresolved = true;
  }

  const out = [];
  const sortedDates = [...byDate.values()].sort((a, b) => dateSortKey(a.date) - dateSortKey(b.date));
  for (const d of sortedDates) {
    for (const cm of [1, 2, 3]) {
      const existing = d.byCm.get(cm);
      out.push(existing || {
        date: d.date,
        cmNumber: `CM - ${cm}`,
        totalLC: null, totalHC: null, totalRadials: null,
        totalAgri: null, totalAgriTreads: null, totalUnknown: null,
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
    if (col >= 7 && col <= 10) return cuttingHeaderOrange;
    if (col === 11) return cuttingHeaderYellow;
    return cuttingHeaderGray;
  };

  for (const rowIndex of [1, 2]) {
    const row = ws.getRow(rowIndex);
    for (let c = 1; c <= 12; c += 1) {
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
    { start: 11, end: 11 },
    { start: 12, end: 12 },
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
  for (let c = 1; c <= 12; c += 1) {
    const cell = row.getCell(c);
    cell.font = { ...(cell.font || {}), bold: true, color: { argb: styles.textDark } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: cuttingHeaderGray } };
    cell.border = styles.baseBorder;
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  }

  for (let c = 1; c <= 12; c += 1) {
    const cell = row.getCell(c);
    const border = { ...(cell.border || {}) };
    border.top = styles.mediumBlack;
    border.bottom = styles.mediumBlack;
    if (c === 1) border.left = styles.mediumBlack;
    if (c === 12) border.right = styles.mediumBlack;
    cell.border = border;
  }
}

function styleDailySummaryHeaderRows(titleRow, fieldRow, styles) {
  const summaryBlueDark = "FF1F4E79";

  for (let c = 1; c <= 7; c += 1) {
    const cell = titleRow.getCell(c);
    cell.font = { bold: true, size: c === 1 ? 16 : 12, color: { argb: styles.textWhite } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: summaryBlueDark } };
    cell.alignment = { vertical: "middle", horizontal: c === 1 ? "left" : "center", wrapText: true };
    cell.border = styles.baseBorder;
  }
  for (let c = 1; c <= 7; c += 1) {
    const cell = fieldRow.getCell(c);
    cell.font = { bold: true, color: { argb: styles.textWhite } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: summaryBlueDark } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = styles.baseBorder;
  }

  for (let c = 1; c <= 7; c += 1) {
    const topCell = titleRow.getCell(c);
    const bottomCell = fieldRow.getCell(c);
    topCell.border = { ...(topCell.border || {}), top: styles.mediumBlack };
    bottomCell.border = { ...(bottomCell.border || {}), bottom: styles.mediumBlack };
    if (c === 1) {
      topCell.border = { ...(topCell.border || {}), left: styles.mediumBlack };
      bottomCell.border = { ...(bottomCell.border || {}), left: styles.mediumBlack };
    }
    if (c === 7) {
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
  wb.calcProperties.fullCalcOnLoad = true;
  wb.calcProperties.forceFullCalc = true;
  const styles = baseStyles();

  const byMonth = groupByMonth(records);
  const summaryByMonth = groupByMonth(summaryRecords);
  const allKeys = [...new Set([...Object.keys(byMonth), ...Object.keys(summaryByMonth)])].sort();

  for (const key of allKeys) {
    const sheetName = monthLabel(key);
    const monthRecords = byMonth[key] || [];
    const monthSummary = summaryByMonth[key] || [];

    const { rows, unresolvedIndices, inferredIndices } = cuttingSheetRows(monthRecords);
    let hasCuttingTotalsRow = false;

    if (rows.length >= 2) {
      rows.push(["TOTALS", "", "", "", "", "", "", "", "", "", "", ""]);
      hasCuttingTotalsRow = true;
    }

    const unresolvedSummaryRowIndices = [];
    if (monthSummary.length > 0) {
      const sorted = withSummaryPlaceholders(monthSummary);
      rows.push([]);
      rows.push(["Cutting Summary", "", "", "", "", "", ""]);
      rows.push(["Date", "CM Number", "LC", "HC", "Radials Total", "Agri", "Unknown"]);
      for (const s of sorted) {
        const nz = (v) => (v != null && v !== 0) ? v : "";
        const radialsTotal = s.totalRadials != null && s.totalRadials !== 0
          ? s.totalRadials
          : (s.totalLC || s.totalHC) ? (s.totalLC ?? 0) + (s.totalHC ?? 0) : "";
        rows.push([
          dateToStr(s.date),
          s.cmNumber,
          nz(s.totalLC),
          nz(s.totalHC),
          radialsTotal,
          nz(s.totalAgri),
          nz(s.totalUnknown),
        ]);
        if (s._unresolved) unresolvedSummaryRowIndices.push(rows.length);
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
      to: { row: 2, column: 12 },
    };
    applyColumnWidths(ws, [13, 16, 14, 26, 10, 10, 16, 16, 14, 14, 16, 65]);

    if (totalsRowIndex) {
      for (let col = 7; col <= 11; col += 1) {
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
      // Thick frame around TOTALS row
      styleThickFrame(ws, totalsRowIndex, totalsRowIndex, 1, 12, styles);
    }

    // Thick frame around the cutting data table (headers + data rows)
    if (monthRecords.length > 0) {
      styleThickFrame(ws, 1, dataEndRow, 1, 12, styles);
    }

    if (monthSummary.length > 0) {
      const summaryTitleRow = rows.findIndex((r) => Array.isArray(r) && r[0] === "Cutting Summary") + 1;
      if (summaryTitleRow > 0) {
        styleBodyRows(ws, 3, summaryTitleRow - 2, styles.baseBorder, 12);
        styleDailySummaryHeaderRows(ws.getRow(summaryTitleRow), ws.getRow(summaryTitleRow + 1), styles);
        styleBodyRows(ws, summaryTitleRow + 2, ws.rowCount, styles.baseBorder, 7);
        collapseRepeatedDates(ws, 3, summaryTitleRow - 2);
        styleFirstDateRows(ws, 3, summaryTitleRow - 2);
        collapseRepeatedDates(ws, summaryTitleRow + 2, ws.rowCount);
        styleFirstDateRows(ws, summaryTitleRow + 2, ws.rowCount);
        styleThickFrame(ws, summaryTitleRow, ws.rowCount, 1, 7, styles);
      } else {
        styleBodyRows(ws, 3, ws.rowCount, styles.baseBorder, 12);
        collapseRepeatedDates(ws, 3, ws.rowCount);
        styleFirstDateRows(ws, 3, ws.rowCount);
      }
    } else {
      styleBodyRows(ws, 3, ws.rowCount, styles.baseBorder, 12);
      collapseRepeatedDates(ws, 3, ws.rowCount);
      styleFirstDateRows(ws, 3, ws.rowCount);
    }

    // Highlight inferred-type rows in pastel light green on the monthly sheet
    const lightGreenPastel = "FFD5F5E3";
    for (const idx of inferredIndices) {
      const excelRow = dataStartRow + idx;
      const row = ws.getRow(excelRow);
      for (let c = 1; c <= 12; c += 1) {
        const cell = row.getCell(c);
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: lightGreenPastel } };
      }
    }

    // Highlight unresolved-type rows in pastel pink on the monthly sheet
    // (applied after green so unresolved takes priority over inferred)
    const lightPinkPastel = "FFFADBD8";
    for (const idx of unresolvedIndices) {
      const excelRow = dataStartRow + idx;
      const row = ws.getRow(excelRow);
      for (let c = 1; c <= 12; c += 1) {
        const cell = row.getCell(c);
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: lightPinkPastel } };
      }
    }

    // Highlight unresolved summary rows in pastel pink
    for (const rowIdx of unresolvedSummaryRowIndices) {
      const row = ws.getRow(rowIdx);
      for (let c = 1; c <= 7; c += 1) {
        const cell = row.getCell(c);
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: lightPinkPastel } };
      }
    }
  }

  if (validationLog.length > 0) {
    const logRows = [
      ["Body Date", "Time", "Message Type", "Cutter", "Issue", "Action Taken", "Raw Text"],
      ...validationLog.map((e) => [e.date, e.time, e.messageType, e.cutter, e.issue, e.action, e.rawText || ""]),
    ];
    const ws = wb.addWorksheet("Validation_Log");
    logRows.forEach((row) => ws.addRow(row));
    styleHeaderRow(ws.getRow(1), styles, false);
    styleBodyRows(ws, 2, ws.rowCount, styles.baseBorder, 7);
    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: 7 },
    };
    applyColumnWidths(ws, [12, 12, 14, 12, 44, 44, 60]);
  }

  await workbookToBrowserDownload(wb, filename);
}
