import { useState, useCallback, useRef, useMemo } from "react";

// ─── Shared Modules ────────────────────────────────────────────────────────────

import {
  parseTime, formatTime, dateToStr, dateSortKey,
  splitWhatsAppMessages, groupByMonth, monthLabel, MONTH_NAMES,
} from "./helpers.js";

import {
  parseCuttingMessages, parseCuttingMessagesNew,
} from "./cuttingParser.js";

import {
  CUTTING_GROUP_MERGES, cuttingSheetRows,
} from "./cuttingSheetBuilder.js";

// ─── Date Filter Helpers ──────────────────────────────────────────────────────

function getQuarter(month) {
  return Math.ceil(month / 3);
}

function quarterLabel(year, q) {
  return `Q${q} ${year}`;
}

function filterRecords(records, mode, filterYear, filterPeriod) {
  if (mode === "all") return records;
  return records.filter(r => {
    if (!r.date) return false;
    const { year, month } = r.date;
    if (mode === "year") {
      return filterYear ? year === parseInt(filterYear, 10) : true;
    }
    if (mode === "month") {
      const yearMatch = filterYear ? year === parseInt(filterYear, 10) : true;
      const periodMatch = filterPeriod ? month === parseInt(filterPeriod, 10) : true;
      return yearMatch && periodMatch;
    }
    if (mode === "quarter") {
      const yearMatch = filterYear ? year === parseInt(filterYear, 10) : true;
      const periodMatch = filterPeriod ? getQuarter(month) === parseInt(filterPeriod, 10) : true;
      return yearMatch && periodMatch;
    }
    return true;
  });
}

function parseLogDate(dateStr) {
  if (!dateStr) return null;
  const m = String(dateStr).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return {
    day: parseInt(m[1], 10),
    month: parseInt(m[2], 10),
    year: parseInt(m[3], 10),
  };
}

function filterValidationLog(logEntries, mode, filterYear, filterPeriod) {
  if (mode === "all") return logEntries;
  return logEntries.filter((entry) => {
    const d = parseLogDate(entry.date);
    if (!d) return false;
    const { year, month } = d;
    if (mode === "year") {
      return filterYear ? year === parseInt(filterYear, 10) : true;
    }
    if (mode === "month") {
      const yearMatch = filterYear ? year === parseInt(filterYear, 10) : true;
      const periodMatch = filterPeriod ? month === parseInt(filterPeriod, 10) : true;
      return yearMatch && periodMatch;
    }
    if (mode === "quarter") {
      const yearMatch = filterYear ? year === parseInt(filterYear, 10) : true;
      const periodMatch = filterPeriod ? getQuarter(month) === parseInt(filterPeriod, 10) : true;
      return yearMatch && periodMatch;
    }
    return true;
  });
}

// ─── Baling Parser ─────────────────────────────────────────────────────────────

function parseBalingMessages(text) {
  const messages = splitWhatsAppMessages(text);
  const records = [];

  for (const msg of messages) {
    const body = msg.body;

    if (body.includes("Daily summary") || body.includes("daily summary")) continue;
    if (body.includes("<Media omitted>")) continue;
    if (body.length < 30) continue;

    const weightMatch = body.match(/Weight\s*[:\-]\s*(\d[\d\s]*)\s*kg/i);
    if (!weightMatch) continue;

    const weight = parseInt(weightMatch[1].replace(/\s/g, ""), 10);

    let machine = "";
    const machineMatch = body.match(/Machine\s*(\d+|One|Two|1|2)/i);
    if (machineMatch) {
      const mv = machineMatch[1].toLowerCase();
      machine = (mv === "one" || mv === "1") ? "BM - 1" : "BM - 2";
    }

    // Use the WhatsApp system timestamp date as authoritative
    const date = msg.tsDate || null;

    let baleNum = "";
    let baleType = "";
    let baleSeries = "";

    const caMatch = body.match(/\b(CA|PCR|HC|LC|SW)\s*O?(\d{1,4})\s*[-–]\s*(\d{1,2}[\/\-]\d{2,4})/i);
    if (caMatch) {
      baleType = caMatch[1].toUpperCase();
      baleNum = caMatch[2].replace(/^0+/, "") || "0";
      baleSeries = caMatch[3].replace("-", "/");
      if (baleSeries.length > 5) {
        const parts = baleSeries.split("/");
        if (parts[1] && parts[1].length === 4) baleSeries = parts[0] + "/" + parts[1].slice(2);
      }
    } else {
      const bMatch = body.match(/\b[Bb](\d{1,4})\s*[-–]\s*(Production|Test|production|test)/i);
      if (bMatch) {
        baleNum = bMatch[1];
        baleType = bMatch[2].charAt(0).toUpperCase() + bMatch[2].slice(1).toLowerCase();
      }
    }

    let operator = "";
    let assistant = "";
    const opMatch = body.match(/Operator\s*[:\-]\s*(.+)/i);
    if (opMatch) operator = opMatch[1].trim().replace(/\s+/g, " ");
    const astMatch = body.match(/Assistant\s*[:\-]\s*(.+)/i);
    if (astMatch) assistant = astMatch[1].trim().replace(/\s+/g, " ");

    let startTime = null;
    let finishTime = null;
    const startMatch = body.match(/START\s*TIME\s*[:\-]\s*(\d{1,2}[.:]\d{2})/i);
    if (startMatch) startTime = parseTime(startMatch[1]);
    const finishMatch = body.match(/FINISH\s*TIME\s*[:\-]\s*(\d{1,2}[.:]\d{2})/i);
    if (finishMatch) finishTime = parseTime(finishMatch[1]);

    if (!startTime) {
      const st2 = body.match(/Start\s+time\s*[-–]\s*(\d{1,2}[.:]\d{2})/i);
      if (st2) startTime = parseTime(st2[1]);
    }
    if (!finishTime) {
      const ft2 = body.match(/Finish\s+time\s*[-–]\s*(\d{1,2}[.:]\d{2})/i);
      if (ft2) finishTime = parseTime(ft2[1]);
    }

    let passenger = 0, fourx4 = 0, motorcycle = 0, lightComm = 0;
    let agriT = 0, agriSW = 0, hcT = 0, hcSW = 0;
    let lcT = 0, lcSW = 0;

    const tMatch = body.match(/\bT\s*[-–]\s*(\d+)/i);
    const swMatch = body.match(/\bSW\s*[-–]\s*(\d+)/i);
    if (tMatch) agriT = parseInt(tMatch[1], 10);
    if (swMatch) agriSW = parseInt(swMatch[1], 10);

    const passMatch = body.match(/[Pp]assenger\s*(?:Qty\s*)?[-–:]\s*(\d+)/i);
    if (passMatch) passenger = parseInt(passMatch[1], 10);
    const f4Match = body.match(/4\s*[xX×]\s*4\s*(?:Qty\s*)?[-–:]\s*(\d+)/i);
    if (f4Match) fourx4 = parseInt(f4Match[1], 10);
    const mcMatch = body.match(/[Mm]otorcycle\s*(?:Qty\s*)?[-–:]\s*(\d+)/i);
    if (mcMatch) motorcycle = parseInt(mcMatch[1], 10);

    const itemLine = body.match(/Item\s*[-–:]\s*(.+)/i);
    if (itemLine) {
      const itemStr = itemLine[1];
      const ip = itemStr.match(/[Pp]assenger\s*(?:Qty\s*)?[-–:]\s*(\d+)/);
      if (ip) passenger = parseInt(ip[1], 10);
      const i4 = itemStr.match(/4\s*[xX×]\s*4\s*(?:Qty\s*)?[-–:]\s*(\d+)/);
      if (i4) fourx4 = parseInt(i4[1], 10);
    }

    if (!baleType && (agriT > 0 || agriSW > 0)) baleType = "CA";
    if (!baleType && (passenger > 0 || fourx4 > 0)) baleType = "PCR";

    records.push({
      date, machine, baleType, baleNum, baleSeries,
      operator, assistant,
      startTime, finishTime,
      passenger, fourx4, motorcycle, lightComm,
      lcT, lcSW, hcT, hcSW,
      agriT, agriSW,
      weight
    });
  }

  return records;
}

// ─── Baling Sheet Builder ──────────────────────────────────────────────────────

function balingSheetRows(records) {
  const headers = [
    "Date", "Baling Machine Number", "Bale Type", "Bale Number", "Bale Series M/Y",
    "Baler Operator", "Assistant Baler Operator", "Start Time", "Finish Time",
    "Passenger", "4 X 4", "Motorcycle", "Light Commercial",
    "LC T (Radials)", "LC SW (Radials)", "HC T (Radials)", "HC SW (Radials)",
    "Agricultural T (Nylons)", "Agricultural SW (Nylons)",
    "HC T (Nylons)", "HC SW (Nylons)",
    "Total Number of Tyres", "Bale Weight KG", "Bale Weight TONS"
  ];

  const sorted = [...records].sort((a, b) => dateSortKey(a.date) - dateSortKey(b.date));

  const rows = [headers];
  for (const r of sorted) {
    const totalTyres = r.passenger + r.fourx4 + r.motorcycle + r.lightComm +
      r.lcT + r.lcSW + r.hcT + r.hcSW + r.agriT + r.agriSW;
    rows.push([
      r.date ? dateToStr(r.date) : "",
      r.machine,
      r.baleType,
      r.baleNum,
      r.baleSeries,
      r.operator,
      r.assistant,
      r.startTime ? formatTime(r.startTime) : "",
      r.finishTime ? formatTime(r.finishTime) : "",
      r.passenger || "",
      r.fourx4 || "",
      r.motorcycle || "",
      r.lightComm || "",
      r.lcT || "",
      r.lcSW || "",
      r.hcT || "",
      r.hcSW || "",
      r.agriT || "",
      r.agriSW || "",
      0, 0,
      totalTyres || "",
      r.weight,
      parseFloat((r.weight / 1000).toFixed(3))
    ]);
  }
  return rows;
}

// ─── Download Helper ───────────────────────────────────────────────────────────

async function downloadXLSX(records, chatType, filename, extras = {}) {
  const { summaryRecords = [], validationLog = [] } = extras;
  const { default: ExcelJS } = await import("exceljs");
  const wb = new ExcelJS.Workbook();

  const brandBlue = "FF1B2E5C";
  const brandBlueLight = "FFEDF1F7";
  const textWhite = "FFFFFFFF";
  const textDark = "FF0F172A";
  const cuttingHeaderGray = "FFC2C8D6";
  const cuttingHeaderOrange = "FFF08A2B";
  const cuttingHeaderYellow = "FFF7D447";
  const summaryGreenDark = "FF2F7D4A";
  const summaryGreenLight = "FFE6F4EA";
  const thinBlack = { style: "thin", color: { argb: "FF000000" } };
  const mediumBlack = { style: "medium", color: { argb: "FF000000" } };

  const baseBorder = {
    top: { style: "thin", color: { argb: "FFD1D5DB" } },
    left: { style: "thin", color: { argb: "FFD1D5DB" } },
    bottom: { style: "thin", color: { argb: "FFD1D5DB" } },
    right: { style: "thin", color: { argb: "FFD1D5DB" } },
  };

  const styleHeaderRow = (row, isLight = false) => {
    row.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: isLight ? textDark : textWhite } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isLight ? brandBlueLight : brandBlue } };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = baseBorder;
    });
  };

  const styleCuttingHeaderRows = (ws) => {
    const colorForColumn = (col) => {
      if (col >= 11 && col <= 12) return cuttingHeaderOrange; // RADIALS
      if (col >= 13 && col <= 15) return cuttingHeaderYellow; // NYLONS
      return cuttingHeaderGray; // Date..PCR (+ Tread Cuts column)
    };

    for (const rowIndex of [1, 2]) {
      const row = ws.getRow(rowIndex);
      for (let c = 1; c <= 16; c += 1) {
        const cell = row.getCell(c);
        cell.font = { ...(cell.font || {}), bold: true, color: { argb: textDark } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colorForColumn(c) } };
        cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
        cell.border = {
          top: thinBlack,
          left: thinBlack,
          bottom: thinBlack,
          right: thinBlack,
        };
      }
    }

    // Thick outer frame per header category block:
    // metadata | PCR | RADIALS | NYLONS | Tread Cuts
    const groups = [
      { start: 1, end: 6 },
      { start: 7, end: 10 },
      { start: 11, end: 12 },
      { start: 13, end: 15 },
      { start: 16, end: 16 },
    ];
    const topRow = ws.getRow(1);
    const bottomRow = ws.getRow(2);

    for (const g of groups) {
      for (let c = g.start; c <= g.end; c += 1) {
        topRow.getCell(c).border = { ...(topRow.getCell(c).border || {}), top: mediumBlack };
        bottomRow.getCell(c).border = { ...(bottomRow.getCell(c).border || {}), bottom: mediumBlack };
      }

      topRow.getCell(g.start).border = { ...(topRow.getCell(g.start).border || {}), left: mediumBlack };
      bottomRow.getCell(g.start).border = { ...(bottomRow.getCell(g.start).border || {}), left: mediumBlack };
      topRow.getCell(g.end).border = { ...(topRow.getCell(g.end).border || {}), right: mediumBlack };
      bottomRow.getCell(g.end).border = { ...(bottomRow.getCell(g.end).border || {}), right: mediumBlack };
    }
  };

  const styleBodyRows = (ws, fromRow, toRow) => {
    for (let r = fromRow; r <= toRow; r += 1) {
      const row = ws.getRow(r);
      row.eachCell((cell) => {
        cell.alignment = { vertical: "middle", horizontal: "left" };
        cell.border = baseBorder;
      });
    }
  };

  const styleTotalsRow = (row) => {
    for (let c = 1; c <= 16; c += 1) {
      const cell = row.getCell(c);
      cell.font = { ...(cell.font || {}), bold: true, color: { argb: textDark } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: cuttingHeaderGray } };
      cell.border = baseBorder;
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    }

    // Thick frame around TOTALS row (outer border only)
    for (let c = 1; c <= 16; c += 1) {
      const cell = row.getCell(c);
      const border = { ...(cell.border || {}) };
      border.top = mediumBlack;
      border.bottom = mediumBlack;
      if (c === 1) border.left = mediumBlack;
      if (c === 16) border.right = mediumBlack;
      cell.border = border;
    }
  };

  const styleDailySummaryHeaderRows = (titleRow, fieldRow) => {
    titleRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: textWhite } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: summaryGreenDark } };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = baseBorder;
    });
    fieldRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: textDark } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: summaryGreenLight } };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = baseBorder;
    });

    // Thick frame around the 2-row Daily Summary header block (A:H)
    for (let c = 1; c <= 8; c += 1) {
      const topCell = titleRow.getCell(c);
      const bottomCell = fieldRow.getCell(c);
      topCell.border = { ...(topCell.border || {}), top: mediumBlack };
      bottomCell.border = { ...(bottomCell.border || {}), bottom: mediumBlack };
      if (c === 1) {
        topCell.border = { ...(topCell.border || {}), left: mediumBlack };
        bottomCell.border = { ...(bottomCell.border || {}), left: mediumBlack };
      }
      if (c === 8) {
        topCell.border = { ...(topCell.border || {}), right: mediumBlack };
        bottomCell.border = { ...(bottomCell.border || {}), right: mediumBlack };
      }
    }
  };

  const styleThickFrame = (ws, fromRow, toRow, fromCol, toCol) => {
    if (fromRow > toRow || fromCol > toCol) return;
    for (let c = fromCol; c <= toCol; c += 1) {
      const topCell = ws.getCell(fromRow, c);
      const bottomCell = ws.getCell(toRow, c);
      topCell.border = { ...(topCell.border || {}), top: mediumBlack };
      bottomCell.border = { ...(bottomCell.border || {}), bottom: mediumBlack };
    }
    for (let r = fromRow; r <= toRow; r += 1) {
      const leftCell = ws.getCell(r, fromCol);
      const rightCell = ws.getCell(r, toCol);
      leftCell.border = { ...(leftCell.border || {}), left: mediumBlack };
      rightCell.border = { ...(rightCell.border || {}), right: mediumBlack };
    }
  };

  const applyColumnWidths = (ws, widths) => {
    ws.columns = widths.map((w) => ({ width: w }));
  };

  const styleFirstDateRows = (ws, fromRow, toRow) => {
    let prevDate = "";
    for (let r = fromRow; r <= toRow; r += 1) {
      const cell = ws.getRow(r).getCell(1);
      const raw = cell.value;
      const dateText = String(
        raw && typeof raw === "object" && "text" in raw ? raw.text : (raw ?? "")
      ).trim();
      if (!/^\d{2}\/\d{2}\/\d{4}$/.test(dateText)) continue;
      if (dateText !== prevDate) {
        cell.font = { ...(cell.font || {}), bold: true };
        prevDate = dateText;
      }
    }
  };

  const collapseRepeatedDates = (ws, fromRow, toRow) => {
    let prevDate = "";
    for (let r = fromRow; r <= toRow; r += 1) {
      const cell = ws.getRow(r).getCell(1);
      const raw = cell.value;
      const dateText = String(
        raw && typeof raw === "object" && "text" in raw ? raw.text : (raw ?? "")
      ).trim();
      if (!/^\d{2}\/\d{2}\/\d{4}$/.test(dateText)) continue;
      if (dateText === prevDate) {
        cell.value = "";
      } else {
        prevDate = dateText;
      }
    }
  };

  const normalizeCmNumber = (cmNumber) => {
    const m = String(cmNumber ?? "").match(/(\d+)/);
    if (!m) return null;
    const n = parseInt(m[1], 10);
    return Number.isNaN(n) ? null : n;
  };

  const withSummaryPlaceholders = (summaryRows) => {
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
      // Merge non-null values so duplicate rows for same date/cm don't lose data.
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
  };

  const byMonth        = groupByMonth(records);
  const summaryByMonth = groupByMonth(summaryRecords);
  const allKeys = [...new Set([...Object.keys(byMonth), ...Object.keys(summaryByMonth)])].sort();

  for (const key of allKeys) {
    const sheetName    = monthLabel(key);
    const monthRecords = byMonth[key] || [];
    const monthSummary = summaryByMonth[key] || [];

    let rows;
    let hasCuttingTotalsRow = false;
    if (chatType === "baling") {
      rows = balingSheetRows(monthRecords);
    } else {
      rows = cuttingSheetRows(monthRecords);
      if (rows.length >= 2) {
        rows.push(["TOTALS", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
        hasCuttingTotalsRow = true;
      }

      // Append daily summary section at the bottom of each month sheet
      if (monthSummary.length > 0) {
        const sorted = withSummaryPlaceholders(monthSummary);
        rows.push([]); // blank separator
        rows.push(["Daily Summary", "", "LC Tyres", "HC Tyres", "Agri Tyres", "LC Treads", "HC Treads", "Agri Treads"]);
        rows.push(["Date", "CM Number",  "LC",       "HC",       "Agri",       "Tread LC",  "Tread HC",  "Tread Agri"]);
        for (const s of sorted) {
          rows.push([
            dateToStr(s.date),
            s.cmNumber,
            s.lc       ?? "",
            s.hc       ?? "",
            s.agri     ?? "",
            s.tread_lc ?? "",
            s.tread_hc ?? "",
            s.tread_agri ?? "",
          ]);
        }
      }
    }

    const ws = wb.addWorksheet(sheetName);
    rows.forEach((row) => ws.addRow(row));

    if (chatType === "cutting") {
      const dataStartRow = 3;
      const dataEndRow = 2 + monthRecords.length;
      const totalsRowIndex = hasCuttingTotalsRow ? dataEndRow + 1 : null;

      for (const merge of CUTTING_GROUP_MERGES) {
        ws.mergeCells(
          merge.s.r + 1,
          merge.s.c + 1,
          merge.e.r + 1,
          merge.e.c + 1
        );
      }
      styleCuttingHeaderRows(ws);
      ws.views = [{ state: "frozen", ySplit: 2 }];
      ws.autoFilter = {
        from: { row: 2, column: 1 },
        to: { row: 2, column: 16 },
      };
      applyColumnWidths(ws, [13, 16, 14, 26, 10, 10, 10, 10, 10, 14, 15, 15, 14, 14, 18, 12]);

      if (totalsRowIndex) {
        // Sum numeric production columns G:P on the cutting data block.
        for (let col = 7; col <= 16; col += 1) {
          const totalCell = ws.getCell(totalsRowIndex, col);
          if (monthRecords.length > 0) {
            totalCell.value = { formula: `SUM(${totalCell.address.replace(/\d+$/, dataStartRow)}:${totalCell.address.replace(/\d+$/, dataEndRow)})` };
          } else {
            totalCell.value = 0;
          }
        }
        styleTotalsRow(ws.getRow(totalsRowIndex));
      }

      if (monthSummary.length > 0) {
        const summaryTitleRow = rows.findIndex((r) => Array.isArray(r) && r[0] === "Daily Summary") + 1;
        if (summaryTitleRow > 0) {
          styleDailySummaryHeaderRows(ws.getRow(summaryTitleRow), ws.getRow(summaryTitleRow + 1));
          styleBodyRows(ws, summaryTitleRow + 2, ws.rowCount);
          collapseRepeatedDates(ws, 3, summaryTitleRow - 2);
          styleFirstDateRows(ws, 3, summaryTitleRow - 2);
          collapseRepeatedDates(ws, summaryTitleRow + 2, ws.rowCount);
          styleFirstDateRows(ws, summaryTitleRow + 2, ws.rowCount);
          styleThickFrame(ws, summaryTitleRow, ws.rowCount, 1, 8);
        } else {
          styleBodyRows(ws, 3, ws.rowCount);
          collapseRepeatedDates(ws, 3, ws.rowCount);
          styleFirstDateRows(ws, 3, ws.rowCount);
        }
      } else {
        styleBodyRows(ws, 3, ws.rowCount);
        collapseRepeatedDates(ws, 3, ws.rowCount);
        styleFirstDateRows(ws, 3, ws.rowCount);
      }
    } else {
      styleHeaderRow(ws.getRow(1), false);
      ws.views = [{ state: "frozen", ySplit: 1 }];
      ws.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: 24 },
      };
      applyColumnWidths(ws, [13, 16, 10, 10, 12, 20, 22, 10, 10, 10, 10, 10, 14, 14, 14, 14, 14, 16, 18, 14, 14, 18, 14, 14]);
      styleBodyRows(ws, 2, ws.rowCount);
      collapseRepeatedDates(ws, 2, ws.rowCount);
      styleFirstDateRows(ws, 2, ws.rowCount);
    }
  }

  // Validation_Log sheet
  if (validationLog.length > 0) {
    const logRows = [
      ["Body Date", "Time", "Message Type", "Cutter", "Issue", "Action Taken"],
      ...validationLog.map(e => [e.date, e.time, e.messageType, e.cutter, e.issue, e.action]),
    ];
    const logWs = wb.addWorksheet("Validation_Log");
    logRows.forEach((row) => logWs.addRow(row));
    styleHeaderRow(logWs.getRow(1), false);
    styleBodyRows(logWs, 2, logWs.rowCount);
    logWs.views = [{ state: "frozen", ySplit: 1 }];
    logWs.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: 6 },
    };
    applyColumnWidths(logWs, [12, 12, 14, 12, 44, 44]);
  }

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob(
    [buffer],
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
  );
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

// ─── Main App ──────────────────────────────────────────────────────────────────

export default function App() {
  const [chatType, setChatType] = useState(null);
  const [cuttingMode, setCuttingMode] = useState("old");
  const [chatText, setChatText] = useState("");
  const [records, setRecords] = useState([]);
  const [summaryRecords, setSummaryRecords] = useState([]);
  const [validationLog, setValidationLog] = useState([]);
  const [parsed, setParsed] = useState(false);
  const [fileName, setFileName] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [filterMode, setFilterMode] = useState("all");
  const [filterYear, setFilterYear] = useState("");
  const [filterPeriod, setFilterPeriod] = useState("");
  const [placeholderPolicy, setPlaceholderPolicy] = useState("keep");
  const [exportFileName, setExportFileName] = useState("");
  const fileRef = useRef(null);

  // Compute available filter options from parsed records
  const filterOptions = useMemo(() => {
    if (!records.length) return { months: [], quarters: [], years: [] };
    const monthSet = new Set();
    const quarterSet = new Set();
    const yearSet = new Set();
    for (const r of records) {
      if (!r.date) continue;
      const { year, month } = r.date;
      monthSet.add(`${year}-${String(month).padStart(2, "0")}`);
      quarterSet.add(`${year}-Q${getQuarter(month)}`);
      yearSet.add(String(year));
    }
    return {
      months: [...monthSet].sort(),
      quarters: [...quarterSet].sort(),
      years: [...yearSet].sort(),
    };
  }, [records]);

  // Filtered records for display and download
  const filteredRecords = useMemo(
    () => filterRecords(records, filterMode, filterYear, filterPeriod),
    [records, filterMode, filterYear, filterPeriod]
  );
  const filteredSummaryRecords = useMemo(
    () => filterRecords(summaryRecords, filterMode, filterYear, filterPeriod),
    [summaryRecords, filterMode, filterYear, filterPeriod]
  );
  const filteredValidationLog = useMemo(
    () => filterValidationLog(validationLog, filterMode, filterYear, filterPeriod),
    [validationLog, filterMode, filterYear, filterPeriod]
  );
  const visibleRecords = useMemo(() => {
    if (chatType !== "cutting" || placeholderPolicy === "keep") return filteredRecords;
    return filteredRecords.filter((r) => !r._syntheticPlaceholder);
  }, [filteredRecords, chatType, placeholderPolicy]);

  const allMonthKeys = useMemo(() => {
    const keys = new Set();
    for (const r of records) {
      if (!r?.date) continue;
      keys.add(`${r.date.year}-${String(r.date.month).padStart(2, "0")}`);
    }
    for (const s of summaryRecords) {
      if (!s?.date) continue;
      keys.add(`${s.date.year}-${String(s.date.month).padStart(2, "0")}`);
    }
    return [...keys].sort();
  }, [records, summaryRecords]);

  const exportMonthKeys = useMemo(() => {
    const keys = new Set();
    for (const r of visibleRecords) {
      if (!r?.date) continue;
      keys.add(`${r.date.year}-${String(r.date.month).padStart(2, "0")}`);
    }
    for (const s of filteredSummaryRecords) {
      if (!s?.date) continue;
      keys.add(`${s.date.year}-${String(s.date.month).padStart(2, "0")}`);
    }
    return [...keys].sort();
  }, [visibleRecords, filteredSummaryRecords]);

  const excludedMonthKeys = useMemo(
    () => allMonthKeys.filter((k) => !exportMonthKeys.includes(k)),
    [allMonthKeys, exportMonthKeys]
  );

  const exportScopeWarning = useMemo(() => {
    if (filterMode === "all" || excludedMonthKeys.length === 0) return "";
    const preview = excludedMonthKeys.slice(0, 3).map(monthLabel).join(", ");
    const extra = excludedMonthKeys.length > 3 ? ` +${excludedMonthKeys.length - 3} more` : "";
    return `Export scope excludes ${excludedMonthKeys.length} month(s): ${preview}${extra}.`;
  }, [filterMode, excludedMonthKeys]);

  const handleFile = useCallback((file) => {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => setChatText(e.target.result);
    reader.readAsText(file);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const handleParse = () => {
    if (!chatText.trim()) return;
    if (chatType === "baling") {
      setRecords(parseBalingMessages(chatText));
      setSummaryRecords([]);
      setValidationLog([]);
    } else if (cuttingMode === "new") {
      const { records, summaryRecords, validationLog } = parseCuttingMessagesNew(chatText);
      setRecords(records);
      setSummaryRecords(summaryRecords);
      setValidationLog(validationLog);
    } else {
      const { records, summaryRecords, validationLog } = parseCuttingMessages(chatText);
      setRecords(records);
      setSummaryRecords(summaryRecords);
      setValidationLog(validationLog);
    }
    setParsed(true);
  };

  const handleDownload = async () => {
    const defaultFilename = chatType === "baling"
      ? "BPR_Production_Data.xlsx"
      : "BPR_Cutting_Data.xlsx";
    const cleaned = String(exportFileName || "")
      .trim()
      .replace(/[\\/:*?"<>|]/g, "")
      .replace(/\s+/g, " ");
    const filename = cleaned
      ? (cleaned.toLowerCase().endsWith(".xlsx") ? cleaned : `${cleaned}.xlsx`)
      : defaultFilename;
    if (exportScopeWarning) {
      const ok = window.confirm(`${exportScopeWarning}\n\nContinue with scoped export?`);
      if (!ok) return;
    }
    try {
      await downloadXLSX(visibleRecords, chatType, filename, {
        summaryRecords: filteredSummaryRecords,
        validationLog: filteredValidationLog
      });
    } catch (error) {
      console.error("Excel export failed:", error);
      window.alert("Excel export failed. Please try again.");
    }
  };

  const handleReset = () => {
    setChatType(null);
    setCuttingMode("old");
    setChatText("");
    setRecords([]);
    setSummaryRecords([]);
    setValidationLog([]);
    setParsed(false);
    setFileName("");
    setFilterMode("all");
    setFilterYear("");
    setFilterPeriod("");
    setPlaceholderPolicy("keep");
    setExportFileName("");
  };

  // ─── Styles ──────────────────────────────────────────────────────────────────

  const accent = "#1B2E5C";
  const accentLight = "#EDF1F7";
  const warmBg = "#F8FAFC";
  const darkText = "#0F172A";
  const mutedText = "#64748B";
  const borderColor = "#E2E8F0";

  return (
    <div style={{
      minHeight: "100vh",
      background: warmBg,
      fontFamily: "'Inter', 'Segoe UI', 'Helvetica Neue', sans-serif",
      color: darkText,
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      padding: "32px 16px"
    }}>
      {/* Header */}
      <div style={{ textAlign: "center", marginBottom: 40, maxWidth: 600 }}>
        <div style={{
          display: "inline-flex",
          alignItems: "center",
          gap: 12,
          marginBottom: 12
        }}>
          <img
            src="/logo.png"
            alt="Blue Pyramid"
            style={{ height: 56, objectFit: "contain" }}
          />
        </div>
        <p style={{
          fontSize: 15, color: mutedText, margin: 0, lineHeight: 1.5
        }}>
          My most humble gift to my love 💙💙💜
        </p>
      </div>

      {/* Step 1: Choose type */}
      {!chatType && (
        <div style={{
          display: "flex", gap: 16, flexWrap: "wrap", justifyContent: "center"
        }}>
          {[
            { id: "baling", icon: "📦", title: "Baling Production", desc: "Bale reports with weights, operators & tyre counts" },
            { id: "cutting", icon: "✂\uFE0F", title: "Cutting Data", desc: "Hourly cutting machine counts per tyre type" }
          ].map((opt) => (
            <button
              key={opt.id}
              onClick={() => setChatType(opt.id)}
              style={{
                width: 260, padding: "28px 24px",
                background: "white",
                border: `2px solid ${borderColor}`,
                borderRadius: 16,
                cursor: "pointer",
                textAlign: "left",
                transition: "all 0.2s",
                boxShadow: "0 1px 3px rgba(0,0,0,0.04)"
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.borderColor = accent;
                e.currentTarget.style.boxShadow = `0 0 0 3px ${accentLight}`;
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.borderColor = borderColor;
                e.currentTarget.style.boxShadow = "0 1px 3px rgba(0,0,0,0.04)";
              }}
            >
              <span style={{ fontSize: 32 }}>{opt.icon}</span>
              <h3 style={{ fontSize: 17, fontWeight: 600, margin: "12px 0 6px" }}>{opt.title}</h3>
              <p style={{ fontSize: 13, color: mutedText, margin: 0, lineHeight: 1.4 }}>{opt.desc}</p>
            </button>
          ))}
        </div>
      )}

      {/* Step 2: Upload / paste */}
      {chatType && !parsed && (
        <div style={{ width: "100%", maxWidth: 560 }}>
          <div style={{
            display: "flex", alignItems: "center", gap: 8, marginBottom: 20
          }}>
            <button onClick={handleReset} style={{
              background: "none", border: "none", cursor: "pointer",
              fontSize: 14, color: mutedText, padding: "4px 0",
              display: "flex", alignItems: "center", gap: 4
            }}>
              ← Back
            </button>
            <span style={{ color: "#D1D5DB" }}>|</span>
            <span style={{ fontSize: 14, fontWeight: 600 }}>
              {chatType === "baling" ? "📦 Baling Production" : "✂️ Cutting Data"}
            </span>
          </div>

          {/* Format selector (cutting only) */}
          {chatType === "cutting" && (
            <div style={{
              display: "flex", alignItems: "center", gap: 12, marginBottom: 16,
              fontSize: 13, color: mutedText
            }}>
              <span style={{ fontWeight: 500 }}>Format:</span>
              {[
                { value: "old", label: "Old WhatsApp format" },
                { value: "new", label: "New WhatsApp format" },
              ].map((opt) => (
                <label key={opt.value} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
                  <input
                    type="radio"
                    name="cuttingMode"
                    value={opt.value}
                    checked={cuttingMode === opt.value}
                    onChange={() => setCuttingMode(opt.value)}
                    style={{ accentColor: accent }}
                  />
                  {opt.label}
                </label>
              ))}
            </div>
          )}

          {/* Drop zone */}
          <div
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={handleDrop}
            onClick={() => fileRef.current?.click()}
            style={{
              border: `2px dashed ${dragOver ? accent : "#D1D5DB"}`,
              borderRadius: 16,
              padding: "40px 24px",
              textAlign: "center",
              cursor: "pointer",
              background: dragOver ? accentLight : "white",
              transition: "all 0.2s",
              marginBottom: 16
            }}
          >
            <input
              ref={fileRef}
              type="file"
              accept=".txt,.zip"
              style={{ display: "none" }}
              onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])}
            />
            <div style={{ fontSize: 36, marginBottom: 8 }}>📂</div>
            <p style={{ fontSize: 15, fontWeight: 500, margin: "0 0 4px" }}>
              {fileName || "Drop your WhatsApp .txt export here"}
            </p>
            <p style={{ fontSize: 13, color: mutedText, margin: 0 }}>
              or click to browse
            </p>
          </div>

          {/* Or paste */}
          <div style={{ marginBottom: 16 }}>
            <p style={{
              fontSize: 12, color: mutedText, textAlign: "center",
              margin: "0 0 8px", textTransform: "uppercase", letterSpacing: 1
            }}>or paste chat text</p>
            <textarea
              value={chatText}
              onChange={(e) => setChatText(e.target.value)}
              placeholder="Paste the exported WhatsApp chat text here..."
              style={{
                width: "100%", minHeight: 140,
                border: "1px solid #E5E7EB", borderRadius: 12,
                padding: 16, fontSize: 13,
                fontFamily: "monospace",
                resize: "vertical",
                background: "white",
                boxSizing: "border-box"
              }}
            />
          </div>

          <button
            onClick={handleParse}
            disabled={!chatText.trim()}
            style={{
              width: "100%",
              padding: "14px 0",
              background: chatText.trim() ? accent : "#D1D5DB",
              color: "white",
              border: "none",
              borderRadius: 12,
              fontSize: 15,
              fontWeight: 600,
              cursor: chatText.trim() ? "pointer" : "default",
              transition: "background 0.2s"
            }}
          >
            Parse Chat →
          </button>
        </div>
      )}

      {/* Step 3: Results */}
      {parsed && (
        <div style={{ width: "100%", maxWidth: 900 }}>
          <div style={{
            display: "flex", alignItems: "center", justifyContent: "space-between",
            marginBottom: 16, flexWrap: "wrap", gap: 12
          }}>
            <div>
              <h2 style={{ fontSize: 20, fontWeight: 700, margin: 0 }}>
                {records.length} records parsed
              </h2>
              <p style={{ fontSize: 13, color: mutedText, margin: "4px 0 0" }}>
                {chatType === "baling" ? "Baling production entries" : "Cutting data entries"}
              </p>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={handleReset} style={{
                padding: "10px 20px", background: "white",
                border: "1px solid #D1D5DB", borderRadius: 10,
                fontSize: 14, cursor: "pointer"
              }}>
                Start Over
              </button>
              <button onClick={handleDownload} disabled={visibleRecords.length === 0} style={{
                padding: "10px 24px",
                background: visibleRecords.length > 0 ? accent : "#D1D5DB", color: "white",
                border: "none", borderRadius: 10,
                fontSize: 14, fontWeight: 600,
                cursor: visibleRecords.length > 0 ? "pointer" : "default"
              }}>
                ⬇ Download Excel
              </button>
            </div>
          </div>

          {/* Date range filter */}
          <div style={{
            display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap",
            marginBottom: 16, padding: "12px 16px",
            background: "white", borderRadius: 12,
            border: `1px solid ${borderColor}`
          }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: mutedText }}>📅 Export:</span>
            {[
              { value: "all", label: "All" },
              { value: "month", label: "Month" },
              { value: "quarter", label: "Quarter" },
              { value: "year", label: "Year" },
            ].map(opt => (
              <button
                key={opt.value}
                onClick={() => { setFilterMode(opt.value); setFilterYear(""); setFilterPeriod(""); }}
                style={{
                  padding: "6px 14px", borderRadius: 8,
                  border: filterMode === opt.value ? `2px solid ${accent}` : `1px solid ${borderColor}`,
                  background: filterMode === opt.value ? accentLight : "white",
                  color: filterMode === opt.value ? accent : darkText,
                  fontWeight: filterMode === opt.value ? 600 : 400,
                  fontSize: 13, cursor: "pointer",
                  transition: "all 0.15s",
                  opacity: filterMode !== "all" && filterMode !== opt.value ? 0.4 : 1
                }}
              >
                {opt.label}
              </button>
            ))}

            {filterMode !== "all" && (
              <>
                <select
                  value={filterYear}
                  onChange={(e) => setFilterYear(e.target.value)}
                  style={{
                    padding: "6px 12px", borderRadius: 8,
                    border: `1px solid ${borderColor}`, fontSize: 13,
                    background: "white", cursor: "pointer"
                  }}
                >
                  <option value="">Year…</option>
                  {filterOptions.years.map(y => (
                    <option key={y} value={y}>{y}</option>
                  ))}
                </select>

                {filterMode === "month" && (
                  <select
                    value={filterPeriod}
                    onChange={(e) => setFilterPeriod(e.target.value)}
                    style={{
                      padding: "6px 12px", borderRadius: 8,
                      border: `1px solid ${borderColor}`, fontSize: 13,
                      background: "white", cursor: "pointer"
                    }}
                  >
                    <option value="">Month…</option>
                    {MONTH_NAMES.map((name, i) => (
                      <option key={i + 1} value={i + 1}>{name}</option>
                    ))}
                  </select>
                )}

                {filterMode === "quarter" && (
                  <select
                    value={filterPeriod}
                    onChange={(e) => setFilterPeriod(e.target.value)}
                    style={{
                      padding: "6px 12px", borderRadius: 8,
                      border: `1px solid ${borderColor}`, fontSize: 13,
                      background: "white", cursor: "pointer"
                    }}
                  >
                    <option value="">Quarter…</option>
                    <option value="1">Q1 (Jan–Mar)</option>
                    <option value="2">Q2 (Apr–Jun)</option>
                    <option value="3">Q3 (Jul–Sep)</option>
                    <option value="4">Q4 (Oct–Dec)</option>
                  </select>
                )}
              </>
            )}

            {filterMode !== "all" && (filterYear || filterPeriod) && (
              <span style={{ fontSize: 12, color: mutedText, marginLeft: 4 }}>
                {visibleRecords.length} of {records.length} records
              </span>
            )}

            {chatType === "cutting" && (
              <select
                value={placeholderPolicy}
                onChange={(e) => setPlaceholderPolicy(e.target.value)}
                style={{
                  marginLeft: "auto",
                  padding: "6px 10px",
                  borderRadius: 8,
                  border: `1px solid ${borderColor}`,
                  fontSize: 12,
                  background: "white",
                  cursor: "pointer"
                }}
              >
                <option value="keep">Keep placeholders</option>
                <option value="strict">Strict source only</option>
              </select>
            )}

            <input
              type="text"
              value={exportFileName}
              onChange={(e) => setExportFileName(e.target.value)}
              placeholder={chatType === "baling" ? "BPR_Production_Data.xlsx" : "BPR_Cutting_Data.xlsx"}
              style={{
                marginLeft: chatType === "cutting" ? 0 : "auto",
                padding: "6px 10px",
                borderRadius: 8,
                border: `1px solid ${borderColor}`,
                fontSize: 12,
                background: "white",
                minWidth: 220
              }}
              aria-label="Export filename"
            />
          </div>

          {exportScopeWarning && (
            <p style={{ margin: "-8px 0 12px", fontSize: 12, color: "#B45309" }}>
              ⚠ {exportScopeWarning}
            </p>
          )}

          {/* Preview table */}
          <div style={{
            background: "white", borderRadius: 16,
            border: "1px solid #E5E7EB",
            overflow: "auto",
            boxShadow: "0 1px 3px rgba(0,0,0,0.04)"
          }}>
            {chatType === "baling" ? (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: "#F9FAFB" }}>
                    {["Date", "Machine", "Type", "Bale #", "Series", "Operator", "Assistant",
                      "Start", "Finish", "Passenger", "4x4", "Agri T", "Agri SW", "Weight (kg)"].map((h) => (
                      <th key={h} style={{
                        padding: "10px 8px", textAlign: "left",
                        borderBottom: "1px solid #E5E7EB",
                        fontWeight: 600, whiteSpace: "nowrap", color: mutedText,
                        fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5
                      }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {visibleRecords.slice(0, 100).map((r, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #F3F4F6" }}>
                      <td style={{ padding: "8px" }}>{r.date ? dateToStr(r.date) : "—"}</td>
                      <td style={{ padding: "8px" }}>{r.machine || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.baleType || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.baleNum || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.baleSeries || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.operator || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.assistant || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.startTime ? formatTime(r.startTime) : "—"}</td>
                      <td style={{ padding: "8px" }}>{r.finishTime ? formatTime(r.finishTime) : "—"}</td>
                      <td style={{ padding: "8px" }}>{r.passenger || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.fourx4 || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.agriT || "—"}</td>
                      <td style={{ padding: "8px" }}>{r.agriSW || "—"}</td>
                      <td style={{ padding: "8px", fontWeight: 600 }}>{r.weight}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: "#F9FAFB" }}>
                    {["Date", "Machine", "Time", "LC", "HC", "Agri", "Tread LC", "Tread HC", "Tread Agri"].map((h) => (
                      <th key={h} style={{
                        padding: "10px 8px", textAlign: "left",
                        borderBottom: "1px solid #E5E7EB",
                        fontWeight: 600, whiteSpace: "nowrap", color: mutedText,
                        fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5
                      }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {visibleRecords.slice(0, 100).map((r, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #F3F4F6" }}>
                      <td style={{ padding: "8px" }}>{dateToStr(r.date)}</td>
                      <td style={{ padding: "8px" }}>{r.cmNumber}</td>
                      <td style={{ padding: "8px" }}>
                        {r.startTime && r.finishTime ? `${formatTime(r.startTime)}-${formatTime(r.finishTime)}` : "—"}
                      </td>
                      <td style={{ padding: "8px", fontWeight: 600 }}>{r.light_commercial ?? "—"}</td>
                      <td style={{ padding: "8px", fontWeight: 600 }}>{r.heavy_commercial_t ?? "—"}</td>
                      <td style={{ padding: "8px", fontWeight: 600 }}>{r.agricultural_t ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.tread_lc ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.tread_hc ?? "—"}</td>
                      <td style={{ padding: "8px" }}>{r.tread_agri ?? "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
            {visibleRecords.length > 100 && (
              <p style={{
                textAlign: "center", padding: 12,
                fontSize: 12, color: mutedText
              }}>
                Showing first 100 of {visibleRecords.length} records. All {visibleRecords.length} will be included in the download.
              </p>
            )}
          </div>

          <p style={{
            fontSize: 12, color: mutedText, textAlign: "center",
            marginTop: 16, lineHeight: 1.5
          }}>
            💡 The Excel file has one sheet per month, sorted chronologically.
          </p>
        </div>
      )}
    </div>
  );
}
