import { dateToStr, dateSortKey, groupByMonth, monthLabel } from "../helpers.js";
import {
  workbookToBrowserDownload,
  baseStyles,
  styleHeaderRow,
  styleBodyRows,
  applyColumnWidths,
} from "./excelCommon.js";

const HEADERS = [
  "Date",
  "Truck No.",
  "WBD",
  "GRV",
  "Depot",
  "Depot Manager",
  "Collection No.",
  "Transporter",
  "Passenger",
  "4 x 4",
  "Motorcycle",
  "Light Commercial",
  "Heavy Commercial",
  "Heavy Commercial SW",
  "Heavy Commercial T",
  "Agricultural",
];

const COLUMN_WIDTHS = [13, 10, 10, 10, 18, 18, 14, 16, 12, 10, 12, 16, 16, 18, 18, 14];

const TYRE_COL_START = 9;
const TYRE_COL_END = 16;
const TOTAL_COLS = 16;

function recordToRow(r) {
  const cellVal = (v) => (v === null || v === undefined ? "" : v);
  return [
    r.date ? dateToStr(r.date) : "",
    cellVal(r.delTruckNo),
    cellVal(r.delWbd),
    cellVal(r.delGrv),
    cellVal(r.delDepot),
    cellVal(r.delDepotManager),
    cellVal(r.delCollectionNo),
    cellVal(r.delTransporter),
    cellVal(r.delPassenger),
    cellVal(r.delFourByFour),
    cellVal(r.delMotorcycle),
    cellVal(r.delLightCommercial),
    cellVal(r.delHeavyCommercial),
    cellVal(r.delHeavyCommercialSW),
    cellVal(r.delHeavyCommercialT),
    cellVal(r.delAgricultural),
  ];
}

function styleTotalsRow(row, styles) {
  const gray = "FFC2C8D6";
  for (let c = 1; c <= TOTAL_COLS; c += 1) {
    const cell = row.getCell(c);
    cell.font = { ...(cell.font || {}), bold: true, color: { argb: styles.textDark } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: gray } };
    cell.border = {
      top: styles.mediumBlack,
      bottom: styles.mediumBlack,
      left: c === 1 ? styles.mediumBlack : styles.baseBorder.left,
      right: c === TOTAL_COLS ? styles.mediumBlack : styles.baseBorder.right,
    };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  }
}

function addValidationLogSheet(wb, validationLog, styles) {
  if (!validationLog || validationLog.length === 0) return;
  const ws = wb.addWorksheet("Validation_Log");
  ws.addRow(["Date", "Source Timestamp", "Sender", "Issue", "Action", "Raw Text"]);
  for (const entry of validationLog) {
    ws.addRow([
      entry.date || "",
      entry.sourceMessageTimestamp || "",
      entry.sender || "",
      entry.issue || "",
      entry.action || "",
      entry.rawMessage || "",
    ]);
  }
  styleHeaderRow(ws.getRow(1), styles);
  styleBodyRows(ws, 2, ws.rowCount, styles.baseBorder, 6);
  applyColumnWidths(ws, [13, 16, 22, 48, 14, 80]);
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 6 } };
}

export async function downloadDeliveriesWorkbook(records, filename, extras = {}) {
  const { validationLog = [] } = extras;
  const { default: ExcelJS } = await import("exceljs");
  const wb = new ExcelJS.Workbook();
  wb.calcProperties.fullCalcOnLoad = true;
  wb.calcProperties.forceFullCalc = true;
  const styles = baseStyles();

  const byMonth = groupByMonth(records);
  const sortedKeys = Object.keys(byMonth).sort();

  if (sortedKeys.length === 0) {
    const ws = wb.addWorksheet("Deliveries");
    ws.addRow(HEADERS);
    styleHeaderRow(ws.getRow(1), styles);
    applyColumnWidths(ws, COLUMN_WIDTHS);
    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: TOTAL_COLS } };
  }

  for (const key of sortedKeys) {
    const monthRecords = [...byMonth[key]].sort(
      (a, b) => dateSortKey(a.date) - dateSortKey(b.date),
    );

    const ws = wb.addWorksheet(monthLabel(key));
    ws.addRow(HEADERS);
    for (const r of monthRecords) {
      ws.addRow(recordToRow(r));
    }

    const dataStartRow = 2;
    const dataEndRow = 1 + monthRecords.length;
    const totalsRowIndex = dataEndRow + 1;

    ws.addRow(["TOTALS"]);

    styleHeaderRow(ws.getRow(1), styles);
    if (monthRecords.length > 0) {
      styleBodyRows(ws, dataStartRow, dataEndRow, styles.baseBorder, TOTAL_COLS);
    }

    for (let col = TYRE_COL_START; col <= TYRE_COL_END; col += 1) {
      const totalCell = ws.getCell(totalsRowIndex, col);
      if (monthRecords.length > 0) {
        const colLetter = totalCell.address.replace(/\d+$/, "");
        totalCell.value = {
          formula: `SUM(${colLetter}${dataStartRow}:${colLetter}${dataEndRow})`,
        };
      } else {
        totalCell.value = 0;
      }
    }

    styleTotalsRow(ws.getRow(totalsRowIndex), styles);

    applyColumnWidths(ws, COLUMN_WIDTHS);
    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: TOTAL_COLS } };
  }

  addValidationLogSheet(wb, validationLog, styles);

  await workbookToBrowserDownload(wb, filename);
}
