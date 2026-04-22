export async function workbookToBrowserDownload(wb, filename) {
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

export function baseStyles() {
  const brandBlue = "FF1B2E5C";
  const brandBlueLight = "FFEDF1F7";
  const textWhite = "FFFFFFFF";
  const textDark = "FF0F172A";
  const thinBlack = { style: "thin", color: { argb: "FF000000" } };
  const mediumBlack = { style: "medium", color: { argb: "FF000000" } };
  const thickBlack = { style: "thick", color: { argb: "FF000000" } };

  const baseBorder = {
    top: { style: "thin", color: { argb: "FFD1D5DB" } },
    left: { style: "thin", color: { argb: "FFD1D5DB" } },
    bottom: { style: "thin", color: { argb: "FFD1D5DB" } },
    right: { style: "thin", color: { argb: "FFD1D5DB" } },
  };

  return { brandBlue, brandBlueLight, textWhite, textDark, thinBlack, mediumBlack, thickBlack, baseBorder };
}

export function styleHeaderRow(row, styles, isLight = false) {
  row.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: isLight ? styles.textDark : styles.textWhite } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: isLight ? styles.brandBlueLight : styles.brandBlue },
    };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = styles.baseBorder;
  });
}

export function styleBodyRows(ws, fromRow, toRow, baseBorder, colCount) {
  for (let r = fromRow; r <= toRow; r += 1) {
    const row = ws.getRow(r);
    const cols = colCount || row.cellCount;
    for (let c = 1; c <= cols; c += 1) {
      const cell = row.getCell(c);
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = baseBorder;
    }
  }
}

export function applyColumnWidths(ws, widths) {
  ws.columns = widths.map((w) => ({ width: w }));
}
