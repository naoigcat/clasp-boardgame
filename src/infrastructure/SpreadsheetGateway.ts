/**
 * Retrieves a sheet from the active spreadsheet by its configured name.
 */
function findSheet(
  sheetName: string,
): GoogleAppsScript.Spreadsheet.Sheet | null {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

/**
 * Clears only data rows below a header when such rows exist.
 *
 * Apps Script rejects zero-height ranges, so this guard keeps empty sheets a
 * valid state rather than an exceptional one.
 */
function clearSheetDataRows(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  columnCount: number,
): void {
  const dataRowCount = sheet.getLastRow() - SHEET_LAYOUT.FIRST_DATA_ROW + 1;
  if (dataRowCount <= 0) {
    return;
  }

  sheet
    .getRange(
      SHEET_LAYOUT.FIRST_DATA_ROW,
      SHEET_LAYOUT.DEFAULT_START_COLUMN,
      dataRowCount,
      columnCount,
    )
    .clearContent();
}
