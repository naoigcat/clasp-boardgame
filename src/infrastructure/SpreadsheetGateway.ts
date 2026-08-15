/**
 * Thin adapters around SpreadsheetApp used by the update services.
 *
 * Isolating these calls keeps services focused on board-game data rules and
 * makes spreadsheet edge cases (missing tabs, empty data ranges) explicit.
 */

/**
 * Retrieves a sheet from the active spreadsheet by its configured name.
 *
 * Returning null lets callers skip work when a tab has been renamed or removed
 * instead of failing the entire update cycle.
 */
function findSheet(
  sheetName: string,
): GoogleAppsScript.Spreadsheet.Sheet | null {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

/**
 * Clears only data rows below a header when such rows exist.
 *
 * Apps Script rejects zero-height ranges, so this guard keeps an empty sheet a
 * valid state rather than an exceptional one. Headers stay intact because data
 * rewrites always start at row 2. The optional start column lets Games preserve
 * its rich-text links while clearing only the value columns it owns.
 */
function clearSheetDataRows(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  columnCount: number,
  startColumn: number = SHEET_LAYOUT.DEFAULT_START_COLUMN,
): void {
  const dataRowCount = sheet.getLastRow() - SHEET_LAYOUT.FIRST_DATA_ROW + 1;
  if (dataRowCount <= 0) {
    return;
  }

  sheet
    .getRange(
      SHEET_LAYOUT.FIRST_DATA_ROW,
      startColumn,
      dataRowCount,
      columnCount,
    )
    .clearContent();
}

/**
 * Replaces a sheet snapshot and removes rows abandoned by a shorter payload.
 *
 * Writing before trimming keeps the previous snapshot available when the new
 * write fails; the helper owns that ordering so each updater preserves the same
 * failure behavior without duplicating range arithmetic.
 */
function writeSheetSnapshot(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rows: readonly SpreadsheetCellRow[],
  columnCount: number,
  startColumn: number = SHEET_LAYOUT.DEFAULT_START_COLUMN,
): void {
  if (rows.length === 0) {
    clearSheetDataRows(sheet, columnCount, startColumn);
    return;
  }

  sheet
    .getRange(
      SHEET_LAYOUT.FIRST_DATA_ROW,
      startColumn,
      rows.length,
      columnCount,
    )
    .setValues(escapeSheetValues(rows));
  clearSurplusSheetDataRows(sheet, rows.length, columnCount, startColumn);
}

/**
 * Clears only rows below a just-written data block.
 *
 * Titles, Rankings, Ratings, and Games write first, then trim abandoned physical
 * rows. Clearing after setValues avoids a window where a failed rewrite would
 * leave a header-only sheet and erase data that cannot be restored until the
 * next successful fetch. Games passes a start column of B so surplus cleanup
 * never touches BoardGameGeek rich-text links in column A.
 */
function clearSurplusSheetDataRows(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  writtenRowCount: number,
  columnCount: number,
  startColumn: number = SHEET_LAYOUT.DEFAULT_START_COLUMN,
): void {
  const firstSurplusRow = SHEET_LAYOUT.FIRST_DATA_ROW + writtenRowCount;
  const surplusRowCount = sheet.getLastRow() - firstSurplusRow + 1;
  if (surplusRowCount <= 0) {
    return;
  }

  sheet
    .getRange(firstSurplusRow, startColumn, surplusRowCount, columnCount)
    .clearContent();
}
