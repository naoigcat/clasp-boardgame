/**
 * Values that Apps Script reads from and writes to an ordinary spreadsheet cell.
 */
type SpreadsheetCellValue = string | number | boolean | Date | null;

/**
 * A mutable sequence of spreadsheet cell values in one row.
 */
type SpreadsheetCellRow = SpreadsheetCellValue[];
