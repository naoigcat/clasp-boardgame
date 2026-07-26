/**
 * Values that Apps Script reads from and writes to an ordinary spreadsheet cell.
 *
 * Rich-text links are handled separately because `getValues` / `setValues` do
 * not preserve hyperlink formatting on their own.
 */
type SpreadsheetCellValue = string | number | boolean | Date | null;

/**
 * A mutable sequence of spreadsheet cell values in one row.
 *
 * Services mutate these arrays in memory and flush them with one `setValues`
 * call so a batch does not thrash the spreadsheet API per cell.
 */
type SpreadsheetCellRow = SpreadsheetCellValue[];
