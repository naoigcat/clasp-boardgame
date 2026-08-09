/**
 * Forces externally sourced strings to store as text in Google Sheets.
 *
 * Titles, tags, and similar fields can begin with `=`, `+`, `-`, `@`, tab, or
 * CR. Digit strings that start with `0` but are not exactly `"0"` are also
 * prefixed so Sheets does not drop leading zeros. A leading apostrophe keeps
 * those values as literals when the workbook is shared. `getValues` later
 * returns the text without that apostrophe, so in-memory matching stays unchanged.
 */
function escapeSheetCellValue(
  value: SpreadsheetCellValue,
): SpreadsheetCellValue {
  if (typeof value !== 'string' || value.length === 0) {
    return value;
  }

  switch (value.charAt(0)) {
    case '=':
    case '+':
    case '-':
    case '@':
    case '\t':
    case '\r':
      return `'${value}`;
    case '0':
      // Keep a lone "0"; escape longer digit strings so leading zeros survive.
      return /^0\d+$/.test(value) ? `'${value}` : value;
    default:
      return value;
  }
}

/**
 * Escapes every cell in a `setValues` payload while leaving non-strings alone.
 *
 * Applied at write time so services keep unescaped strings for comparisons and
 * retries; only the spreadsheet API receives the formula-safe copy.
 */
function escapeSheetValues(
  rows: readonly SpreadsheetCellRow[],
): SpreadsheetCellRow[] {
  return rows.map((row) => row.map(escapeSheetCellValue));
}
