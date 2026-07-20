/**
 * Returns a new date offset by the specified number of calendar days.
 *
 * The input is never mutated because spreadsheet values can be reused while a
 * batch is being evaluated.
 */
function addDays(date: Date, days: number): Date {
  const result = new Date(date.getTime());
  result.setDate(result.getDate() + days);
  return result;
}

/**
 * Determines whether a timestamp remains within a refresh window.
 */
function isWithinRefreshWindow(
  timestamp: Date,
  current: Date,
  refreshIntervalDays: number,
): boolean {
  return addDays(timestamp, refreshIntervalDays) > current;
}
