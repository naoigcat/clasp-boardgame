/**
 * Returns a new date offset by the specified number of calendar days.
 *
 * The input is never mutated because spreadsheet Date values can be reused
 * while a batch is being evaluated; mutating them would corrupt later writes.
 */
function addDays(date: Date, days: number): Date {
  const result = new Date(date.getTime());
  result.setDate(result.getDate() + days);
  return result;
}

/**
 * Determines whether a timestamp remains within a refresh window.
 *
 * Used by Games to skip BoardGameGeek rows that were attempted recently enough
 * that another fetch would waste quota without improving freshness.
 */
function isWithinRefreshWindow(
  timestamp: Date,
  current: Date,
  refreshIntervalDays: number,
): boolean {
  return addDays(timestamp, refreshIntervalDays) > current;
}

/**
 * Determines whether an Apps Script invocation has used its soft time budget.
 *
 * Games and Titles stop starting new upstream requests once this returns true so
 * progress can be flushed before the platform's hard execution limit aborts the
 * run. Remaining rows resume on the next shared trigger.
 */
function hasExceededRuntime(
  startedAtMilliseconds: number,
  maxRuntimeMilliseconds: number,
): boolean {
  return Date.now() - startedAtMilliseconds >= maxRuntimeMilliseconds;
}
