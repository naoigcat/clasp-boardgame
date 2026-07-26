/**
 * Converts an unknown thrown value into a stable, log-gable message.
 *
 * Apps Script and TypeScript both allow non-Error throwables. Normalizing here
 * keeps Logger output and spreadsheet error cells readable without each caller
 * reimplementing the same defensive check.
 */
function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
