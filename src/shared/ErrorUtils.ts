/**
 * Converts an unknown thrown value into a stable, log-friendly message.
 */
function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
