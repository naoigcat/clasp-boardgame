/**
 * Runs the unified spreadsheet update from the menu or a time-driven trigger.
 *
 * Apps Script supplies an event object only for triggered executions. Detecting
 * that presence keeps a manual menu click on the "start a new cycle" path while
 * scheduled runs resume the saved queue phase instead of restarting.
 */
function update(event?: unknown): void {
  UpdateCoordinator.run(event !== undefined);
}
