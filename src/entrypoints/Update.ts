/**
 * Runs the unified spreadsheet update from the menu or a time-driven trigger.
 *
 * Apps Script supplies an event only for triggered executions; using its
 * presence keeps the menu action separate from queue-resume behavior.
 */
function update(event?: unknown): void {
  UpdateCoordinator.run(event !== undefined);
}
