/**
 * Creates and removes time-driven triggers without allowing duplicates.
 *
 * The unified update queue depends on exactly one recurring trigger. Duplicate
 * handlers would resume the same phase concurrently and contend for sheet locks.
 */
class TriggerManager {
  /**
   * Ensures that exactly one time-driven trigger exists for a handler.
   *
   * Retaining the oldest trigger avoids continually changing its schedule while
   * still repairing duplicate triggers left by prior executions or failed
   * cleanups.
   */
  static ensureSingle(handlerFunction: string, intervalMinutes: number): void {
    const matchingTriggers = ScriptApp.getProjectTriggers().filter(
      (trigger) => trigger.getHandlerFunction() === handlerFunction,
    );

    // Keep the first match; extras are leftovers from overlapping starts.
    matchingTriggers
      .slice(1)
      .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

    if (matchingTriggers.length === 0) {
      ScriptApp.newTrigger(handlerFunction)
        .timeBased()
        .everyMinutes(intervalMinutes)
        .create();
    }
  }

  /**
   * Removes every trigger that invokes a handler.
   *
   * Used when a cycle finishes or when a manual update supersedes the previous
   * queue so orphaned triggers cannot keep firing after state is cleared.
   */
  static removeAll(handlerFunction: string): void {
    ScriptApp.getProjectTriggers()
      .filter((trigger) => trigger.getHandlerFunction() === handlerFunction)
      .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  }
}
