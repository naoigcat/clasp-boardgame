/**
 * Creates and removes time-driven triggers without allowing duplicates.
 */
class TriggerManager {
  /**
   * Ensures that exactly one time-driven trigger exists for a handler.
   *
   * Retaining the oldest trigger avoids continually changing its schedule while
   * still repairing duplicate triggers left by prior executions.
   */
  static ensureSingle(
    handlerFunction: string,
    intervalMinutes: number,
  ): void {
    const matchingTriggers = ScriptApp.getProjectTriggers().filter(
      (trigger) => trigger.getHandlerFunction() === handlerFunction,
    );

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
   */
  static removeAll(handlerFunction: string): void {
    ScriptApp.getProjectTriggers()
      .filter((trigger) => trigger.getHandlerFunction() === handlerFunction)
      .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  }
}
