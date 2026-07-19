class Triggers {
  static ensure(handlerFunction: string, intervalMinutes: number): void {
    const triggers = ScriptApp.getProjectTriggers().filter(
      (trigger) => trigger.getHandlerFunction() === handlerFunction,
    );
    triggers.slice(1).forEach((trigger) => ScriptApp.deleteTrigger(trigger));
    if (triggers.length === 0) {
      ScriptApp.newTrigger(handlerFunction)
        .timeBased()
        .everyMinutes(intervalMinutes)
        .create();
    }
  }

  static deleteAll(handlerFunction: string): void {
    ScriptApp.getProjectTriggers()
      .filter((trigger) => trigger.getHandlerFunction() === handlerFunction)
      .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  }
}
