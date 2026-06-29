class Triggers {
  static ensure(handlerFunction: string, intervalMinutes: number): void {
    const exists = ScriptApp.getProjectTriggers().some(
      (trigger) => trigger.getHandlerFunction() === handlerFunction,
    );
    if (!exists) {
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
