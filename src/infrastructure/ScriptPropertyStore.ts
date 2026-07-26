/**
 * Provides typed access to Apps Script project properties.
 *
 * Script properties hold both durable configuration (tokens, user IDs) and
 * transient queue state. Wrapping PropertiesService keeps those concerns out of
 * individual services.
 */
class ScriptPropertyStore {
  /**
   * Returns an optional string property without exposing the Apps Script API to
   * every service that needs configuration.
   */
  static getOptionalValue(propertyKey: string): string | null {
    return PropertiesService.getScriptProperties().getProperty(propertyKey);
  }

  /**
   * Removes a transient property that should not survive a completed update.
   *
   * Queue phase keys are deleted on finish so a later scheduled fire cannot
   * resume work that was already marked complete.
   */
  static remove(propertyKey: string): void {
    PropertiesService.getScriptProperties().deleteProperty(propertyKey);
  }

  /**
   * Persists the next state of a multi-invocation update.
   *
   * The value must survive the current execution so the next trigger knows
   * whether to continue Games batches or switch to Titles.
   */
  static set(propertyKey: string, value: string): void {
    PropertiesService.getScriptProperties().setProperty(propertyKey, value);
  }
}
