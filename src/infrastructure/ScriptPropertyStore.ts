/**
 * Provides typed access to Apps Script project properties.
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
   */
  static remove(propertyKey: string): void {
    PropertiesService.getScriptProperties().deleteProperty(propertyKey);
  }

  /**
   * Persists the next state of a multi-invocation update.
   */
  static set(propertyKey: string, value: string): void {
    PropertiesService.getScriptProperties().setProperty(propertyKey, value);
  }
}
