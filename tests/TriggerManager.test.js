/**
 * TriggerManager single-trigger queue tests.
 *
 * Covers creating a missing handler trigger, retaining the first match while
 * deleting duplicates, and clearing every trigger for a handler.
 */
const assert = require('node:assert/strict');
const test = require('node:test');

const { loadScripts } = require('./helpers/appScriptHarness');

/**
 * Builds a ScriptApp double that records create/delete operations against an
 * in-memory project trigger list, matching the order getProjectTriggers returns.
 */
function createScriptApp(initialTriggers = []) {
  const triggers = [...initialTriggers];
  const deleted = [];
  const created = [];

  return {
    triggers,
    deleted,
    created,
    ScriptApp: {
      getProjectTriggers() {
        return [...triggers];
      },
      deleteTrigger(trigger) {
        deleted.push(trigger);
        const index = triggers.indexOf(trigger);
        if (index >= 0) {
          triggers.splice(index, 1);
        }
      },
      newTrigger(handlerFunction) {
        const builder = {
          timeBased() {
            return builder;
          },
          everyMinutes(intervalMinutes) {
            builder.intervalMinutes = intervalMinutes;
            return builder;
          },
          create() {
            const trigger = {
              handlerFunction,
              intervalMinutes: builder.intervalMinutes,
              getHandlerFunction() {
                return handlerFunction;
              },
            };
            created.push(trigger);
            triggers.push(trigger);
            return trigger;
          },
        };
        return builder;
      },
    },
  };
}

/**
 * Creates a trigger double identified by handler name for assertion identity.
 */
function createTrigger(handlerFunction, id) {
  return {
    id,
    getHandlerFunction() {
      return handlerFunction;
    },
  };
}

/**
 * Loads TriggerManager into a sandbox with the provided ScriptApp double.
 */
function loadTriggerManager(ScriptApp) {
  return loadScripts({ ScriptApp }, [
    {
      path: 'src/infrastructure/TriggerManager.ts',
      exports: ['TriggerManager'],
    },
  ]);
}

test('ensureSingle creates a time-driven trigger when none exist', () => {
  const { ScriptApp, created, triggers } = createScriptApp();
  const { TriggerManager } = loadTriggerManager(ScriptApp);

  TriggerManager.ensureSingle('update', 5);

  assert.equal(created.length, 1);
  assert.equal(created[0].handlerFunction, 'update');
  assert.equal(created[0].intervalMinutes, 5);
  assert.equal(triggers.length, 1);
});

test('ensureSingle keeps the first matching trigger and deletes duplicates', () => {
  const oldest = createTrigger('update', 'oldest');
  const duplicate = createTrigger('update', 'duplicate');
  const other = createTrigger('otherHandler', 'other');
  const { ScriptApp, created, deleted, triggers } = createScriptApp([
    oldest,
    duplicate,
    other,
  ]);
  const { TriggerManager } = loadTriggerManager(ScriptApp);

  TriggerManager.ensureSingle('update', 5);

  assert.deepEqual(deleted, [duplicate]);
  assert.equal(created.length, 0);
  assert.deepEqual(triggers, [oldest, other]);
});

test('ensureSingle leaves a single matching trigger unchanged', () => {
  const existing = createTrigger('update', 'only');
  const { ScriptApp, created, deleted, triggers } = createScriptApp([existing]);
  const { TriggerManager } = loadTriggerManager(ScriptApp);

  TriggerManager.ensureSingle('update', 5);

  assert.equal(created.length, 0);
  assert.equal(deleted.length, 0);
  assert.deepEqual(triggers, [existing]);
});

test('removeAll deletes every trigger for the handler and ignores others', () => {
  const first = createTrigger('update', 'first');
  const second = createTrigger('update', 'second');
  const other = createTrigger('otherHandler', 'other');
  const { ScriptApp, deleted, triggers } = createScriptApp([
    first,
    second,
    other,
  ]);
  const { TriggerManager } = loadTriggerManager(ScriptApp);

  TriggerManager.removeAll('update');

  assert.deepEqual(deleted, [first, second]);
  assert.deepEqual(triggers, [other]);
});
