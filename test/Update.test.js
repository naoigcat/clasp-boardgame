/**
 * UpdateCoordinator queue lifecycle tests.
 *
 * Covers lock skipping, menu-driven restarts, Games-to-Titles phase transitions,
 * and cleanup of the shared trigger when the queued work finishes.
 */
const assert = require('node:assert/strict');
const test = require('node:test');

const { loadScripts } = require('./helpers/appScriptHarness');

/**
 * Creates service and Apps Script doubles for update-queue behavior tests.
 */
function createUpdateSandbox(gameResults, titleResults, lockAcquired = true) {
  const calls = [];
  const properties = new Map();

  return {
    calls,
    LockService: {
      getScriptLock() {
        return {
          tryLock() {
            return lockAcquired;
          },
          releaseLock() {},
        };
      },
    },
    Logger: {
      log() {},
    },
    PropertiesService: {
      getScriptProperties() {
        return {
          deleteProperty(key) {
            calls.push(['deleteProperty', key]);
            properties.delete(key);
          },
          getProperty(key) {
            calls.push(['getProperty', key]);
            return properties.get(key) || null;
          },
          setProperty(key, value) {
            calls.push(['setProperty', key, value]);
            properties.set(key, value);
          },
        };
      },
    },
    TriggerManager: {
      removeAll(handler) {
        calls.push(['deleteTrigger', handler]);
      },
      ensureSingle(handler, intervalMinutes) {
        calls.push(['ensureTrigger', handler, intervalMinutes]);
      },
    },
    GameUpdater: {
      run() {
        calls.push(['games']);
        return gameResults.shift();
      },
    },
    RankingUpdater: {
      run() {
        calls.push(['rankings']);
      },
    },
    RatingUpdater: {
      run() {
        calls.push(['ratings']);
      },
    },
    TitleUpdater: {
      run() {
        calls.push(['titles']);
        return titleResults.shift();
      },
    },
  };
}

/**
 * Loads the coordinator through its public Apps Script entry point.
 */
function loadUpdate(sandbox) {
  return loadScripts(sandbox, [
    { path: 'src/config/AppConfig.ts', exports: [] },
    {
      path: 'src/infrastructure/ScriptPropertyStore.ts',
      exports: ['ScriptPropertyStore'],
    },
    {
      path: 'src/services/UpdateCoordinator.ts',
      exports: ['UpdateCoordinator'],
    },
    { path: 'src/entrypoints/Update.ts', exports: ['update'] },
  ]);
}

test('update runs single-execution work before scheduling the games phase', () => {
  const sandbox = createUpdateSandbox([], []);
  const context = loadUpdate(sandbox);

  context.update();

  assert.deepEqual(sandbox.calls, [
    ['deleteTrigger', 'updateGames'],
    ['deleteTrigger', 'updateRankings'],
    ['deleteTrigger', 'updateTitles'],
    ['deleteTrigger', 'updateRatings'],
    ['deleteTrigger', 'update'],
    ['deleteProperty', 'UPDATE_STEP'],
    ['rankings'],
    ['ratings'],
    ['setProperty', 'UPDATE_STEP', 'games'],
    ['ensureTrigger', 'update', 5],
  ]);
});

test('scheduled updates complete games before starting titles with the same handler', () => {
  const sandbox = createUpdateSandbox([true, false], [false]);
  const context = loadUpdate(sandbox);

  context.update();
  sandbox.calls.length = 0;

  context.update({ triggerUid: 'first' });
  assert.deepEqual(sandbox.calls, [
    ['getProperty', 'UPDATE_STEP'],
    ['games'],
  ]);

  sandbox.calls.length = 0;
  context.update({ triggerUid: 'second' });
  assert.deepEqual(sandbox.calls, [
    ['getProperty', 'UPDATE_STEP'],
    ['games'],
    ['setProperty', 'UPDATE_STEP', 'titles'],
  ]);

  sandbox.calls.length = 0;
  context.update({ triggerUid: 'third' });
  assert.deepEqual(sandbox.calls, [
    ['getProperty', 'UPDATE_STEP'],
    ['titles'],
    ['deleteProperty', 'UPDATE_STEP'],
    ['deleteTrigger', 'update'],
  ]);
});

test('update does not start a second phase while another update holds the lock', () => {
  const sandbox = createUpdateSandbox([], [], false);
  const context = loadUpdate(sandbox);

  context.update();

  assert.deepEqual(sandbox.calls, []);
});
