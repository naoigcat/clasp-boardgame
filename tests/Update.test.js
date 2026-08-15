/**
 * UpdateCoordinator queue lifecycle tests.
 *
 * Covers lock skipping, menu-driven restarts that schedule Games before sync
 * imports, Games-to-Titles phase transitions, and cleanup of the shared trigger
 * when the queued work finishes.
 */
const assert = require('node:assert/strict');
const test = require('node:test');

const { loadScripts } = require('./helpers/appScriptHarness');

/**
 * Creates service and Apps Script doubles for update-queue behavior tests.
 */
function createUpdateSandbox(
  gameResults,
  titleResults,
  lockAcquired = true,
  options = {},
) {
  const calls = [];
  const logs = [];
  const properties = new Map();

  return {
    calls,
    logs,
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
      log(message) {
        logs.push(message);
      },
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
        if (options.rankingError) {
          throw options.rankingError;
        }
      },
    },
    RatingUpdater: {
      run() {
        calls.push(['ratings']);
        if (options.ratingError) {
          throw options.ratingError;
        }
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
    { path: 'src/shared/ErrorUtils.ts', exports: ['getErrorMessage'] },
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

test('update schedules the games phase before single-execution sync imports', () => {
  const sandbox = createUpdateSandbox([], []);
  const context = loadUpdate(sandbox);

  context.update();

  assert.deepEqual(sandbox.calls, [
    ['deleteTrigger', 'updateGames'],
    ['deleteTrigger', 'updateRankings'],
    ['deleteTrigger', 'updateTitles'],
    ['deleteTrigger', 'updateRatings'],
    ['deleteTrigger', 'update'],
    ['setProperty', 'UPDATE_STEP', 'games'],
    ['ensureTrigger', 'update', 5],
    ['rankings'],
    ['ratings'],
  ]);
});

test('update ensures games resume state before sync imports that can hard-timeout', () => {
  const sandbox = createUpdateSandbox([], []);
  const context = loadUpdate(sandbox);

  context.update();

  const stepIndex = sandbox.calls.findIndex(
    (call) =>
      call[0] === 'setProperty' &&
      call[1] === 'UPDATE_STEP' &&
      call[2] === 'games',
  );
  const ensureIndex = sandbox.calls.findIndex(
    (call) => call[0] === 'ensureTrigger',
  );
  const rankingsIndex = sandbox.calls.findIndex(
    (call) => call[0] === 'rankings',
  );

  assert.ok(stepIndex !== -1, 'expected UPDATE_STEP=games to be persisted');
  assert.ok(ensureIndex !== -1, 'expected games trigger to be ensured');
  assert.ok(rankingsIndex !== -1, 'expected rankings import to run');
  assert.ok(
    stepIndex < rankingsIndex && ensureIndex < rankingsIndex,
    'Games STEP and trigger must be ready before sync imports that can hard-timeout',
  );
});

test('scheduled updates complete games before starting titles with the same handler', () => {
  const sandbox = createUpdateSandbox([true, false], [false]);
  const context = loadUpdate(sandbox);

  context.update();
  sandbox.calls.length = 0;

  context.update({ triggerUid: 'first' });
  assert.deepEqual(sandbox.calls, [['getProperty', 'UPDATE_STEP'], ['games']]);

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

test('update still schedules games when ranking import fails', () => {
  const sandbox = createUpdateSandbox([], [], true, {
    rankingError: new Error('BGA catalog unavailable'),
  });
  const context = loadUpdate(sandbox);

  context.update();

  assert.deepEqual(sandbox.calls, [
    ['deleteTrigger', 'updateGames'],
    ['deleteTrigger', 'updateRankings'],
    ['deleteTrigger', 'updateTitles'],
    ['deleteTrigger', 'updateRatings'],
    ['deleteTrigger', 'update'],
    ['setProperty', 'UPDATE_STEP', 'games'],
    ['ensureTrigger', 'update', 5],
    ['rankings'],
    ['ratings'],
  ]);
  assert.deepEqual(sandbox.logs, [
    'Ranking update failed: BGA catalog unavailable',
  ]);
});

test('update still schedules games when rating import fails', () => {
  const sandbox = createUpdateSandbox([], [], true, {
    ratingError: new Error('Bodoge HTML unrecognized'),
  });
  const context = loadUpdate(sandbox);

  context.update();

  assert.deepEqual(sandbox.calls, [
    ['deleteTrigger', 'updateGames'],
    ['deleteTrigger', 'updateRankings'],
    ['deleteTrigger', 'updateTitles'],
    ['deleteTrigger', 'updateRatings'],
    ['deleteTrigger', 'update'],
    ['setProperty', 'UPDATE_STEP', 'games'],
    ['ensureTrigger', 'update', 5],
    ['rankings'],
    ['ratings'],
  ]);
  assert.deepEqual(sandbox.logs, [
    'Rating update failed: Bodoge HTML unrecognized',
  ]);
});

test('update still schedules games when both sync imports fail', () => {
  const sandbox = createUpdateSandbox([], [], true, {
    rankingError: new Error('BGA catalog unavailable'),
    ratingError: new Error('Bodoge HTML unrecognized'),
  });
  const context = loadUpdate(sandbox);

  context.update();

  assert.deepEqual(sandbox.calls, [
    ['deleteTrigger', 'updateGames'],
    ['deleteTrigger', 'updateRankings'],
    ['deleteTrigger', 'updateTitles'],
    ['deleteTrigger', 'updateRatings'],
    ['deleteTrigger', 'update'],
    ['setProperty', 'UPDATE_STEP', 'games'],
    ['ensureTrigger', 'update', 5],
    ['rankings'],
    ['ratings'],
  ]);
  assert.deepEqual(sandbox.logs, [
    'Ranking update failed: BGA catalog unavailable',
    'Rating update failed: Bodoge HTML unrecognized',
  ]);
});
