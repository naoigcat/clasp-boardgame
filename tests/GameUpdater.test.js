/**
 * GameUpdater batch and refresh-policy tests.
 *
 * Focuses on empty sheets, oldest-first batching, derived-column clearing on
 * success and failure, advancing failure timestamps so permanent errors rotate
 * out, short failure backoff versus the seven-day success window, formula
 * escaping of external strings on write, and writing only the real 2–10
 * player-recommendation columns before rank.
 */
const assert = require('node:assert/strict');
const test = require('node:test');
const { createHash } = require('node:crypto');

const { createClockDate, loadScripts } = require('./helpers/appScriptHarness');

/** Zero-based Games value columns owned by spreadsheet array formulas. */
const GAME_ARRAY_FORMULA_INPUT_COLUMNS = [1, 4, 21, 22, 23];
/** Zero-based Games value column that stores the last update attempt. */
const GAME_LAST_UPDATED_AT_COLUMN = 24;
/** Zero-based Games value column that stores the last row-level error. */
const GAME_ERROR_MESSAGE_COLUMN = 25;
/** Zero-based Games value column for the first (2-player) recommendation. */
const GAME_PLAYER_RECOMMENDATION_START_COLUMN = 7;
/** Zero-based Games value column for the 10-player recommendation. */
const GAME_TEN_PLAYER_RECOMMENDATION_COLUMN = 15;
/** Zero-based Games value column for BoardGameGeek board-game rank. */
const GAME_BOARD_GAME_RANK_COLUMN = 16;

/**
 * Creates a Games sheet double whose first source-link cell is empty.
 */
function createEmptyGamesSheet() {
  const writes = [];
  const clears = [];
  const operations = [];
  const rangeReads = [];

  return {
    writes,
    clears,
    operations,
    rangeReads,
    getLastRow() {
      // Header only: no data rows to read.
      return 1;
    },
    getRange(a1NotationOrRow, column, numRows, numColumns) {
      if (typeof a1NotationOrRow === 'string') {
        throw new Error(`Unexpected open-ended range: ${a1NotationOrRow}`);
      }

      rangeReads.push({
        row: a1NotationOrRow,
        column,
        numRows,
        numColumns,
      });

      return {
        getRichTextValues() {
          return Array.from({ length: numRows }, () => [null]);
        },
        getValues() {
          return Array.from({ length: numRows }, () =>
            Array(numColumns).fill(''),
          );
        },
        setValues(values) {
          const write = {
            a1NotationOrRow,
            column,
            numRows,
            numColumns,
            values,
          };
          writes.push(write);
          operations.push({ type: 'setValues', ...write });
        },
        clearContent() {
          const clear = {
            row: a1NotationOrRow,
            column,
            numRows,
            numColumns,
          };
          clears.push(clear);
          operations.push({ type: 'clearContent', ...clear });
        },
      };
    },
  };
}

/**
 * Gives each derived cell a unique value so accidental resets are observable.
 */
function createGameRow(index, url) {
  const values = Array(26).fill('');
  GAME_ARRAY_FORMULA_INPUT_COLUMNS.forEach((column) => {
    values[column] = `derived-${index}-${column}`;
  });
  values[GAME_LAST_UPDATED_AT_COLUMN] = new Date(2020, 0, index + 1);

  return {
    rowNumber: index + 2,
    gameLink: {
      getText() {
        return `Game ${index}`;
      },
      getLinkUrl() {
        return url;
      },
    },
    values,
  };
}

/**
 * Keeps the double limited to managed rows so the batch boundary is unambiguous.
 *
 * `surplusRowCount` simulates abandoned C–AB cells below the first blank link.
 * `failOnSetValues` simulates a failed matched-row write.
 */
function createGamesSheet(
  rows,
  { surplusRowCount = 0, failOnSetValues = false } = {},
) {
  const writes = [];
  const clears = [];
  const operations = [];
  const rangeReads = [];

  return {
    writes,
    clears,
    operations,
    rangeReads,
    getLastRow() {
      // Header plus managed rows, plus any abandoned physical rows left below
      // the blank-link end marker after a shortened Games list.
      return rows.length + 1 + surplusRowCount;
    },
    getRange(a1NotationOrRow, column, numRows, numColumns) {
      if (typeof a1NotationOrRow === 'string') {
        throw new Error(`Unexpected open-ended range: ${a1NotationOrRow}`);
      }

      const range = {
        getRichTextValues() {
          rangeReads.push({
            row: a1NotationOrRow,
            column,
            numRows,
            numColumns,
          });
          // Surplus rows below the managed block have empty links; the updater
          // stops at the first blank, so only managed rows are returned as data.
          return Array.from({ length: numRows }, (_, index) => [
            index + a1NotationOrRow - 2 < rows.length
              ? rows[index + a1NotationOrRow - 2].gameLink
              : null,
          ]);
        },
        getValues() {
          rangeReads.push({
            row: a1NotationOrRow,
            column,
            numRows,
            numColumns,
          });
          return Array.from({ length: numRows }, (_, index) => {
            const row = rows[index + a1NotationOrRow - 2];
            return column === 1
              ? [row?.rowNumber ?? '']
              : row
                ? row.values.slice()
                : Array(numColumns).fill('stale');
          });
        },
        setValues(values) {
          if (failOnSetValues) {
            throw new Error('setValues failed');
          }
          const write = {
            a1NotationOrRow,
            column,
            numRows,
            numColumns,
            values,
          };
          writes.push(write);
          operations.push({ type: 'setValues', ...write });
        },
        clearContent() {
          const clear = {
            row: a1NotationOrRow,
            column,
            numRows,
            numColumns,
          };
          clears.push(clear);
          operations.push({ type: 'clearContent', ...clear });
        },
      };

      return range;
    },
  };
}

/** Reconstructs the managed value columns from bounded range writes. */
function writtenGameRows(sheet) {
  const rows = new Map();
  for (const write of sheet.writes) {
    const row = rows.get(write.a1NotationOrRow) ?? Array(26).fill('');
    row.splice(write.column - 3, write.numColumns, ...write.values[0]);
    rows.set(write.a1NotationOrRow, row);
  }
  return [...rows.entries()]
    .sort(([first], [second]) => first - second)
    .map(([, values]) => values);
}

/**
 * Keeps assertions aligned with the formula-owned columns of a refreshed row.
 */
function getArrayFormulaInputs(values) {
  return GAME_ARRAY_FORMULA_INPUT_COLUMNS.map((column) => values[column]);
}

/**
 * Builds a minimal XmlService-like attribute for BoardGameGeek fixtures.
 */
function createXmlAttribute(value) {
  return {
    getValue() {
      return value;
    },
  };
}

/**
 * Builds a minimal XmlService-like element for BoardGameGeek fixtures.
 */
function createXmlElement({ attributes = {}, child = {}, children = {} } = {}) {
  return {
    getAttribute(name) {
      return Object.prototype.hasOwnProperty.call(attributes, name)
        ? createXmlAttribute(attributes[name])
        : null;
    },
    getChild(name) {
      return Object.prototype.hasOwnProperty.call(child, name)
        ? child[name]
        : null;
    },
    getChildren(name) {
      return children[name] ?? [];
    },
  };
}

/**
 * Builds one poll results element for a player count.
 */
function createPlayerCountResults(playerCount, recommendation) {
  return createXmlElement({
    attributes: { numplayers: playerCount },
    children: {
      result: [
        createXmlElement({
          attributes: { value: recommendation, numvotes: '7' },
        }),
        createXmlElement({
          attributes: { value: 'Not Recommended', numvotes: '1' },
        }),
      ],
    },
  });
}

/**
 * Returns a BoardGameGeek item with 10- and 11-player poll entries so tests can
 * assert that only the real 2–10 recommendation columns are written.
 */
function createGameItemWithTenAndElevenPlayerRecommendations() {
  return createXmlElement({
    children: {
      poll: [
        createXmlElement({
          attributes: { name: 'suggested_numplayers' },
          children: {
            results: [
              createPlayerCountResults('10', 'Best'),
              createPlayerCountResults('11', 'Recommended'),
            ],
          },
        }),
      ],
    },
    child: {
      statistics: createXmlElement({
        child: {
          ratings: createXmlElement({
            child: {
              ranks: createXmlElement({
                children: {
                  rank: [
                    createXmlElement({
                      attributes: { name: 'boardgame', value: '42' },
                    }),
                  ],
                },
              }),
              bayesaverage: createXmlElement({
                attributes: { value: '7.5' },
              }),
              averageweight: createXmlElement({
                attributes: { value: '2.1' },
              }),
            },
          }),
        },
      }),
      minplaytime: createXmlElement({ attributes: { value: '30' } }),
      maxplaytime: createXmlElement({ attributes: { value: '45' } }),
      yearpublished: createXmlElement({ attributes: { value: '2020' } }),
    },
  });
}

/**
 * Supplies persistent queue properties and the Apps Script hashing API.
 */
function addGameQueueServices(sandbox) {
  const properties = new Map();
  const getProperties = sandbox.PropertiesService?.getScriptProperties;
  sandbox.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => properties.get(key) ?? null,
      setProperty: (key, value) => properties.set(key, value),
      deleteProperty: (key) => properties.delete(key),
      ...getProperties?.(),
    }),
  };
  sandbox.Utilities = {
    DigestAlgorithm: { SHA_256: 'sha256' },
    computeDigest: (_algorithm, value) => [
      ...createHash('sha256').update(value).digest(),
    ],
    ...sandbox.Utilities,
  };
}

function loadGameUpdater(sandbox) {
  addGameQueueServices(sandbox);
  return loadScripts(sandbox, [
    { path: 'src/config/AppConfig.ts', exports: [] },
    { path: 'src/config/SheetSchema.ts', exports: [] },
    { path: 'src/config/TitleRules.ts', exports: [] },
    { path: 'src/shared/DateUtils.ts', exports: [] },
    { path: 'src/shared/ErrorUtils.ts', exports: [] },
    { path: 'src/shared/SpreadsheetUtils.ts', exports: [] },
    { path: 'src/shared/XmlUtils.ts', exports: [] },
    { path: 'src/infrastructure/SpreadsheetGateway.ts', exports: [] },
    { path: 'src/infrastructure/ScriptPropertyStore.ts', exports: [] },
    { path: 'src/services/GameUpdater.ts', exports: ['GameUpdater'] },
  ]);
}

/**
 * Loads GameUpdater with the HTTP client path used for BoardGameGeek fetches.
 */
function loadGameUpdaterWithHttp(sandbox) {
  addGameQueueServices(sandbox);
  return loadScripts(sandbox, [
    { path: 'src/config/AppConfig.ts', exports: [] },
    { path: 'src/config/SheetSchema.ts', exports: [] },
    { path: 'src/config/TitleRules.ts', exports: [] },
    { path: 'src/shared/DateUtils.ts', exports: [] },
    { path: 'src/shared/ErrorUtils.ts', exports: [] },
    { path: 'src/shared/SpreadsheetUtils.ts', exports: [] },
    { path: 'src/shared/XmlUtils.ts', exports: [] },
    { path: 'src/infrastructure/HttpClient.ts', exports: ['HttpClient'] },
    {
      path: 'src/infrastructure/ScriptPropertyStore.ts',
      exports: ['ScriptPropertyStore'],
    },
    { path: 'src/infrastructure/SpreadsheetGateway.ts', exports: [] },
    { path: 'src/services/GameUpdater.ts', exports: ['GameUpdater'] },
  ]);
}

test('GameUpdater keeps the Games phase when the Games sheet is missing', () => {
  const logs = [];
  const context = loadGameUpdater({
    Logger: {
      log(message) {
        logs.push(message);
      },
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName() {
            return null;
          },
        };
      },
    },
  });

  // false would let UpdateCoordinator advance to Titles and treat a renamed or
  // deleted Games tab as a completed BoardGameGeek batch with no execution log.
  assert.equal(context.GameUpdater.run(), true);
  assert.deepEqual(logs, [
    'Games sheet "Games" is missing; keeping the Games phase until the tab is restored.',
  ]);
});

test('GameUpdater treats a null rich-text value as the end of the Games data', () => {
  const gamesSheet = createEmptyGamesSheet();
  const context = loadGameUpdater({
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  assert.equal(context.GameUpdater.run(), false);
  assert.deepEqual(gamesSheet.writes, []);
  // Header-only sheets must not issue a data-range read at all.
  assert.deepEqual(gamesSheet.rangeReads, []);
});

test('GameUpdater bounds Games sheet reads to getLastRow instead of open-ended A1 ranges', () => {
  const rows = [
    createGameRow(0, 'https://boardgamegeek.com/boardgame/1'),
    createGameRow(1, 'https://boardgamegeek.com/boardgame/2'),
  ];
  const gamesSheet = createGamesSheet(rows);
  // Simulate a tall sheet whose last content row is only the two managed games;
  // open-ended `$A$2:$A` would otherwise scan to the sheet maximum.
  let reportedLastRow = 3;
  gamesSheet.getLastRow = () => reportedLastRow;
  const context = loadGameUpdater({
    Date,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  context.GameUpdater.fetchGameItem = () => ({});
  context.GameUpdater.applyGameItem = (row, _gameItem, _gameId, current) => {
    row.values[GAME_LAST_UPDATED_AT_COLUMN] = current;
  };

  assert.equal(context.GameUpdater.run(), false);

  assert.deepEqual(gamesSheet.rangeReads.slice(0, 3), [
    { row: 2, column: 1, numRows: 2, numColumns: 1 },
    { row: 2, column: 2, numRows: 2, numColumns: 1 },
    { row: 2, column: 3, numRows: 2, numColumns: 26 },
  ]);
  assert.ok(
    gamesSheet.rangeReads.every((range) => range.row + range.numRows - 1 <= 3),
  );
  assert.equal(writtenGameRows(gamesSheet).length, 2);
  // A regression that reintroduced `$A$2:$A` would throw in createGamesSheet.
  reportedLastRow = 1;
  gamesSheet.rangeReads.length = 0;
  gamesSheet.writes.length = 0;
  assert.equal(context.GameUpdater.run(), false);
  assert.deepEqual(gamesSheet.rangeReads, []);
});

test('GameUpdater writes matching Games rows without clearing surplus rows', () => {
  const rows = [
    createGameRow(0, 'https://boardgamegeek.com/boardgame/1'),
    createGameRow(1, 'https://boardgamegeek.com/boardgame/2'),
  ];
  // Two abandoned physical rows remain below the blank-link end marker after a
  // shortened list; unmatched C–AB cells must remain untouched.
  const gamesSheet = createGamesSheet(rows, { surplusRowCount: 2 });
  const context = loadGameUpdater({
    Date,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  context.GameUpdater.fetchGameItem = () => ({});
  context.GameUpdater.applyGameItem = (row, _gameItem, _gameId, current) => {
    row.values[GAME_LAST_UPDATED_AT_COLUMN] = current;
  };

  assert.equal(context.GameUpdater.run(), false);

  // Only matched rows are written; surplus rows may belong to newly added games.
  assert.deepEqual(
    gamesSheet.operations.map((operation) => operation.type),
    Array(6).fill('setValues'),
  );
  assert.deepEqual(gamesSheet.clears, []);
  assert.equal(writtenGameRows(gamesSheet).length, 2);
  assert.equal(gamesSheet.writes[0].column, 4);
});

test('GameUpdater preserves surplus rows when every managed game is still fresh', () => {
  const rows = [
    createGameRow(0, 'https://boardgamegeek.com/boardgame/1'),
    createGameRow(1, 'https://boardgamegeek.com/boardgame/2'),
  ];
  // Within the seven-day refresh window so no game needs a fetch.
  rows.forEach((row) => {
    row.values[GAME_LAST_UPDATED_AT_COLUMN] = new Date(2024, 5, 14);
  });
  const gamesSheet = createGamesSheet(rows, { surplusRowCount: 2 });
  const { ClockDate } = createClockDate(Date.UTC(2024, 5, 15));
  let fetchCount = 0;
  const context = loadGameUpdater({
    Date: ClockDate,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  context.GameUpdater.fetchGameItem = () => {
    fetchCount += 1;
    return {};
  };

  assert.equal(context.GameUpdater.run(), false);
  assert.equal(fetchCount, 0);
  // Fresh rows clear formula inputs without touching any unmatched rows.
  assert.deepEqual(
    gamesSheet.operations.map((operation) => operation.type),
    Array(6).fill('setValues'),
  );
  assert.deepEqual(gamesSheet.clears, []);
  assert.equal(writtenGameRows(gamesSheet).length, 2);
  assert.equal(gamesSheet.writes[0].column, 4);
  assert.deepEqual(
    getArrayFormulaInputs(writtenGameRows(gamesSheet)[0]),
    Array(5).fill(null),
  );
});

test('GameUpdater preserves data rows when the managed Games list is empty', () => {
  // First blank link ends the managed block immediately, but abandoned C–AB
  // cells can remain after every column-B link is deleted.
  const gamesSheet = createGamesSheet([], { surplusRowCount: 3 });
  let fetchCount = 0;
  const context = loadGameUpdater({
    Date,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  context.GameUpdater.fetchGameItem = () => {
    fetchCount += 1;
    return {};
  };

  assert.equal(context.GameUpdater.run(), false);
  assert.equal(fetchCount, 0);
  // No identity matches means no values may be written or cleared.
  assert.deepEqual(
    gamesSheet.operations.map((operation) => operation.type),
    [],
  );
  assert.deepEqual(gamesSheet.clears, []);
  assert.deepEqual(gamesSheet.writes, []);
});

test('GameUpdater leaves Games intact when the first matched-row write fails', () => {
  const rows = [createGameRow(0, 'https://boardgamegeek.com/boardgame/1')];
  const gamesSheet = createGamesSheet(rows, {
    surplusRowCount: 2,
    failOnSetValues: true,
  });
  const context = loadGameUpdater({
    Date,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  context.GameUpdater.fetchGameItem = () => ({});
  context.GameUpdater.applyGameItem = (row, _gameItem, _gameId, current) => {
    row.values[GAME_LAST_UPDATED_AT_COLUMN] = current;
  };

  assert.throws(() => context.GameUpdater.run(), /setValues failed/);

  // Clearing must not run after a write failure; otherwise managed C–AB values
  // would disappear even though BoardGameGeek cannot restore the prior snapshot.
  assert.deepEqual(gamesSheet.operations, []);
  assert.deepEqual(gamesSheet.clears, []);
  assert.deepEqual(gamesSheet.writes, []);
});

test('GameUpdater defers untouched rows at the runtime boundary and clears them next time', () => {
  const rows = Array.from({ length: 51 }, (_, index) =>
    createGameRow(index, `https://boardgamegeek.com/boardgame/${index + 1}`),
  );
  const gamesSheet = createGamesSheet(rows);
  // Start after the fixture last-updated dates so rows stay stale while Date.now
  // advances only for the soft runtime budget.
  const { ClockDate, clock } = createClockDate(Date.UTC(2024, 0, 1));
  const context = loadGameUpdater({
    Date: ClockDate,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  // Fifty successful fetches exhaust the soft 180s budget before the 51st row.
  const runtimeStepMs =
    context.UPDATE_QUEUE_CONFIG.MAX_RUNTIME_MILLISECONDS / 50;
  context.GameUpdater.fetchGameItem = () => {
    clock.nowMs += runtimeStepMs;
    return {};
  };
  context.GameUpdater.applyGameItem = (row, _gameItem, _gameId, current) => {
    row.values[GAME_LAST_UPDATED_AT_COLUMN] = current;
  };

  assert.equal(context.GameUpdater.run(), true);

  const writtenRows = writtenGameRows(gamesSheet);
  // null clears the cell for ARRAYFORMULA; '' would leave a blocking blank.
  assert.deepEqual(getArrayFormulaInputs(writtenRows[0]), Array(5).fill(null));
  assert.equal(writtenRows.length, 50);
  writtenRows.forEach((values, index) => {
    rows[index].values = values;
  });
  gamesSheet.writes.length = 0;
  assert.equal(context.GameUpdater.run(), false);
  assert.deepEqual(
    getArrayFormulaInputs(writtenGameRows(gamesSheet)[50]),
    Array(5).fill(null),
  );
});

test('GameUpdater clears formula inputs when a game refresh fails', () => {
  const row = createGameRow(0, 'not-a-boardgamegeek-url');
  // Uncalculated derived cells often hold ''; writing them back would block
  // ARRAYFORMULA for the full refresh window after the failed attempt.
  GAME_ARRAY_FORMULA_INPUT_COLUMNS.forEach((column) => {
    row.values[column] = '';
  });
  const gamesSheet = createGamesSheet([row]);
  const context = loadGameUpdater({
    Date,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  assert.equal(context.GameUpdater.run(), false);

  assert.deepEqual(
    getArrayFormulaInputs(writtenGameRows(gamesSheet)[0]),
    Array(5).fill(null),
  );
  assert.ok(
    writtenGameRows(gamesSheet)[0][GAME_LAST_UPDATED_AT_COLUMN] instanceof Date,
  );
});

test('GameUpdater retries failed rows after the short failure backoff, not the seven-day success window', () => {
  const row = createGameRow(0, 'https://boardgamegeek.com/boardgame/42');
  const rows = [row];
  const gamesSheet = createGamesSheet(rows);
  const failureAt = new Date(2024, 0, 1, 12, 0, 0);
  const { ClockDate, clock } = createClockDate(failureAt.getTime());
  let fetchCount = 0;
  const context = loadGameUpdater({
    Date: ClockDate,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  context.GameUpdater.fetchGameItem = () => {
    fetchCount += 1;
    throw new Error('HTTP 503');
  };

  assert.equal(context.GameUpdater.run(), false);
  assert.equal(fetchCount, 1);
  assert.match(
    writtenGameRows(gamesSheet)[0][GAME_ERROR_MESSAGE_COLUMN],
    /^HTTP 503$/,
  );

  // Persist the failure stamp the same way a later trigger rereads the sheet.
  writtenGameRows(gamesSheet).forEach((values, index) => {
    rows[index].values = values;
  });
  gamesSheet.writes.length = 0;

  // Still inside the one-day failure backoff: must not re-fetch yet.
  clock.nowMs = failureAt.getTime() + 12 * 60 * 60 * 1000;
  assert.equal(context.GameUpdater.run(), false);
  assert.equal(fetchCount, 1);

  // Past the failure backoff but well inside the seven-day success window:
  // treating failures like successes would skip this retry until day seven.
  clock.nowMs =
    failureAt.getTime() +
    context.UPDATE_QUEUE_CONFIG.GAME_FAILURE_BACKOFF_DAYS * 24 * 60 * 60 * 1000;
  assert.equal(context.GameUpdater.run(), false);
  assert.equal(fetchCount, 2);
});

test('GameUpdater keeps successful rows ineligible for the full seven-day refresh window', () => {
  const row = createGameRow(0, 'https://boardgamegeek.com/boardgame/42');
  const rows = [row];
  const gamesSheet = createGamesSheet(rows);
  const successAt = new Date(2024, 0, 1, 12, 0, 0);
  const { ClockDate, clock } = createClockDate(successAt.getTime());
  let fetchCount = 0;
  const context = loadGameUpdater({
    Date: ClockDate,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  context.GameUpdater.fetchGameItem = () => {
    fetchCount += 1;
    return {};
  };
  context.GameUpdater.applyGameItem = (
    gameRow,
    _gameItem,
    _gameId,
    current,
  ) => {
    gameRow.values[GAME_LAST_UPDATED_AT_COLUMN] = current;
    gameRow.values[GAME_ERROR_MESSAGE_COLUMN] = '';
  };

  assert.equal(context.GameUpdater.run(), false);
  assert.equal(fetchCount, 1);

  writtenGameRows(gamesSheet).forEach((values, index) => {
    rows[index].values = values;
  });
  gamesSheet.writes.length = 0;

  // One day is enough for a failed row to retry, but a success must still wait.
  clock.nowMs =
    successAt.getTime() +
    context.UPDATE_QUEUE_CONFIG.GAME_FAILURE_BACKOFF_DAYS * 24 * 60 * 60 * 1000;
  assert.equal(context.GameUpdater.run(), false);
  assert.equal(fetchCount, 1);

  clock.nowMs =
    successAt.getTime() +
    context.UPDATE_QUEUE_CONFIG.GAME_REFRESH_INTERVAL_DAYS *
      24 *
      60 *
      60 *
      1000;
  assert.equal(context.GameUpdater.run(), false);
  assert.equal(fetchCount, 2);
});

test('GameUpdater advances past permanently failing head rows on the next batch', () => {
  const failingRows = Array.from({ length: 50 }, (_, index) =>
    createGameRow(index, `https://boardgamegeek.com/boardgame/${index + 1000}`),
  );
  const successRow = createGameRow(
    50,
    'https://boardgamegeek.com/boardgame/42',
  );
  const rows = [...failingRows, successRow];
  const gamesSheet = createGamesSheet(rows);
  const batchStartedAtMs = Date.UTC(2024, 0, 1);
  const { ClockDate, clock } = createClockDate(batchStartedAtMs);
  const context = loadGameUpdater({
    Date: ClockDate,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  // Each failed head row spends an equal share of the soft budget so the first
  // invocation stops before the still-stale success row is reached.
  const runtimeStepMs =
    context.UPDATE_QUEUE_CONFIG.MAX_RUNTIME_MILLISECONDS / 50;
  context.GameUpdater.fetchGameItem = (gameReference) => {
    clock.nowMs += runtimeStepMs;
    if (gameReference.id === '42') {
      return {};
    }

    throw new Error('permanent upstream failure');
  };
  context.GameUpdater.applyGameItem = (row, _gameItem, _gameId, current) => {
    row.values[GAME_LAST_UPDATED_AT_COLUMN] = current;
  };

  assert.equal(context.GameUpdater.run(), true);
  assert.equal(writtenGameRows(gamesSheet).length, 50);

  // The sheet double reads from the source row objects, so persist the first
  // batch write before asserting that the next batch can move past failures.
  writtenGameRows(gamesSheet).forEach((values, index) => {
    rows[index].values = values;
  });
  gamesSheet.writes.length = 0;
  clock.nowMs = batchStartedAtMs;

  assert.equal(context.GameUpdater.run(), false);
  assert.deepEqual(
    getArrayFormulaInputs(writtenGameRows(gamesSheet)[50]),
    Array(5).fill(null),
  );
});

test('GameUpdater escapes formula-like recommendations and errors before setValues', () => {
  const formulaRecommendation = '=HYPERLINK("https://evil.example","x")';
  const formulaError = '+cmd|a';
  const successRow = createGameRow(0, 'https://boardgamegeek.com/boardgame/1');
  const failureRow = createGameRow(1, 'https://boardgamegeek.com/boardgame/2');
  const gamesSheet = createGamesSheet([successRow, failureRow]);
  const context = loadGameUpdater({
    Date,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  context.GameUpdater.fetchGameItem = (gameReference) => {
    if (gameReference.id === '2') {
      throw new Error(formulaError);
    }

    return {};
  };
  context.GameUpdater.applyGameItem = (row, _gameItem, _gameId, current) => {
    row.values[GAME_PLAYER_RECOMMENDATION_START_COLUMN] = formulaRecommendation;
    row.values[GAME_LAST_UPDATED_AT_COLUMN] = current;
    row.values[GAME_ERROR_MESSAGE_COLUMN] = '';
  };

  assert.equal(context.GameUpdater.run(), false);

  const [writtenSuccess, writtenFailure] = writtenGameRows(gamesSheet);
  // Recommendation labels and error text are external; a leading apostrophe
  // keeps Sheets from executing them as formulas when the workbook is shared.
  assert.equal(
    writtenSuccess[GAME_PLAYER_RECOMMENDATION_START_COLUMN],
    `'${formulaRecommendation}`,
  );
  assert.equal(writtenFailure[GAME_ERROR_MESSAGE_COLUMN], `'${formulaError}`);
  // Cleared error cells stay empty rather than becoming a quoted blank.
  assert.equal(writtenSuccess[GAME_ERROR_MESSAGE_COLUMN], '');
});

test('GameUpdater writes 2–10 player recommendations and leaves board-game rank intact', () => {
  const context = loadGameUpdater({});
  const row = {
    gameLink: null,
    values: Array(26).fill(''),
  };

  context.GameUpdater.applyGameItem(
    row,
    createGameItemWithTenAndElevenPlayerRecommendations(),
    '9999',
    new Date(2024, 0, 1),
  );

  assert.equal(row.values[GAME_TEN_PLAYER_RECOMMENDATION_COLUMN], 'Best');
  assert.equal(row.values[GAME_BOARD_GAME_RANK_COLUMN], 42);
  assert.notEqual(row.values[GAME_BOARD_GAME_RANK_COLUMN], 'Recommended');
});

test('GameUpdater keeps prior player recommendations when applyGameItem fails mid-parse', () => {
  const row = createGameRow(0, 'https://boardgamegeek.com/boardgame/42');
  row.values[GAME_PLAYER_RECOMMENDATION_START_COLUMN] = 'Best';
  row.values[GAME_TEN_PLAYER_RECOMMENDATION_COLUMN] = 'Recommended';
  row.values[GAME_BOARD_GAME_RANK_COLUMN] = 99;
  const priorRecommendations = row.values.slice(
    GAME_PLAYER_RECOMMENDATION_START_COLUMN,
    GAME_TEN_PLAYER_RECOMMENDATION_COLUMN + 1,
  );
  const priorFormulaInputs = getArrayFormulaInputs(row.values);
  const gamesSheet = createGamesSheet([row]);
  const context = loadGameUpdater({
    Date,
    Logger: {
      log() {},
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  // Poll data is present so recommendations would be rewritten if applied before
  // statistics parsing; omitting statistics forces the mid-apply failure path.
  context.GameUpdater.fetchGameItem = () =>
    createXmlElement({
      children: {
        poll: [
          createXmlElement({
            attributes: { name: 'suggested_numplayers' },
            children: {
              results: [
                createPlayerCountResults('2', 'Not Recommended'),
                createPlayerCountResults('10', 'Not Recommended'),
              ],
            },
          }),
        ],
      },
    });

  assert.equal(context.GameUpdater.run(), false);

  const written = writtenGameRows(gamesSheet)[0];
  // Copy into a host-realm array: sheet doubles hold VM arrays from loadScripts.
  assert.deepEqual(
    Array.from(
      written.slice(
        GAME_PLAYER_RECOMMENDATION_START_COLUMN,
        GAME_TEN_PLAYER_RECOMMENDATION_COLUMN + 1,
      ),
    ),
    priorRecommendations,
  );
  assert.equal(written[GAME_BOARD_GAME_RANK_COLUMN], 99);
  assert.match(
    written[GAME_ERROR_MESSAGE_COLUMN],
    /Required XML child "statistics" was not found/,
  );
  assert.ok(written[GAME_LAST_UPDATED_AT_COLUMN] instanceof Date);
  // Derived columns are sheet-owned; failure must still null them even though
  // BoardGameGeek metadata from before the attempt is preserved.
  assert.notDeepEqual(priorFormulaInputs, Array(5).fill(null));
  assert.deepEqual(getArrayFormulaInputs(written), Array(5).fill(null));
});

test('GameUpdater throttles before raising a non-2xx BoardGameGeek response', () => {
  const row = createGameRow(0, 'https://boardgamegeek.com/boardgame/42');
  const gamesSheet = createGamesSheet([row]);
  const sleepCalls = [];
  const context = loadGameUpdaterWithHttp({
    Date,
    Logger: {
      log() {},
    },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty() {
            return null;
          },
        };
      },
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
    UrlFetchApp: {
      fetch(url, options = {}) {
        if (options.muteHttpExceptions !== true) {
          throw new Error(`Request failed for ${url}: 503`);
        }

        return {
          getResponseCode() {
            return 503;
          },
          getContentText() {
            return 'unavailable';
          },
        };
      },
    },
    Utilities: {
      sleep(milliseconds) {
        sleepCalls.push(milliseconds);
      },
    },
  });

  assert.equal(context.GameUpdater.run(), false);

  assert.deepEqual(sleepCalls, [
    context.BOARD_GAME_GEEK_CONFIG.REQUEST_DELAY_MILLISECONDS,
  ]);
  assert.match(
    writtenGameRows(gamesSheet)[0][GAME_ERROR_MESSAGE_COLUMN],
    /^HTTP 503$/,
  );
  assert.deepEqual(
    getArrayFormulaInputs(writtenGameRows(gamesSheet)[0]),
    Array(5).fill(null),
  );
  assert.ok(
    writtenGameRows(gamesSheet)[0][GAME_LAST_UPDATED_AT_COLUMN] instanceof Date,
  );
});

test('GameUpdater matches row numbers after edits during metadata fetching', async (t) => {
  const scenarios = [
    {
      name: 'sorting moves the update with its row number',
      edit: (rows) => rows.reverse(),
      expected: [
        [2, 3],
        [3, 2],
      ],
      pending: false,
    },
    {
      name: 'inserting a game preserves the new row',
      edit: (rows) =>
        rows.unshift(createGameRow(8, 'https://boardgamegeek.com/boardgame/9')),
      expected: [
        [2, 3],
        [3, 4],
      ],
      pending: false,
    },
    {
      name: 'deleting a game never writes its values onto the next row',
      edit: (rows) => rows.shift(),
      expected: [[3, 2]],
      pending: true,
    },
    {
      name: 'duplicate row numbers skip both targets',
      edit: (rows) => {
        rows[1].rowNumber = rows[0].rowNumber;
      },
      expected: [],
      pending: true,
    },
    {
      name: 'blank row number is left untouched',
      edit: (rows) => {
        rows[0].rowNumber = '';
      },
      expected: [[3, 3]],
      pending: true,
    },
    {
      name: 'changed link with the same row number clears only its old timestamp',
      edit: (rows) => {
        rows[0].gameLink = createGameRow(
          8,
          'https://boardgamegeek.com/boardgame/9',
        ).gameLink;
      },
      expected: [[3, 3]],
      pending: true,
      clears: [{ row: 2, column: 27, numRows: 1, numColumns: 1 }],
    },
  ];
  for (const scenario of scenarios) {
    await t.test(scenario.name, () => {
      const rows = [
        createGameRow(0, 'https://boardgamegeek.com/boardgame/1'),
        createGameRow(1, 'https://boardgamegeek.com/boardgame/2'),
      ];
      const sheet = createGamesSheet(rows);
      const context = loadGameUpdater({
        Date,
        Logger: { log() {} },
        SpreadsheetApp: {
          getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }),
        },
      });
      let edited = false;
      context.GameUpdater.fetchGameItem = () => {
        if (!edited) {
          scenario.edit(rows);
          edited = true;
        }
        return {};
      };
      context.GameUpdater.applyGameItem = (row, _item, _id, current) => {
        row.values[GAME_BOARD_GAME_RANK_COLUMN] = row.rowNumber;
        row.values[GAME_LAST_UPDATED_AT_COLUMN] = current;
      };
      assert.equal(context.GameUpdater.run(), scenario.pending);
      assert.deepEqual(
        sheet.writes
          .filter((write) => write.column === 10)
          .map((write) => [
            write.values[0][GAME_BOARD_GAME_RANK_COLUMN - 7],
            write.a1NotationOrRow,
          ]),
        scenario.expected,
      );
      assert.ok(
        sheet.writes.every(
          (write) =>
            [4, 7, 10].includes(write.column) &&
            write.numRows === 1 &&
            write.numColumns === (write.column === 10 ? 19 : 1),
        ),
      );
      assert.deepEqual(sheet.clears, scenario.clears ?? []);
    });
  }
});

test('GameUpdater skips duplicate numbers already present when the batch starts', () => {
  const rows = [
    createGameRow(0, 'https://boardgamegeek.com/boardgame/1'),
    createGameRow(1, 'https://boardgamegeek.com/boardgame/2'),
  ];
  rows[1].rowNumber = rows[0].rowNumber;
  const sheet = createGamesSheet(rows);
  const context = loadGameUpdater({
    Date,
    Logger: { log() {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }),
    },
  });
  assert.equal(
    context.GameUpdater.writeRows(sheet, context.GameUpdater.loadRows(sheet)),
    true,
  );
  assert.deepEqual(sheet.writes, []);
});

test('GameUpdater accepts decimal row numbers stored as text', () => {
  const ids = ['2', '5', '6', '8', '9', '13'];
  const rows = ids.map((id, index) => ({
    ...createGameRow(index, `https://boardgamegeek.com/boardgame/${index + 1}`),
    rowNumber: id,
  }));
  const sheet = createGamesSheet(rows);
  const logs = [];
  const context = loadGameUpdater({
    Date,
    Logger: { log: (message) => logs.push(message) },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }),
    },
  });
  let fetches = 0;
  context.GameUpdater.fetchGameItem = () => {
    fetches += 1;
    return {};
  };
  context.GameUpdater.applyGameItem = (row, _item, _id, current) => {
    row.values[GAME_LAST_UPDATED_AT_COLUMN] = current;
  };
  assert.equal(context.GameUpdater.run(), false);
  assert.equal(fetches, ids.length);
  assert.equal(sheet.writes.length, ids.length * 3);
  assert.deepEqual(
    rows.map((row) => row.rowNumber),
    ids,
  );
  assert.deepEqual(logs, []);
});

test('GameUpdater excludes permanent identity errors without fetching or blocking later batches', () => {
  for (const fresh of [false, true]) {
    // Numeric and text forms of the same ID must also count as duplicates.
    const ids = ['', 'invalid', false, 1, 2.5, '9007199254740992', 5, '5', 13];
    const rows = ids.map((id, index) => ({
      ...createGameRow(
        index,
        `https://boardgamegeek.com/boardgame/${index + 1}`,
      ),
      rowNumber: id,
    }));
    if (fresh) {
      rows.forEach((row) => {
        row.values[GAME_LAST_UPDATED_AT_COLUMN] = new Date();
      });
    }
    const sheet = createGamesSheet(rows);
    const logs = [];
    const context = loadGameUpdater({
      Date,
      Logger: { log: (message) => logs.push(message) },
      SpreadsheetApp: {
        getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }),
      },
    });
    let fetches = 0;
    context.GameUpdater.fetchGameItem = () => {
      fetches += 1;
      return {};
    };
    context.GameUpdater.applyGameItem = (row, _item, _id, current) => {
      row.values[GAME_LAST_UPDATED_AT_COLUMN] = current;
    };
    for (let batch = 0; batch < 3; batch += 1) {
      assert.equal(context.GameUpdater.run(), false);
      // Persist the valid row only, as the next trigger rereads the sheet.
      rows[8].values = writtenGameRows(sheet).at(-1);
    }
    assert.equal(fetches, fresh ? 0 : 1);
    assert.equal(sheet.writes.length, 9);
    assert.ok(sheet.writes.every((write) => write.a1NotationOrRow === 10));
    assert.equal(logs.length, 8 * 3);
    assert.ok(
      logs.every((message) => message.includes('excluded from this update')),
    );

    // Repairing an excluded row makes it eligible in the next invocation.
    rows[0].rowNumber = 20;
    rows[0].values[GAME_LAST_UPDATED_AT_COLUMN] = new Date(2020, 0, 1);
    assert.equal(context.GameUpdater.run(), false);
    assert.equal(fetches, fresh ? 1 : 2);
    assert.ok(sheet.writes.some((write) => write.a1NotationOrRow === 2));

    // Exercise the real coordinator through Titles and trigger cleanup.
    const properties = context.PropertiesService.getScriptProperties();
    properties.setProperty('UPDATE_STEP', 'games');
    let titleRuns = 0;
    let removedTriggers = 0;
    context.TitleUpdater = {
      run: () => {
        titleRuns += 1;
        return false;
      },
    };
    context.TriggerManager = {
      removeAll: () => {
        removedTriggers += 1;
      },
    };
    loadScripts(context, [
      {
        path: 'src/services/UpdateCoordinator.ts',
        exports: ['UpdateCoordinator'],
      },
    ]);
    context.UpdateCoordinator.resumePendingPhase();
    assert.equal(properties.getProperty('UPDATE_STEP'), 'titles');
    context.UpdateCoordinator.resumePendingPhase();
    assert.equal(properties.getProperty('UPDATE_STEP'), null);
    assert.equal(titleRuns, 1);
    assert.equal(removedTriggers, 1);

    // A sheet containing only excluded rows must also finish immediately.
    rows[0].rowNumber = '';
    rows[8].rowNumber = 5;
    sheet.writes.length = 0;
    const previousFetches = fetches;
    assert.equal(context.GameUpdater.run(), false);
    assert.equal(fetches, previousFetches);
    assert.deepEqual(sheet.writes, []);
  }
});

test('GameUpdater checks the destination again immediately before each write', () => {
  const rows = [
    createGameRow(0, 'https://boardgamegeek.com/boardgame/1'),
    createGameRow(1, 'https://boardgamegeek.com/boardgame/2'),
  ];
  const sheet = createGamesSheet(rows);
  const getRange = sheet.getRange;
  sheet.getRange = (...args) => {
    const range = getRange(...args);
    const setValues = range.setValues;
    range.setValues = (values) => {
      setValues(values);
      rows.reverse();
    };
    return range;
  };
  const context = loadGameUpdater({ Date, Logger: { log() {} } });
  const snapshot = context.GameUpdater.loadRows(sheet);
  assert.equal(context.GameUpdater.writeRows(sheet, snapshot), true);
  assert.equal(sheet.writes.length, 1);
  assert.equal(sheet.writes[0].a1NotationOrRow, 2);
  assert.deepEqual(sheet.clears, []);
});

test('GameUpdater keeps triplicate identities ambiguous without losing position zero', () => {
  const context = loadGameUpdater({});
  const rows = [2, 3, 3, 3, 4].map((rowNumber) => ({ rowNumber }));
  const positions = context.GameUpdater.indexRowNumbers(rows);

  assert.equal(positions.get(2), 0);
  assert.equal(positions.get(3), null);
  assert.equal(positions.get(4), 4);
  assert.equal(positions.has(5), false);
});

test('GameUpdater preserves every input column edited during successful and failed fetches', () => {
  for (const status of [200, 503]) {
    const rows = [createGameRow(0, 'https://boardgamegeek.com/boardgame/42')];
    const sheet = createGamesSheet(rows);
    const inputs = [0, 2, 3, 5, 6];
    const context = loadGameUpdaterWithHttp({
      Date,
      Logger: { log() {} },
      PropertiesService: {
        getScriptProperties: () => ({ getProperty: () => null }),
      },
      SpreadsheetApp: {
        getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }),
      },
      UrlFetchApp: {
        fetch() {
          inputs.forEach((column) => {
            rows[0].values[column] = `edited-${column}`;
          });
          return {
            getResponseCode: () => status,
            getContentText: () => 'response',
          };
        },
      },
      Utilities: { sleep() {} },
      XmlService: {
        parse: () => ({
          getRootElement: () =>
            createXmlElement({
              child: {
                item: createGameItemWithTenAndElevenPlayerRecommendations(),
              },
            }),
        }),
      },
    });
    assert.equal(context.GameUpdater.run(), false);
    const finalValues = rows[0].values.slice();
    for (const write of sheet.writes) {
      finalValues.splice(
        write.column - 3,
        write.numColumns,
        ...write.values[0],
      );
      assert.ok(
        inputs.every(
          (column) =>
            column + 3 < write.column ||
            column + 3 >= write.column + write.numColumns,
        ),
      );
    }
    inputs.forEach((column) =>
      assert.equal(finalValues[column], `edited-${column}`),
    );
    assert.deepEqual(getArrayFormulaInputs(finalValues), Array(5).fill(null));
    assert.equal(
      finalValues[GAME_ERROR_MESSAGE_COLUMN],
      status === 503 ? 'HTTP 503' : '',
    );
    assert.ok(finalValues[GAME_LAST_UPDATED_AT_COLUMN] instanceof Date);
  }
});

for (const fails of [false, true]) {
  test(`GameUpdater fetches a changed fresh link on the next batch after ${fails ? 'failure' : 'success'}`, () => {
    const rows = [
      createGameRow(0, 'https://boardgamegeek.com/boardgame/1'),
      createGameRow(1, 'https://boardgamegeek.com/boardgame/2'),
    ];
    rows[1].values[GAME_LAST_UPDATED_AT_COLUMN] = new Date();
    rows[1].values[GAME_BOARD_GAME_RANK_COLUMN] = 123;
    rows[1].values[GAME_ERROR_MESSAGE_COLUMN] = fails ? 'old error' : '';
    rows[1].values[3] = 'manual location';
    const sheet = createGamesSheet(rows);
    const context = loadGameUpdater({
      Date,
      Logger: { log() {} },
      SpreadsheetApp: {
        getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }),
      },
    });
    const fetched = [];
    context.GameUpdater.fetchGameItem = (ref) => {
      fetched.push(ref.id);
      if (ref.id === '1')
        rows[1].gameLink = createGameRow(
          1,
          'https://boardgamegeek.com/boardgame/999',
        ).gameLink;
      if (ref.id === '999' && fails) throw new Error('new request failed');
      return {};
    };
    context.GameUpdater.applyGameItem = (row, _item, id, current) => {
      row.values[GAME_LAST_UPDATED_AT_COLUMN] = current;
      row.values[GAME_BOARD_GAME_RANK_COLUMN] = Number(id);
      row.values[GAME_ERROR_MESSAGE_COLUMN] = '';
    };
    assert.equal(context.GameUpdater.run(), true);
    for (const operation of sheet.operations) {
      const index = (operation.a1NotationOrRow ?? operation.row) - 2;
      const offset = operation.column - 3;
      const values =
        operation.type === 'setValues'
          ? operation.values[0]
          : Array(operation.numColumns).fill('');
      rows[index].values.splice(offset, operation.numColumns, ...values);
    }
    sheet.writes.length = 0;
    sheet.clears.length = 0;
    sheet.operations.length = 0;
    assert.equal(context.GameUpdater.run(), false);
    assert.deepEqual(fetched, ['1', '999']);
    const written = writtenGameRows(sheet)[1];
    assert.equal(written[GAME_BOARD_GAME_RANK_COLUMN], fails ? 123 : 999);
    assert.equal(
      written[GAME_ERROR_MESSAGE_COLUMN],
      fails ? 'new request failed' : '',
    );
    assert.equal(rows[1].values[3], 'manual location');
    assert.deepEqual(sheet.clears, []);
  });
}

test('GameUpdater preserves the refresh timestamp for a label-only link edit', () => {
  const url = 'https://boardgamegeek.com/boardgame/2';
  const rows = [createGameRow(0, url)];
  const sheet = createGamesSheet(rows);
  const context = loadGameUpdater({ Date, Logger: { log() {} } });
  const snapshot = context.GameUpdater.loadRows(sheet);
  rows[0].gameLink = { getText: () => 'manual label', getLinkUrl: () => url };
  assert.equal(context.GameUpdater.writeRows(sheet, snapshot), true);
  assert.deepEqual(sheet.writes, []);
  assert.deepEqual(sheet.clears, []);
});

test('GameUpdater reports a timestamp invalidation failure without overwriting the changed game', () => {
  const rows = [createGameRow(0, 'https://boardgamegeek.com/boardgame/2')];
  const sheet = createGamesSheet(rows);
  const context = loadGameUpdater({ Date, Logger: { log() {} } });
  const snapshot = context.GameUpdater.loadRows(sheet);
  rows[0].gameLink = createGameRow(
    0,
    'https://boardgamegeek.com/boardgame/999',
  ).gameLink;
  const getRange = sheet.getRange;
  sheet.getRange = (...args) => {
    const range = getRange(...args);
    range.clearContent = () => {
      throw new Error('clearContent failed');
    };
    return range;
  };
  assert.throws(
    () => context.GameUpdater.writeRows(sheet, snapshot),
    /clearContent failed/,
  );
  assert.deepEqual(sheet.writes, []);
  assert.deepEqual(sheet.clears, []);
  assert.equal(
    rows[0].gameLink.getLinkUrl(),
    'https://boardgamegeek.com/boardgame/999',
  );
});

test('GameUpdater saves progress within the time budget and finishes large sheets', () => {
  for (const staleLastRow of [false, true]) {
    const { ClockDate, clock } = createClockDate(Date.UTC(2026, 0, 1));
    const rows = Array.from({ length: 500 }, (_, index) => {
      const row = createGameRow(
        index,
        `https://boardgamegeek.com/boardgame/${index + 1}`,
      );
      row.values[GAME_LAST_UPDATED_AT_COLUMN] = new Date(clock.nowMs);
      return row;
    });
    if (staleLastRow) rows[499].values[GAME_LAST_UPDATED_AT_COLUMN] = '';
    const sheet = createGamesSheet(rows);
    const getRange = sheet.getRange;
    let invocationStart = clock.nowMs;
    sheet.getRange = (...args) => {
      const range = getRange(...args);
      for (const method of [
        'getValues',
        'getRichTextValues',
        'setValues',
        'clearContent',
      ]) {
        const original = range[method];
        range[method] = (...values) => {
          clock.nowMs += 100;
          assert.ok(
            clock.nowMs - invocationStart < 360000,
            'execution must not reach the hard timeout',
          );
          const result = original(...values);
          if (method === 'setValues') {
            rows[args[0] - 2].values.splice(
              args[1] - 3,
              args[3],
              ...values[0][0],
            );
          }
          return result;
        };
      }
      return range;
    };
    const context = loadGameUpdater({
      Date: ClockDate,
      Logger: { log() {} },
      SpreadsheetApp: {
        getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }),
      },
    });
    let fetches = 0;
    context.GameUpdater.fetchGameItem = () => {
      fetches++;
      return {};
    };
    context.GameUpdater.applyGameItem = (row, _item, _id, now) => {
      row.values[GAME_LAST_UPDATED_AT_COLUMN] = now;
    };
    let pending = true;
    let batches = 0;
    while (pending && batches < 10) {
      invocationStart = clock.nowMs;
      pending = context.GameUpdater.run();
      batches++;
      assert.ok(
        clock.nowMs - invocationStart < 182000,
        'sheet writes must share the soft budget',
      );
    }
    assert.equal(
      pending,
      false,
      'queue must finish rather than restart the same prefix',
    );
    assert.ok(batches > 1);
    assert.equal(fetches, staleLastRow ? 1 : 0);
    assert.ok(
      rows.every((row) =>
        getArrayFormulaInputs(row.values).every((value) => value === null),
      ),
    );
    assert.ok(rows[499].values[GAME_LAST_UPDATED_AT_COLUMN] instanceof Date);
  }
});

test('GameUpdater keeps its saved position when a resumed sheet write fails', () => {
  const { ClockDate, clock } = createClockDate(Date.UTC(2026, 0, 1));
  const rows = Array.from({ length: 3 }, (_, index) => {
    const row = createGameRow(
      index,
      `https://boardgamegeek.com/boardgame/${index + 1}`,
    );
    row.values[GAME_LAST_UPDATED_AT_COLUMN] = new Date(clock.nowMs);
    return row;
  });
  const sheet = createGamesSheet(rows);
  const getRange = sheet.getRange;
  let fail = false;
  sheet.getRange = (...args) => {
    const range = getRange(...args);
    const setValues = range.setValues;
    range.setValues = (values) => {
      if (fail) throw new Error('resumed write failed');
      setValues(values);
      if (args[0] === 2 && args[1] === 10) clock.nowMs += 180000;
    };
    return range;
  };
  const context = loadGameUpdater({
    Date: ClockDate,
    Logger: { log() {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }),
    },
  });
  assert.equal(context.GameUpdater.run(), true);
  const properties = context.PropertiesService.getScriptProperties();
  const saved = properties.getProperty('GAME_BATCH_ROW');
  assert.equal(JSON.parse(saved).row, 1);
  fail = true;
  assert.throws(() => context.GameUpdater.run(), /resumed write failed/);
  assert.equal(properties.getProperty('GAME_BATCH_ROW'), saved);
  fail = false;
  sheet.writes.length = 0;
  assert.equal(context.GameUpdater.run(), false);
  assert.deepEqual(
    [...new Set(sheet.writes.map((write) => write.a1NotationOrRow))],
    [3, 4],
  );
  assert.equal(properties.getProperty('GAME_BATCH_ROW'), null);
});
