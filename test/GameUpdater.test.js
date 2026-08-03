/**
 * GameUpdater batch and refresh-policy tests.
 *
 * Focuses on empty sheets, oldest-first batching, derived-column clearing on
 * success, advancing failure timestamps so permanent errors rotate out, and
 * writing only the real 2–10 player-recommendation columns before rank.
 */
const assert = require('node:assert/strict');
const test = require('node:test');

const { loadScripts } = require('./helpers/appScriptHarness');

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
 * `surplusRowCount` simulates abandoned B–AA cells below the first blank link.
 * `failOnSetValues` proves surplus clearing runs only after a successful rewrite.
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
            index < rows.length ? rows[index].gameLink : null,
          ]);
        },
        getValues() {
          rangeReads.push({
            row: a1NotationOrRow,
            column,
            numRows,
            numColumns,
          });
          return Array.from({ length: numRows }, (_, index) =>
            index < rows.length
              ? rows[index].values.slice()
              : Array(numColumns).fill('stale'),
          );
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
 * Supplies a Date stand-in whose `now()` advances under test control.
 *
 * Prototype sharing keeps host-created Date values passing `instanceof Date`
 * checks inside the VM while runtime-budget helpers read the fake clock.
 */
function createClockDate(initialMs = 0) {
  const clock = { nowMs: initialMs };
  function ClockDate(...args) {
    if (args.length === 0) {
      return new Date(clock.nowMs);
    }

    return new Date(...args);
  }
  ClockDate.now = () => clock.nowMs;
  ClockDate.parse = Date.parse;
  ClockDate.UTC = Date.UTC;
  ClockDate.prototype = Date.prototype;
  return { ClockDate, clock };
}

/**
 * Loads the dependencies evaluated by GameUpdater before it begins API work.
 */
function loadGameUpdater(sandbox) {
  return loadScripts(sandbox, [
    { path: 'src/config/AppConfig.ts', exports: [] },
    { path: 'src/config/SheetSchema.ts', exports: [] },
    { path: 'src/config/TitleRules.ts', exports: [] },
    { path: 'src/shared/DateUtils.ts', exports: [] },
    { path: 'src/shared/ErrorUtils.ts', exports: [] },
    { path: 'src/shared/XmlUtils.ts', exports: [] },
    { path: 'src/infrastructure/SpreadsheetGateway.ts', exports: [] },
    { path: 'src/services/GameUpdater.ts', exports: ['GameUpdater'] },
  ]);
}

/**
 * Loads GameUpdater with the HTTP client path used for BoardGameGeek fetches.
 */
function loadGameUpdaterWithHttp(sandbox) {
  return loadScripts(sandbox, [
    { path: 'src/config/AppConfig.ts', exports: [] },
    { path: 'src/config/SheetSchema.ts', exports: [] },
    { path: 'src/config/TitleRules.ts', exports: [] },
    { path: 'src/shared/DateUtils.ts', exports: [] },
    { path: 'src/shared/ErrorUtils.ts', exports: [] },
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

  assert.deepEqual(gamesSheet.rangeReads, [
    { row: 2, column: 1, numRows: 2, numColumns: 1 },
    { row: 2, column: 2, numRows: 2, numColumns: 26 },
  ]);
  assert.equal(gamesSheet.writes[0].values.length, 2);
  // A regression that reintroduced `$A$2:$A` would throw in createGamesSheet.
  reportedLastRow = 1;
  gamesSheet.rangeReads.length = 0;
  gamesSheet.writes.length = 0;
  assert.equal(context.GameUpdater.run(), false);
  assert.deepEqual(gamesSheet.rangeReads, []);
});

test('GameUpdater writes Games values then clears surplus B–AA rows', () => {
  const rows = [
    createGameRow(0, 'https://boardgamegeek.com/boardgame/1'),
    createGameRow(1, 'https://boardgamegeek.com/boardgame/2'),
  ];
  // Two abandoned physical rows remain below the blank-link end marker after a
  // shortened list; their B–AA cells must be trimmed after a successful write.
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

  // Write first so a failed setValues cannot wipe managed Games values, then
  // trim only abandoned B–AA cells; column A is left alone for rich-text links.
  assert.deepEqual(
    gamesSheet.operations.map((operation) => operation.type),
    ['setValues', 'clearContent'],
  );
  assert.deepEqual(gamesSheet.clears, [
    { row: 4, column: 2, numRows: 2, numColumns: 26 },
  ]);
  assert.equal(gamesSheet.writes[0].values.length, 2);
  assert.equal(gamesSheet.writes[0].column, 2);
});

test('GameUpdater clears surplus B–AA rows when every managed game is still fresh', () => {
  const rows = [
    createGameRow(0, 'https://boardgamegeek.com/boardgame/1'),
    createGameRow(1, 'https://boardgamegeek.com/boardgame/2'),
  ];
  // Within the seven-day refresh window so countPendingRows is zero; surplus
  // cleanup must still run after column A was shortened.
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
  // Fresh rows still rewrite current B–AA values, then trim abandoned cells so
  // a pending-free early exit cannot skip surplus cleanup.
  assert.deepEqual(
    gamesSheet.operations.map((operation) => operation.type),
    ['setValues', 'clearContent'],
  );
  assert.deepEqual(gamesSheet.clears, [
    { row: 4, column: 2, numRows: 2, numColumns: 26 },
  ]);
  assert.equal(gamesSheet.writes[0].values.length, 2);
  assert.equal(gamesSheet.writes[0].column, 2);
  assert.deepEqual(
    getArrayFormulaInputs(gamesSheet.writes[0].values[0]),
    Array(5).fill(null),
  );
});

test('GameUpdater clears all B–AA data rows when the managed Games list is empty', () => {
  // First blank link ends the managed block immediately, but abandoned B–AA
  // cells can remain after every column-A link is deleted.
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
  // Empty managed lists skip setValues and clear B–AA from the first data row,
  // matching Titles/Ratings empty-write cleanup without touching column A.
  assert.deepEqual(
    gamesSheet.operations.map((operation) => operation.type),
    ['clearContent'],
  );
  assert.deepEqual(gamesSheet.clears, [
    { row: 2, column: 2, numRows: 3, numColumns: 26 },
  ]);
  assert.deepEqual(gamesSheet.writes, []);
});

test('GameUpdater leaves Games intact when setValues fails before surplus clear', () => {
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

  // Clearing must not run after a write failure; otherwise managed B–AA values
  // would disappear even though BoardGameGeek cannot restore the prior snapshot.
  assert.deepEqual(gamesSheet.operations, []);
  assert.deepEqual(gamesSheet.clears, []);
  assert.deepEqual(gamesSheet.writes, []);
});

test('GameUpdater clears formula inputs for refreshed and skipped Games rows', () => {
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

  const writtenRows = gamesSheet.writes[0].values;
  // null clears the cell for ARRAYFORMULA; '' would leave a blocking blank.
  assert.deepEqual(getArrayFormulaInputs(writtenRows[0]), Array(5).fill(null));
  assert.deepEqual(
    getArrayFormulaInputs(writtenRows[50]),
    Array(5).fill(null),
  );
});

test('GameUpdater preserves formula inputs when a game refresh fails', () => {
  const row = createGameRow(0, 'not-a-boardgamegeek-url');
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
    getArrayFormulaInputs(gamesSheet.writes[0].values[0]),
    getArrayFormulaInputs(row.values),
  );
  assert.ok(
    gamesSheet.writes[0].values[0][GAME_LAST_UPDATED_AT_COLUMN] instanceof Date,
  );
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
  assert.deepEqual(
    getArrayFormulaInputs(gamesSheet.writes[0].values[50]),
    Array(5).fill(null),
  );

  // The sheet double reads from the source row objects, so persist the first
  // batch write before asserting that the next batch can move past failures.
  gamesSheet.writes[0].values.forEach((values, index) => {
    rows[index].values = values;
  });
  gamesSheet.writes.length = 0;
  clock.nowMs = batchStartedAtMs;

  assert.equal(context.GameUpdater.run(), false);
  assert.deepEqual(
    getArrayFormulaInputs(gamesSheet.writes[0].values[50]),
    Array(5).fill(null),
  );
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

  const written = gamesSheet.writes[0].values[0];
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
  assert.deepEqual(getArrayFormulaInputs(written), priorFormulaInputs);
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
    gamesSheet.writes[0].values[0][GAME_ERROR_MESSAGE_COLUMN],
    /^HTTP 503$/,
  );
  assert.deepEqual(
    getArrayFormulaInputs(gamesSheet.writes[0].values[0]),
    getArrayFormulaInputs(row.values),
  );
  assert.ok(
    gamesSheet.writes[0].values[0][GAME_LAST_UPDATED_AT_COLUMN] instanceof Date,
  );
});
