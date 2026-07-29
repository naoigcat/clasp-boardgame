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
  const rangeReads = [];

  return {
    writes,
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
          writes.push({
            a1NotationOrRow,
            column,
            numRows,
            numColumns,
            values,
          });
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
 */
function createGamesSheet(rows) {
  const writes = [];
  const rangeReads = [];

  return {
    writes,
    rangeReads,
    getLastRow() {
      // Header plus every managed data row supplied to the double.
      return rows.length + 1;
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
          return rows.slice(0, numRows).map((row) => [row.gameLink]);
        },
        getValues() {
          rangeReads.push({
            row: a1NotationOrRow,
            column,
            numRows,
            numColumns,
          });
          return rows.slice(0, numRows).map((row) => row.values.slice());
        },
        setValues(values) {
          writes.push({
            a1NotationOrRow,
            column,
            numRows,
            numColumns,
            values,
          });
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

test('GameUpdater clears formula inputs only for Games rows refreshed within the batch', () => {
  const rows = Array.from({ length: 51 }, (_, index) =>
    createGameRow(index, `https://boardgamegeek.com/boardgame/${index + 1}`),
  );
  const gamesSheet = createGamesSheet(rows);
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

  assert.equal(context.GameUpdater.run(), true);

  const writtenRows = gamesSheet.writes[0].values;
  // null clears the cell for ARRAYFORMULA; '' would leave a blocking blank.
  assert.deepEqual(getArrayFormulaInputs(writtenRows[0]), Array(5).fill(null));
  assert.deepEqual(
    getArrayFormulaInputs(writtenRows[50]),
    getArrayFormulaInputs(rows[50].values),
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
    createGameRow(index, 'not-a-boardgamegeek-url'),
  );
  const successRow = createGameRow(
    50,
    'https://boardgamegeek.com/boardgame/42',
  );
  const rows = [...failingRows, successRow];
  const gamesSheet = createGamesSheet(rows);
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

  assert.equal(context.GameUpdater.run(), true);
  assert.deepEqual(
    getArrayFormulaInputs(gamesSheet.writes[0].values[50]),
    getArrayFormulaInputs(successRow.values),
  );

  // The sheet double reads from the source row objects, so persist the first
  // batch write before asserting that the next batch can move past failures.
  gamesSheet.writes[0].values.forEach((values, index) => {
    rows[index].values = values;
  });
  gamesSheet.writes.length = 0;

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
