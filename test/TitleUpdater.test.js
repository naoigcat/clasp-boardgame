/**
 * TitleUpdater normalization and batch-ordering tests.
 *
 * Pure spelling rules are loaded without Apps Script doubles; run() scenarios
 * cover Rankings URL intake, failed-row deferral, and the final retry pass.
 */
const assert = require('node:assert/strict');
const test = require('node:test');

const { loadScripts } = require('./helpers/appScriptHarness');

/** Titles sheet column indexes mirrored from SheetSchema for assertions. */
const TITLE_URL_COLUMN = 0;
const TITLE_SOURCE_COLUMN = 1;
const TITLE_NORMALIZED_COLUMN = 2;
const TITLE_ERROR_COLUMN = 3;

/**
 * Loads only the title-normalization dependencies so the tests stay independent
 * from Apps Script services that are irrelevant to spelling rules.
 */
function loadTitleUpdater() {
  return loadScripts({}, [
    { path: 'src/config/TitleRules.ts', exports: [] },
    { path: 'src/services/TitleUpdater.ts', exports: ['TitleUpdater'] },
  ]);
}

/**
 * Loads the Apps Script dependency chain used by TitleUpdater.run.
 */
function loadTitleUpdaterService(sandbox) {
  return loadScripts(sandbox, [
    { path: 'src/config/AppConfig.ts', exports: [] },
    { path: 'src/config/SheetSchema.ts', exports: [] },
    { path: 'src/config/TitleRules.ts', exports: [] },
    { path: 'src/shared/ErrorUtils.ts', exports: [] },
    { path: 'src/infrastructure/HttpClient.ts', exports: ['HttpClient'] },
    { path: 'src/infrastructure/SpreadsheetGateway.ts', exports: [] },
    { path: 'src/services/TitleUpdater.ts', exports: ['TitleUpdater'] },
  ]);
}

/**
 * Creates Titles and Rankings sheet doubles around queued HTTP responses.
 */
function createTitleSandbox({ titleRows, responses, failOnSetValues = false }) {
  const responseQueue = [...responses];
  const writes = [];
  const clears = [];
  const operations = [];
  const titlesSheet = {
    writes,
    clears,
    operations,
    getLastRow() {
      // Header row plus every physical data row, including mid-sheet blanks.
      return titleRows.length + 1;
    },
    getRange(a1NotationOrRow, column, numRows, numColumns) {
      if (typeof a1NotationOrRow === 'string') {
        return {
          getValues() {
            return titleRows.map((row) => row.slice());
          },
        };
      }

      return {
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
        setValues(values) {
          if (failOnSetValues) {
            throw new Error('setValues failed');
          }
          const write = {
            row: a1NotationOrRow,
            column,
            numRows,
            numColumns,
            values: JSON.parse(JSON.stringify(values)),
          };
          writes.push(write);
          operations.push({ type: 'setValues', ...write });
        },
      };
    },
  };

  return {
    titlesSheet,
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            if (name === 'Titles') {
              return titlesSheet;
            }
            if (name === 'Rankings') {
              return {
                getRange() {
                  return {
                    getValues() {
                      return [];
                    },
                  };
                },
              };
            }
            return null;
          },
        };
      },
    },
    UrlFetchApp: {
      fetch(url, options = {}) {
        const response = responseQueue.shift();
        assert.ok(response, `Unexpected fetch: ${url}`);
        // Mirror UrlFetchApp: non-2xx throws unless muteHttpExceptions is set.
        if (
          (response.status < 200 || response.status >= 300) &&
          options.muteHttpExceptions !== true
        ) {
          throw new Error(`Request failed for ${url}: ${response.status}`);
        }
        return {
          getResponseCode() {
            return response.status;
          },
          getContentText() {
            return response.body;
          },
        };
      },
    },
    Utilities: {
      sleep() {},
    },
    Logger: {
      log() {},
    },
  };
}

/**
 * Builds the Board Game Arena markup fragment consumed by TitleUpdater.
 */
function titlePage(gameName) {
  return `id="game_name" class="block gamename"\n>${gameName}</a`;
}

test('TitleUpdater applies generic cleanup before exact aliases', () => {
  const context = loadTitleUpdater();

  assert.equal(
    context.TitleUpdater.normalizeTitle('ザ・クルー 深海に眠る遺跡 第2版'),
    'ザ・クルー：深海に眠る遺跡',
  );
});

test('TitleUpdater keeps family aliases while preserving a title suffix', () => {
  const context = loadTitleUpdater();

  assert.equal(
    context.TitleUpdater.normalizeTitle('チケット・トゥ・ライド：ヨーロッパ'),
    'チケットトゥライド：ヨーロッパ',
  );
});

test('TitleUpdater handles a source-specific title that needs no generic cleanup', () => {
  const context = loadTitleUpdater();

  assert.equal(
    context.TitleUpdater.normalizeTitle('タペストリー ～文明の錦の御旗～'),
    'タペストリー',
  );
});

test('TitleUpdater normalizes edition-only titles to an empty string', () => {
  const context = loadTitleUpdater();

  assert.equal(context.TitleUpdater.normalizeTitle('《新版》'), '');
  assert.equal(context.TitleUpdater.normalizeTitle('第1版'), '');
  assert.equal(context.TitleUpdater.normalizeTitle('-Deluxe-'), '');
});

test('TitleUpdater records empty normalization as an error and ends the preferred queue', () => {
  const titleRows = [['https://example.com/empty-norm', '', '', '']];
  const sandbox = createTitleSandbox({
    titleRows,
    responses: [{ status: 200, body: titlePage('《新版》') }],
  });
  const context = loadTitleUpdaterService(sandbox);

  // First pass records the empty normalization as a failure and schedules the
  // one-shot retry pass used for other permanent title errors.
  assert.equal(context.TitleUpdater.run(), true);

  let written = sandbox.titlesSheet.writes[0].values;
  assert.equal(written[0][TITLE_SOURCE_COLUMN], '《新版》');
  assert.equal(written[0][TITLE_NORMALIZED_COLUMN], '');
  assert.equal(written[0][TITLE_ERROR_COLUMN], 'normalized title is empty');

  titleRows.splice(0, titleRows.length, ...written);
  sandbox.titlesSheet.writes.length = 0;

  // Without recording an error, this row would stay preferred and run() would
  // keep returning true, holding the shared trigger open indefinitely.
  assert.equal(context.TitleUpdater.run(), false);
  written = sandbox.titlesSheet.writes[0].values;
  assert.equal(written[0][TITLE_NORMALIZED_COLUMN], '');
  assert.equal(written[0][TITLE_ERROR_COLUMN], 'normalized title is empty');
});

test('TitleUpdater prefers unfailed title rows over permanently failing head rows', () => {
  const titleRows = [
    ['https://example.com/fail-1', '', '', 'previous error'],
    ['https://example.com/fail-2', '', '', 'previous error'],
    ['https://example.com/ok', '', '', ''],
  ];
  const sandbox = createTitleSandbox({
    titleRows,
    responses: [
      { status: 200, body: titlePage('カタン') },
      { status: 500, body: 'error' },
      { status: 500, body: 'error' },
    ],
  });
  const context = loadTitleUpdaterService(sandbox);

  assert.equal(context.TitleUpdater.run(), true);

  let written = sandbox.titlesSheet.writes[0].values;
  assert.equal(written[0][TITLE_NORMALIZED_COLUMN], '');
  assert.equal(written[0][TITLE_ERROR_COLUMN], 'previous error');
  assert.equal(written[1][TITLE_NORMALIZED_COLUMN], '');
  assert.equal(written[1][TITLE_ERROR_COLUMN], 'previous error');
  assert.equal(written[2][TITLE_SOURCE_COLUMN], 'カタン');
  assert.equal(written[2][TITLE_NORMALIZED_COLUMN], 'カタン');
  assert.equal(written[2][TITLE_ERROR_COLUMN], '');

  // Persist the first write so the retry pass sees the successful row as done.
  titleRows.splice(0, titleRows.length, ...written);
  sandbox.titlesSheet.writes.length = 0;

  assert.equal(context.TitleUpdater.run(), false);
  written = sandbox.titlesSheet.writes[0].values;
  assert.match(
    written[0][TITLE_ERROR_COLUMN],
    /Board Game Arena returned HTTP 500/,
  );
  assert.match(
    written[1][TITLE_ERROR_COLUMN],
    /Board Game Arena returned HTTP 500/,
  );
  assert.equal(written[2][TITLE_NORMALIZED_COLUMN], 'カタン');
});

test('TitleUpdater retries failed titles once when only failures remain', () => {
  const titleRows = [['https://example.com/retry', '', '', 'previous error']];
  const sandbox = createTitleSandbox({
    titleRows,
    responses: [{ status: 200, body: titlePage('カルカソンヌ') }],
  });
  const context = loadTitleUpdaterService(sandbox);

  assert.equal(context.TitleUpdater.run(), false);

  const written = sandbox.titlesSheet.writes[0].values;
  assert.equal(written[0][TITLE_URL_COLUMN], 'https://example.com/retry');
  assert.equal(written[0][TITLE_NORMALIZED_COLUMN], 'カルカソンヌ');
  assert.equal(written[0][TITLE_ERROR_COLUMN], '');
});

test('TitleUpdater clears surplus Titles rows after compacting mid-sheet blanks', () => {
  const titleRows = [
    ['https://example.com/a', 'カタン', 'カタン', ''],
    ['', '', '', ''],
    ['https://example.com/b', '', '', ''],
  ];
  const sandbox = createTitleSandbox({
    titleRows,
    responses: [{ status: 200, body: titlePage('カルカソンヌ') }],
  });
  const context = loadTitleUpdaterService(sandbox);

  assert.equal(context.TitleUpdater.run(), false);

  // Write first so a failed setValues cannot wipe Titles, then trim only the
  // abandoned physical row left by compacting out the blank URL slot.
  assert.deepEqual(
    sandbox.titlesSheet.operations.map((operation) => operation.type),
    ['setValues', 'clearContent'],
  );
  assert.deepEqual(sandbox.titlesSheet.clears, [
    { row: 4, column: 1, numRows: 1, numColumns: 4 },
  ]);
  const written = sandbox.titlesSheet.writes[0].values;
  assert.equal(written.length, 2);
  assert.equal(written[0][TITLE_URL_COLUMN], 'https://example.com/a');
  assert.equal(written[1][TITLE_URL_COLUMN], 'https://example.com/b');
  assert.equal(written[1][TITLE_NORMALIZED_COLUMN], 'カルカソンヌ');
});

test('TitleUpdater leaves Titles intact when compacted setValues fails', () => {
  const titleRows = [
    ['https://example.com/a', 'カタン', 'カタン', ''],
    ['', '', '', ''],
    ['https://example.com/b', '', '', ''],
  ];
  const sandbox = createTitleSandbox({
    titleRows,
    responses: [{ status: 200, body: titlePage('カルカソンヌ') }],
    failOnSetValues: true,
  });
  const context = loadTitleUpdaterService(sandbox);

  assert.throws(() => context.TitleUpdater.run(), /setValues failed/);

  // Clearing must not run after a write failure; otherwise source titles and
  // manual corrections would disappear even though Rankings cannot restore them.
  assert.deepEqual(sandbox.titlesSheet.operations, []);
  assert.deepEqual(sandbox.titlesSheet.clears, []);
  assert.deepEqual(titleRows, [
    ['https://example.com/a', 'カタン', 'カタン', ''],
    ['', '', '', ''],
    ['https://example.com/b', '', '', ''],
  ]);
});
