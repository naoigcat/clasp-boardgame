/**
 * RatingUpdater import and snapshot-preservation tests.
 *
 * Verifies pagination until Bodoge's empty marker, title alias expansion, and
 * keeping the previous Ratings sheet when HTML is unrecognized mid-import.
 */
const assert = require('node:assert/strict');
const test = require('node:test');

const {
  createSheet,
  getCalls,
  loadScripts,
} = require('./helpers/appScriptHarness');

/**
 * Creates the Apps Script doubles required to exercise RatingUpdater end to end.
 */
function createRatingSandbox({ ratingsSheet, userId, responses }) {
  const responseQueue = [...responses];

  return {
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Ratings' ? ratingsSheet : null;
          },
        };
      },
    },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty(propertyName) {
            return propertyName === 'BODOGE_USER_ID' ? userId : null;
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
  };
}

/**
 * Loads the production dependency chain used by RatingUpdater.
 */
function loadRatingUpdater(sandbox) {
  return loadScripts(sandbox, [
    { path: 'src/config/AppConfig.ts', exports: ['BODOGE_CONFIG'] },
    { path: 'src/config/SheetSchema.ts', exports: [] },
    { path: 'src/config/TitleRules.ts', exports: [] },
    { path: 'src/infrastructure/HttpClient.ts', exports: ['HttpClient'] },
    {
      path: 'src/infrastructure/ScriptPropertyStore.ts',
      exports: ['ScriptPropertyStore'],
    },
    { path: 'src/infrastructure/SpreadsheetGateway.ts', exports: [] },
    { path: 'src/services/RatingUpdater.ts', exports: ['RatingUpdater'] },
  ]);
}

/**
 * Builds one Bodoge rating card in the markup shape consumed by RatingUpdater.
 */
function ratingCard(title, rating) {
  return [
    '<a class="list--interests-item-title">',
    `<div class="list--interests-item-title-japanese">${title}</div>`,
    `<div class="rating--result-stars" data-rating-mode="result" data-rating-result="${rating}">`,
    '</div>',
    '</a>',
  ].join('');
}

/**
 * Markup Bodoge uses for an empty played-games page or a page past the last result.
 */
function emptyPlayedGamesPage() {
  return '<p class="empty">検索結果が存在しないか、マイボードゲームが未登録のユーザーです</p>';
}

test('RatingUpdater writes aliased ratings without clearing a header-only sheet', () => {
  const ratingsSheet = createSheet('Ratings', 1);
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: ratingCard('#hashtag', '4') },
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  context.RatingUpdater.run();

  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), [
    {
      type: 'setValues',
      row: 2,
      column: 1,
      numRows: 1,
      numColumns: 2,
      values: [['ハッシュタグ', '4']],
    },
  ]);
});

test('RatingUpdater writes ratings then clears surplus Ratings rows', () => {
  const ratingsSheet = createSheet('Ratings', 4);
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: ratingCard('#hashtag', '4') },
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  context.RatingUpdater.run();

  // Write first so a failed setValues cannot wipe Ratings, then trim only the
  // abandoned physical rows left by a shorter replacement snapshot.
  assert.deepEqual(
    ratingsSheet.calls
      .filter(
        (call) => call.type === 'setValues' || call.type === 'clearContent',
      )
      .map((call) => call.type),
    ['setValues', 'clearContent'],
  );
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), [
    {
      type: 'clearContent',
      row: 3,
      column: 1,
      numRows: 2,
      numColumns: 2,
    },
  ]);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), [
    {
      type: 'setValues',
      row: 2,
      column: 1,
      numRows: 1,
      numColumns: 2,
      values: [['ハッシュタグ', '4']],
    },
  ]);
});

test('RatingUpdater leaves Ratings intact when setValues fails', () => {
  const ratingsSheet = createSheet('Ratings', 4, { failOnSetValues: true });
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: ratingCard('カタン', '5') },
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  assert.throws(() => context.RatingUpdater.run(), /setValues failed/);

  // Clearing must not run after a write failure; otherwise the previous
  // Ratings snapshot would disappear until the next successful Bodoge import.
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});

test('RatingUpdater leaves the Ratings sheet untouched without a configured user ID', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: null,
    responses: [],
  });
  const context = loadRatingUpdater(sandbox);

  context.RatingUpdater.run();

  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});

test('RatingUpdater keeps existing rows when Bodoge HTML is unrecognized', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: '<html><body>login required</body></html>' },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  assert.throws(
    () => context.RatingUpdater.run(),
    /Unrecognized Bodoge ratings page HTML/,
  );
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});

test('RatingUpdater keeps existing rows when a later Bodoge page is unrecognized', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: ratingCard('カタン', '5') },
      { status: 200, body: '<html><body>maintenance</body></html>' },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  assert.throws(
    () => context.RatingUpdater.run(),
    /Unrecognized Bodoge ratings page HTML/,
  );
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});

test('RatingUpdater clears Ratings when Bodoge reports an explicit empty list', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [{ status: 200, body: emptyPlayedGamesPage() }],
  });
  const context = loadRatingUpdater(sandbox);

  context.RatingUpdater.run();

  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), [
    {
      type: 'clearContent',
      row: 2,
      column: 1,
      numRows: 2,
      numColumns: 2,
    },
  ]);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});

test('RatingUpdater keeps existing rows when Bodoge keeps returning rating cards past the page cap', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  const maxPageCount = 100;
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    // One response per allowed page, all with cards and no empty marker, so the
    // import must stop on the page cap instead of looping until runtime expires.
    responses: Array.from({ length: maxPageCount }, (_, index) => ({
      status: 200,
      body: ratingCard(`ゲーム${index + 1}`, '3'),
    })),
  });
  const context = loadRatingUpdater(sandbox);

  assert.equal(context.BODOGE_CONFIG.MAX_PAGE_COUNT, maxPageCount);
  assert.throws(
    () => context.RatingUpdater.run(),
    /Bodoge ratings pagination exceeded 100 pages/,
  );
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});
