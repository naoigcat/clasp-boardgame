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
    { path: 'src/shared/SpreadsheetUtils.ts', exports: [] },
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
 *
 * Joins with newlines to mirror pretty-printed Bodoge HTML; the card regex must
 * span those newlines or parsePage would treat real pages as unrecognized.
 */
function ratingCard(title, rating) {
  return [
    '<a class="list--interests-item-title">',
    `<div class="list--interests-item-title-japanese">${title}</div>`,
    `<div class="rating--result-stars" data-rating-mode="result" data-rating-result="${rating}">`,
    '</div>',
    '</a>',
  ].join('\n');
}

/**
 * Markup Bodoge uses for an empty played-games page or a page past the last result.
 */
function emptyPlayedGamesPage() {
  return '<p class="empty">検索結果が存在しないか、マイボードゲームが未登録のユーザーです</p>';
}

test('RatingUpdater logs and skips when the Ratings sheet is missing', () => {
  const logs = [];
  const sandbox = createRatingSandbox({
    ratingsSheet: null,
    userId: 'user-1',
    responses: [],
  });
  sandbox.Logger = {
    log(message) {
      logs.push(message);
    },
  };
  const context = loadRatingUpdater(sandbox);

  context.RatingUpdater.run();

  assert.deepEqual(logs, [
    'Ratings sheet "Ratings" is missing; skipping rating import until the tab is restored.',
  ]);
});

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

test('RatingUpdater escapes formula-like titles before setValues', () => {
  const formulaTitle = '+cmd|a';
  const ratingsSheet = createSheet('Ratings', 1);
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: ratingCard(formulaTitle, '5') },
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  context.RatingUpdater.run();

  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), [
    {
      type: 'setValues',
      row: 2,
      column: 1,
      numRows: 1,
      numColumns: 2,
      // Bodoge titles are external text; prefix so Sheets stores them as literals.
      values: [[`'${formulaTitle}`, '5']],
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

/**
 * Card chrome without a Japanese title div — hasCards true, extraction fails.
 */
function cardWithoutJapaneseTitle(englishTitle = 'Catan') {
  return [
    '<a class="list--interests-item-title">',
    `<div class="list--interests-item-title-english">${englishTitle}</div>`,
    '<div class="rating--result-stars" data-rating-mode="result" data-rating-result="5">',
    '</div>',
    '</a>',
  ].join('\n');
}

/**
 * Card with a Japanese title but no star-rating markup — rating extraction fails.
 */
function cardWithoutRatingMarkup(title = 'カタン') {
  return [
    '<a class="list--interests-item-title">',
    `<div class="list--interests-item-title-japanese">${title}</div>`,
    '</a>',
  ].join('\n');
}

/**
 * Card whose rating attribute is present but empty — treated as markup failure.
 */
function cardWithEmptyRating(title = 'カタン') {
  return ratingCard(title, '');
}

test('RatingUpdater keeps existing rows when cards match but no titles can be extracted', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  // Fail on the first unextractable card page rather than waiting for the empty
  // marker; otherwise a later empty page could be mistaken for a clean empty list.
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: cardWithoutJapaneseTitle() },
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  assert.throws(
    () => context.RatingUpdater.run(),
    /Bodoge ratings page contained cards without extractable Japanese titles/,
  );
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});

test('RatingUpdater keeps existing rows when title cleanup leaves only an empty string', () => {
  // Regression: edition/expansion markers, whitespace, or slash-only titles
  // normalize to "" and must not be written as blank Ratings rows.
  const emptyAfterCleanupTitles = ['（拡張）', '：新版', '   ', '/'];
  for (const title of emptyAfterCleanupTitles) {
    const ratingsSheet = createSheet('Ratings', 3);
    const sandbox = createRatingSandbox({
      ratingsSheet,
      userId: 'user-1',
      responses: [
        { status: 200, body: ratingCard(title, '4') },
        { status: 200, body: emptyPlayedGamesPage() },
      ],
    });
    const context = loadRatingUpdater(sandbox);

    assert.throws(
      () => context.RatingUpdater.run(),
      /Bodoge ratings page contained cards without extractable Japanese titles/,
      `expected abort for cleaned-empty title ${JSON.stringify(title)}`,
    );
    assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
    assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
  }
});

test('RatingUpdater keeps existing rows when cards have titles but no rating markup', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  // Regression: a changed rating attribute must not write title-only rows that
  // replace the previous complete Ratings snapshot with empty values.
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: cardWithoutRatingMarkup() },
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  assert.throws(
    () => context.RatingUpdater.run(),
    /Bodoge ratings page contained cards without extractable ratings/,
  );
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});

test('RatingUpdater keeps existing rows when every card rating attribute is empty', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: cardWithEmptyRating() },
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  assert.throws(
    () => context.RatingUpdater.run(),
    /Bodoge ratings page contained cards without extractable ratings/,
  );
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});

test('RatingUpdater keeps existing rows when a later page has titles but no extractable ratings', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  // Earlier pages with real rows must not be written when a later page yields
  // titles without ratings (partial import / changed rating markup mid-run).
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: ratingCard('カタン', '5') },
      { status: 200, body: cardWithoutRatingMarkup('チッキットゥライド') },
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  assert.throws(
    () => context.RatingUpdater.run(),
    /Bodoge ratings page contained cards without extractable ratings/,
  );
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});

test('RatingUpdater keeps existing rows when a later page has cards but no extractable titles', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  // Regression: earlier pages with real rows must not be written when a later
  // page matches card HTML yet yields zero titles (partial import).
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: ratingCard('カタン', '5') },
      { status: 200, body: cardWithoutJapaneseTitle('Ticket to Ride') },
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  assert.throws(
    () => context.RatingUpdater.run(),
    /Bodoge ratings page contained cards without extractable Japanese titles/,
  );
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});

test('RatingUpdater skips an excluded-title-only page after earlier importable rows', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  // Exclusions still extract a source title, so they must not be treated as the
  // same failure mode as missing Japanese titles mid-pagination.
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: ratingCard('カタン', '5') },
      { status: 200, body: ratingCard('ドミニオン：基本カードセット', '3') },
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  context.RatingUpdater.run();

  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), [
    {
      type: 'setValues',
      row: 2,
      column: 1,
      numRows: 1,
      numColumns: 2,
      values: [['カタン', '5']],
    },
  ]);
});

test('RatingUpdater keeps existing rows when every card title is excluded', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  // Extractable-but-excluded cards are not unextractable; reaching the empty
  // marker with zero importable rows must still abort instead of clearing.
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    responses: [
      { status: 200, body: ratingCard('ドミニオン：基本カードセット', '3') },
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  assert.throws(
    () => context.RatingUpdater.run(),
    /Bodoge ratings pages contained cards but yielded no importable titles/,
  );
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), []);
});

test('RatingUpdater writes ratings when cards fill exactly MAX_PAGE_COUNT pages then empty marker', () => {
  const ratingsSheet = createSheet('Ratings', 1);
  const maxPageCount = 100;
  const expectedRows = Array.from({ length: maxPageCount }, (_, index) => [
    `ゲーム${index + 1}`,
    '3',
  ]).sort(([firstTitle], [secondTitle]) =>
    firstTitle > secondTitle ? 1 : firstTitle < secondTitle ? -1 : 0,
  );
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    // A full card-page budget must still peek one page for Bodoge's empty
    // marker; otherwise a legitimate MAX_PAGE_COUNT-page collection would
    // throw and leave the previous Ratings snapshot forever.
    responses: [
      ...Array.from({ length: maxPageCount }, (_, index) => ({
        status: 200,
        body: ratingCard(`ゲーム${index + 1}`, '3'),
      })),
      { status: 200, body: emptyPlayedGamesPage() },
    ],
  });
  const context = loadRatingUpdater(sandbox);

  assert.equal(context.BODOGE_CONFIG.MAX_PAGE_COUNT, maxPageCount);
  context.RatingUpdater.run();

  assert.deepEqual(getCalls(ratingsSheet, 'setValues'), [
    {
      type: 'setValues',
      row: 2,
      column: 1,
      numRows: maxPageCount,
      numColumns: 2,
      values: expectedRows,
    },
  ]);
  assert.deepEqual(getCalls(ratingsSheet, 'clearContent'), []);
});

test('RatingUpdater keeps existing rows when Bodoge keeps returning rating cards past the page cap', () => {
  const ratingsSheet = createSheet('Ratings', 3);
  const maxPageCount = 100;
  const sandbox = createRatingSandbox({
    ratingsSheet,
    userId: 'user-1',
    // Card pages through the empty-marker peek mean runaway pagination; abort
    // without writing so a truncated import cannot replace a complete snapshot.
    responses: Array.from({ length: maxPageCount + 1 }, (_, index) => ({
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
