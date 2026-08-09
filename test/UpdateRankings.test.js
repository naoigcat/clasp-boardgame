/**
 * RankingUpdater catalog import tests.
 *
 * Exercises embedded JSON parsing, ignored implementation tags, player-count
 * boolean columns, and refusing to clear Rankings when the catalog is empty.
 */
const assert = require('node:assert/strict');
const test = require('node:test');

const {
  createSandbox,
  createSheet,
  getCalls,
  loadRankings,
  rankingPage,
} = require('./helpers/appScriptHarness');

/** Tag ID used to verify catalog-to-sheet tag resolution. */
const SMALL_CARD_GAME_TAG_ID = 101;
/** Japanese tag name paired with the test ID. */
const SMALL_CARD_GAME_TAG_NAME = 'カードゲーム';

/** Builds a representative Board Game Arena game payload. */
function rankingGame() {
  return {
    id: 1,
    name: 'azul',
    tags: [[SMALL_CARD_GAME_TAG_ID]],
    games_played: 123,
    average_duration: 30,
    default_num_players: 2,
    player_numbers: [2, 3, 4],
    watched: false,
  };
}

/** Builds a representative Board Game Arena tag payload. */
function rankingTag() {
  return {
    id: SMALL_CARD_GAME_TAG_ID,
    name: SMALL_CARD_GAME_TAG_NAME,
  };
}

/** Returns the exact Rankings row expected from the representative payload. */
function expectedRankingRow() {
  return [
    'https://ja.boardgamearena.com/gamepanel?game=azul',
    null,
    SMALL_CARD_GAME_TAG_NAME,
    123,
    30,
    2,
    true,
    true,
    true,
    false,
    false,
    false,
    false,
    false,
    false,
  ];
}

test('RankingUpdater skips clearContent for a header-only Rankings sheet and writes fetched games', () => {
  const rankingsSheet = createSheet('Rankings', 1);
  const sandbox = createSandbox({
    sheets: { Rankings: rankingsSheet },
    responses: [
      {
        status: 200,
        body: rankingPage([rankingGame()], [rankingTag()]),
      },
    ],
  });
  const context = loadRankings(sandbox);

  context.RankingUpdater.run();

  assert.deepEqual(getCalls(rankingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(rankingsSheet, 'setValues'), [
    {
      type: 'setValues',
      row: 2,
      column: 1,
      numRows: 1,
      numColumns: 15,
      values: [expectedRankingRow()],
    },
  ]);
});

test('RankingUpdater writes fetched games then clears surplus Rankings rows', () => {
  const rankingsSheet = createSheet('Rankings', 4);
  const sandbox = createSandbox({
    sheets: { Rankings: rankingsSheet },
    responses: [
      {
        status: 200,
        body: rankingPage([rankingGame()], [rankingTag()]),
      },
    ],
  });
  const context = loadRankings(sandbox);

  context.RankingUpdater.run();

  // Write first so a failed setValues cannot wipe Rankings, then trim only the
  // abandoned physical rows left by a shorter replacement snapshot.
  assert.deepEqual(
    rankingsSheet.calls
      .filter(
        (call) => call.type === 'setValues' || call.type === 'clearContent',
      )
      .map((call) => call.type),
    ['setValues', 'clearContent'],
  );
  assert.deepEqual(getCalls(rankingsSheet, 'clearContent'), [
    { type: 'clearContent', row: 3, column: 1, numRows: 2, numColumns: 15 },
  ]);
  assert.deepEqual(getCalls(rankingsSheet, 'setValues'), [
    {
      type: 'setValues',
      row: 2,
      column: 1,
      numRows: 1,
      numColumns: 15,
      values: [expectedRankingRow()],
    },
  ]);
});

test('RankingUpdater leaves Rankings intact when setValues fails', () => {
  const rankingsSheet = createSheet('Rankings', 4, { failOnSetValues: true });
  const sandbox = createSandbox({
    sheets: { Rankings: rankingsSheet },
    responses: [
      {
        status: 200,
        body: rankingPage([rankingGame()], [rankingTag()]),
      },
    ],
  });
  const context = loadRankings(sandbox);

  assert.throws(() => context.RankingUpdater.run(), /setValues failed/);

  // Clearing must not run after a write failure; otherwise the previous
  // Rankings snapshot would disappear until the next successful catalog fetch.
  assert.deepEqual(getCalls(rankingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(rankingsSheet, 'setValues'), []);
});

test('RankingUpdater keeps existing rows when fetched HTML has no games', () => {
  const rankingsSheet = createSheet('Rankings', 4);
  const sandbox = createSandbox({
    sheets: { Rankings: rankingsSheet },
    responses: [
      {
        status: 200,
        body: rankingPage([], [rankingTag()]),
      },
    ],
  });
  const context = loadRankings(sandbox);

  context.RankingUpdater.run();

  assert.deepEqual(getCalls(rankingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(rankingsSheet, 'setValues'), []);
});

test('RankingUpdater escapes formula-like tag names before setValues', () => {
  const formulaTagName = '=HYPERLINK("https://evil.example","x")';
  const rankingsSheet = createSheet('Rankings', 1);
  const sandbox = createSandbox({
    sheets: { Rankings: rankingsSheet },
    responses: [
      {
        status: 200,
        body: rankingPage(
          [rankingGame()],
          [{ id: SMALL_CARD_GAME_TAG_ID, name: formulaTagName }],
        ),
      },
    ],
  });
  const context = loadRankings(sandbox);

  context.RankingUpdater.run();

  const written = getCalls(rankingsSheet, 'setValues')[0].values[0];
  // Tag names are external text; a leading apostrophe keeps Sheets from
  // executing them as formulas when the workbook is shared.
  assert.equal(written[2], `'${formulaTagName}`);
  assert.equal(written[0], 'https://ja.boardgamearena.com/gamepanel?game=azul');
});
