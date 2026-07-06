const assert = require('node:assert/strict');
const test = require('node:test');

const {
  createSandbox,
  createSheet,
  getCalls,
  loadRankings,
  rankingPage,
} = require('./helpers/appScriptHarness');

const SMALL_CARD_GAME_TAG_ID = 101;
const SMALL_CARD_GAME_TAG_NAME = 'カードゲーム';

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

function rankingTag() {
  return {
    id: SMALL_CARD_GAME_TAG_ID,
    name: SMALL_CARD_GAME_TAG_NAME,
  };
}

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

test('UpdateRankings skips clearContent for a header-only Rankings sheet and writes fetched games', () => {
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

  context.UpdateRankings.run();

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

test('UpdateRankings clears existing rows before writing fetched games', () => {
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

  context.UpdateRankings.run();

  assert.deepEqual(getCalls(rankingsSheet, 'clearContent'), [
    { type: 'clearContent', row: 2, column: 1, numRows: 3, numColumns: 15 },
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

test('UpdateRankings keeps existing rows when fetched HTML has no games', () => {
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

  context.UpdateRankings.run();

  assert.deepEqual(getCalls(rankingsSheet, 'clearContent'), []);
  assert.deepEqual(getCalls(rankingsSheet, 'setValues'), []);
});
