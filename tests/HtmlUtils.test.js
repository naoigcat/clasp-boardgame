/**
 * HtmlUtils embedded-JSON extraction tests.
 *
 * Guards the bracket-matching parser against quoted brackets that would break a
 * greedy regular expression over Board Game Arena page source.
 */
const assert = require('node:assert/strict');
const test = require('node:test');

const { loadScripts } = require('./helpers/appScriptHarness');

/**
 * Loads the standalone HTML parser without any Apps Script dependencies.
 */
function loadHtmlUtils() {
  return loadScripts({}, [
    {
      path: 'src/shared/HtmlUtils.ts',
      exports: ['extractEmbeddedJsonArray'],
    },
  ]);
}

test('extractEmbeddedJsonArray respects brackets inside JSON strings', () => {
  const context = loadHtmlUtils();
  const page =
    '{"game_list":[{"name":"[bracket]","tags":[[101]]}],"game_tags":[]}';

  const extracted = context.extractEmbeddedJsonArray(page, 'game_list');

  assert.deepEqual(JSON.parse(JSON.stringify(extracted)), [
    { name: '[bracket]', tags: [[101]] },
  ]);
});

test('extractEmbeddedJsonArray reports a missing property clearly', () => {
  const context = loadHtmlUtils();

  assert.throws(
    () => context.extractEmbeddedJsonArray('{"game_list":[]}', 'game_tags'),
    /game_tags/,
  );
});
