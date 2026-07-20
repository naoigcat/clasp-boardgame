const assert = require('node:assert/strict');
const test = require('node:test');

const { loadScripts } = require('./helpers/appScriptHarness');

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
