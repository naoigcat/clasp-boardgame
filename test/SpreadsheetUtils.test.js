/**
 * SpreadsheetUtils formula-injection escape tests.
 *
 * Guards the setValues helper that prefixes formula-triggering characters so
 * external titles and tags cannot become executable sheet formulas.
 */
const assert = require('node:assert/strict');
const test = require('node:test');

const { loadScripts } = require('./helpers/appScriptHarness');

/**
 * Loads the standalone sheet-value escape helpers.
 */
function loadSpreadsheetUtils() {
  return loadScripts({}, [
    {
      path: 'src/shared/SpreadsheetUtils.ts',
      exports: ['escapeSheetCellValue', 'escapeSheetValues'],
    },
  ]);
}

test('escapeSheetCellValue prefixes formula-triggering characters', () => {
  const context = loadSpreadsheetUtils();

  assert.equal(context.escapeSheetCellValue('=1+1'), "'=1+1");
  assert.equal(context.escapeSheetCellValue('+cmd'), "'+cmd");
  assert.equal(context.escapeSheetCellValue('-1'), "'-1");
  assert.equal(context.escapeSheetCellValue('@SUM(A1)'), "'@SUM(A1)");
  assert.equal(context.escapeSheetCellValue('\tTAB'), "'\tTAB");
  assert.equal(context.escapeSheetCellValue('\rCR'), "'\rCR");
});

test('escapeSheetCellValue prefixes digit strings with a leading zero', () => {
  const context = loadSpreadsheetUtils();

  assert.equal(context.escapeSheetCellValue('01'), "'01");
  assert.equal(context.escapeSheetCellValue('0123'), "'0123");
  assert.equal(context.escapeSheetCellValue('00'), "'00");
});

test('escapeSheetCellValue leaves safe values unchanged', () => {
  const context = loadSpreadsheetUtils();

  assert.equal(context.escapeSheetCellValue('カタン'), 'カタン');
  assert.equal(
    context.escapeSheetCellValue('https://example.com'),
    'https://example.com',
  );
  assert.equal(context.escapeSheetCellValue(''), '');
  assert.equal(context.escapeSheetCellValue('0'), '0');
  assert.equal(context.escapeSheetCellValue('0a'), '0a');
  assert.equal(context.escapeSheetCellValue(4), 4);
  assert.equal(context.escapeSheetCellValue(true), true);
  assert.equal(context.escapeSheetCellValue(null), null);
});

test('escapeSheetValues escapes only string cells in each row', () => {
  const context = loadSpreadsheetUtils();

  assert.deepEqual(
    context.escapeSheetValues([
      ['=HYPERLINK("x")', null, 'カードゲーム', 12, true],
      ['safe', '', '-leading', 0, false],
    ]),
    [
      ['\'=HYPERLINK("x")', null, 'カードゲーム', 12, true],
      ['safe', '', "'-leading", 0, false],
    ],
  );
});
