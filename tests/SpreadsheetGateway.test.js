const assert = require('node:assert/strict');
const test = require('node:test');
const { loadScripts } = require('./helpers/appScriptHarness');

function loadGateway() {
  return loadScripts({}, [
    { path: 'src/config/SheetSchema.ts', exports: [] },
    {
      path: 'src/infrastructure/SpreadsheetGateway.ts',
      exports: ['readSheetDataRows'],
    },
  ]);
}

test('bounded reads return no rows without requesting an empty range', () => {
  const context = loadGateway();
  for (const lastRow of [0, 1]) {
    const sheet = {
      getLastRow: () => lastRow,
      getRange: () => assert.fail('An empty sheet has no data range'),
    };
    assert.equal(context.readSheetDataRows(sheet, 4).length, 0);
  }
});

test('bounded reads preserve values and honor the requested columns', () => {
  const context = loadGateway();
  const values = [
    ['first', null],
    ['second', 42],
  ];
  const calls = [];
  const sheet = {
    getLastRow: () => 3,
    getRange: (...args) => {
      calls.push(args);
      return { getValues: () => values };
    },
  };

  assert.equal(context.readSheetDataRows(sheet, 2, 3), values);
  assert.equal(context.readSheetDataRows(sheet, 2), values);
  assert.deepEqual(calls, [
    [2, 3, 2, 2],
    [2, 1, 2, 2],
  ]);
});
