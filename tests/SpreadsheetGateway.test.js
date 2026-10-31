const assert = require('node:assert/strict');
const test = require('node:test');
const {
  createSheet,
  getCalls,
  loadScripts,
} = require('./helpers/appScriptHarness');

function loadGateway() {
  return loadScripts({}, [
    { path: 'src/config/SheetSchema.ts', exports: [] },
    { path: 'src/shared/SpreadsheetUtils.ts', exports: [] },
    {
      path: 'src/infrastructure/SpreadsheetGateway.ts',
      exports: ['readSheetDataRows', 'writeSheetSnapshot'],
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

test('snapshots expand only the missing rows before writing at the requested columns', () => {
  const context = loadGateway();
  for (const count of [0, 998, 999, 1000, 1002]) {
    const sheet = createSheet('Ratings', 2, { maxRows: 1000 });
    const values = Array.from({ length: count }, () => ['title', 5]);
    context.writeSheetSnapshot(sheet, values, 2, 4);
    assert.deepEqual(
      getCalls(sheet, 'insertRowsAfter'),
      count > 999
        ? [
            {
              type: 'insertRowsAfter',
              afterPosition: 1000,
              howMany: count - 999,
            },
          ]
        : [],
    );
    const writes = getCalls(sheet, 'setValues');
    assert.equal(writes.length, count === 0 ? 0 : 1);
    if (count > 0) {
      assert.equal(writes[0].row, 2);
      assert.equal(writes[0].column, 4);
      assert.deepEqual(writes[0].values, values);
    }
  }
});

test('expansion and write failures never clear the previous snapshot', () => {
  const context = loadGateway();
  const values = Array.from({ length: 4 }, () => ['new title']);
  for (const failure of ['insert', 'write']) {
    const sheet = createSheet('Ratings', 3, {
      maxRows: 3,
      failOnSetValues: failure === 'write',
    });
    if (failure === 'insert') {
      sheet.insertRowsAfter = () => {
        throw new Error('insert failed');
      };
    }
    assert.throws(() => context.writeSheetSnapshot(sheet, values, 1), /failed/);
    assert.deepEqual(getCalls(sheet, 'clearContent'), []);
    assert.deepEqual(getCalls(sheet, 'setValues'), []);
  }
});
