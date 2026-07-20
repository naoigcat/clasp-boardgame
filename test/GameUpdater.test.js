const assert = require('node:assert/strict');
const test = require('node:test');

const { loadScripts } = require('./helpers/appScriptHarness');

/**
 * Creates a Games sheet double whose first source-link cell is empty.
 */
function createEmptyGamesSheet() {
  const writes = [];

  return {
    writes,
    getRange(a1NotationOrRow, column, numRows, numColumns) {
      if (typeof a1NotationOrRow === 'string') {
        return {
          getRichTextValues() {
            return [[null]];
          },
          getValues() {
            return a1NotationOrRow === '$B$2:$Z'
              ? [Array(25).fill('')]
              : [['']];
          },
        };
      }

      return {
        setValues(values) {
          writes.push({
            a1NotationOrRow,
            column,
            numRows,
            numColumns,
            values,
          });
        },
      };
    },
  };
}

/**
 * Loads the dependencies evaluated by GameUpdater before it begins API work.
 */
function loadGameUpdater(sandbox) {
  return loadScripts(sandbox, [
    { path: 'src/config/AppConfig.ts', exports: [] },
    { path: 'src/config/SheetSchema.ts', exports: [] },
    { path: 'src/infrastructure/SpreadsheetGateway.ts', exports: [] },
    { path: 'src/services/GameUpdater.ts', exports: ['GameUpdater'] },
  ]);
}

test('GameUpdater treats a null rich-text value as the end of the Games data', () => {
  const gamesSheet = createEmptyGamesSheet();
  const context = loadGameUpdater({
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'Games' ? gamesSheet : null;
          },
        };
      },
    },
  });

  assert.equal(context.GameUpdater.run(), false);
  assert.deepEqual(gamesSheet.writes, []);
});
