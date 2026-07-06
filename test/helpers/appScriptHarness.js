const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');

function compileSource(relativePath, exportNames) {
  const sourcePath = path.join(__dirname, '..', '..', relativePath);
  const source = fs.readFileSync(sourcePath, 'utf8');
  const output = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.None,
      target: ts.ScriptTarget.ES2020,
    },
  }).outputText;
  const exports = exportNames
    .map((name) => `globalThis.${name} = ${name};`)
    .join('\n');

  return `${output}\n${exports}`;
}

function loadScripts(sandbox, scripts) {
  const context = vm.createContext(sandbox);

  for (const script of scripts) {
    vm.runInContext(compileSource(script.path, script.exports), context, {
      filename: script.path,
    });
  }

  return context;
}

function createSheet(name, lastRow) {
  const calls = [];
  return {
    name,
    calls,
    getLastRow() {
      calls.push({ type: 'getLastRow' });
      return lastRow;
    },
    getRange(row, column, numRows, numColumns) {
      calls.push({ type: 'getRange', row, column, numRows, numColumns });
      if (numRows <= 0) {
        throw new Error(`Invalid range row count: ${numRows}`);
      }
      if (numColumns <= 0) {
        throw new Error(`Invalid range column count: ${numColumns}`);
      }

      return {
        clearContent() {
          calls.push({ type: 'clearContent', row, column, numRows, numColumns });
        },
        setValues(values) {
          calls.push({
            type: 'setValues',
            row,
            column,
            numRows,
            numColumns,
            values: JSON.parse(JSON.stringify(values)),
          });
        },
      };
    },
  };
}

function createSandbox({ sheets, responses }) {
  const responseQueue = [...responses];

  return {
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return sheets[name] || null;
          },
        };
      },
    },
    UrlFetchApp: {
      fetch(url) {
        const response = responseQueue.shift();
        assert.ok(response, `Unexpected fetch: ${url}`);
        assert.equal(typeof response.status, 'number');
        assert.equal(typeof response.body, 'string');

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
  };
}

function loadRankings(sandbox) {
  return loadScripts(sandbox, [
    { path: 'src/HttpClient.ts', exports: ['urlFetch'] },
    { path: 'src/UpdateRankings.ts', exports: ['UpdateRankings'] },
  ]);
}

function rankingPage(games, tags) {
  return [
    '{"game_list":',
    JSON.stringify(games),
    ',',
    '"game_tags":',
    JSON.stringify(tags),
    ',',
    '"top_tags":[]}',
  ].join('');
}

function getCalls(sheet, type) {
  return sheet.calls.filter((call) => call.type === type);
}

module.exports = {
  createSandbox,
  createSheet,
  getCalls,
  loadRankings,
  rankingPage,
};
