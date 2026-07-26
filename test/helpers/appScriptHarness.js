/**
 * Shared doubles and loaders for Node-based Apps Script unit tests.
 *
 * Production sources are global TypeScript files without imports. Tests compile
 * them with `module: None` and evaluate them in a vm sandbox that supplies only
 * the Apps Script services each scenario needs.
 */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');

/**
 * Transpiles one non-module Apps Script source file and exposes selected names
 * to a test sandbox, mirroring Apps Script's global execution model.
 */
function compileSource(relativePath, exportNames) {
  const sourcePath = path.join(__dirname, '..', '..', relativePath);
  const source = fs.readFileSync(sourcePath, 'utf8');
  const output = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.None,
      // Match the production compiler so global script declarations behave alike.
      target: ts.ScriptTarget.ES5,
    },
  }).outputText;
  const exports = exportNames
    .map((name) => `globalThis.${name} = ${name};`)
    .join('\n');

  return `${output}\n${exports}`;
}

/**
 * Evaluates source files in order so a test can provide only the dependencies
 * required for the behavior under test.
 */
function loadScripts(sandbox, scripts) {
  const context = vm.createContext(sandbox);

  for (const script of scripts) {
    vm.runInContext(compileSource(script.path, script.exports), context, {
      filename: script.path,
    });
  }

  return context;
}

/**
 * Creates a minimal sheet double that records range mutations for assertions.
 */
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

/**
 * Creates a minimal Apps Script sandbox with queued HTTP responses.
 *
 * Responses are consumed in fetch order so tests can assert both the number of
 * upstream calls and the sheet writes that follow a successful parse.
 */
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
    Logger: {
      log() {},
    },
  };
}

/**
 * Loads the dependency chain required by RankingUpdater integration tests.
 */
function loadRankings(sandbox) {
  return loadScripts(sandbox, [
    { path: 'src/config/AppConfig.ts', exports: [] },
    { path: 'src/config/SheetSchema.ts', exports: [] },
    { path: 'src/shared/ErrorUtils.ts', exports: [] },
    { path: 'src/shared/HtmlUtils.ts', exports: [] },
    { path: 'src/infrastructure/HttpClient.ts', exports: ['HttpClient'] },
    { path: 'src/infrastructure/SpreadsheetGateway.ts', exports: [] },
    {
      path: 'src/services/RankingUpdater.ts',
      exports: ['RankingUpdater'],
    },
  ]);
}

/**
 * Builds the embedded JSON fragment expected from Board Game Arena's page.
 *
 * Only the property names and surrounding punctuation matter to the HTML
 * extractor; wrapping them in a larger page would not change parser behavior.
 */
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

/**
 * Returns recorded sheet operations of a requested type.
 */
function getCalls(sheet, type) {
  return sheet.calls.filter((call) => call.type === type);
}

module.exports = {
  createSandbox,
  createSheet,
  getCalls,
  loadScripts,
  loadRankings,
  rankingPage,
};
