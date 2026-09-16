/**
 * Shared doubles and loaders for Node-based Apps Script unit tests.
 *
 * Production sources are global TypeScript files without imports. Tests compile
 * them with `module: None` and evaluate them in a vm sandbox that supplies only
 * the Apps Script services each scenario needs.
 */
const assert = require('node:assert/strict');
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const vm = require('node:vm');

/**
 * Compiles one non-module Apps Script source file and exposes selected names
 * to a test sandbox, mirroring Apps Script's global execution model. TypeScript
 * 7 no longer exposes the JavaScript transpiler API, so use its no-check CLI
 * for the same transpile-only behavior that the old API provided here.
 */
function compileSource(relativePath, exportNames) {
  const sourcePath = path.join(__dirname, '..', '..', relativePath);
  const outputDirectory = fs.mkdtempSync(
    path.join(os.tmpdir(), 'clasp-boardgame-ts-'),
  );

  let output;
  try {
    execFileSync(
      'tsc',
      [
        '--ignoreConfig',
        '--noCheck',
        '--target',
        'ES2015',
        '--module',
        'ES2015',
        '--outDir',
        outputDirectory,
        sourcePath,
      ],
      { stdio: 'pipe' },
    );

    const outputPath = path.join(
      outputDirectory,
      `${path.basename(relativePath, path.extname(relativePath))}.js`,
    );
    output = fs.readFileSync(outputPath, 'utf8');
  } finally {
    fs.rmSync(outputDirectory, { force: true, recursive: true });
  }

  // Keep top-level Apps Script declarations accessible as VM global properties.
  // ES5 transpilation did this implicitly; ES2015 preserves const/let instead.
  output = output.replace(/^(const|let)\s+/gm, 'var ');
  const exports = exportNames
    .map((name) => `globalThis.${name} = ${name};`)
    .join('\n');

  return `${output}\n${exports}`;
}

/**
 * Evaluates source files in order so a test can provide only the dependencies
 * required for the behavior under test.
 *
 * Host `Error` is shared into the sandbox so `instanceof Error` in production
 * helpers such as `getErrorMessage` stays true for doubles that throw from
 * outside the VM. Without that share, the VM's separate `Error` constructor
 * makes host-thrown errors fall through to `String(error)` and assert log
 * text that Apps Script would never emit.
 */
function loadScripts(sandbox, scripts) {
  if (!Object.hasOwn(sandbox, 'Error')) {
    sandbox.Error = Error;
  }

  const context = vm.createContext(sandbox);

  for (const script of scripts) {
    // An absolute filename lets Node associate VM coverage with the source file.
    vm.runInContext(compileSource(script.path, script.exports), context, {
      filename: path.join(__dirname, '..', '..', script.path),
    });
  }

  return context;
}

/**
 * Creates a minimal sheet double that records range mutations for assertions.
 *
 * `failOnSetValues` lets regression tests prove writers clear only after a
 * successful rewrite, matching Titles' write-then-trim surplus pattern.
 */
function createSheet(name, lastRow, { failOnSetValues = false } = {}) {
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
          calls.push({
            type: 'clearContent',
            row,
            column,
            numRows,
            numColumns,
          });
        },
        setValues(values) {
          if (failOnSetValues) {
            throw new Error('setValues failed');
          }
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
      fetch(url, options = {}) {
        const response = responseQueue.shift();
        assert.ok(response, `Unexpected fetch: ${url}`);
        assert.equal(typeof response.status, 'number');
        assert.equal(typeof response.body, 'string');

        // Mirror UrlFetchApp: non-2xx throws unless muteHttpExceptions is set.
        if (
          (response.status < 200 || response.status >= 300) &&
          options.muteHttpExceptions !== true
        ) {
          throw new Error(`Request failed for ${url}: ${response.status}`);
        }

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
    { path: 'src/shared/SpreadsheetUtils.ts', exports: [] },
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
