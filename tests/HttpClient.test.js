/**
 * HttpClient request-option tests.
 *
 * Confirms muteHttpExceptions is always enabled so callers can inspect non-2xx
 * statuses instead of receiving UrlFetchApp's default throw.
 */
const assert = require('node:assert/strict');
const test = require('node:test');

const { loadScripts } = require('./helpers/appScriptHarness');

/**
 * Loads HttpClient against a UrlFetchApp double that records fetch options.
 */
function loadHttpClient(sandbox) {
  return loadScripts(sandbox, [
    { path: 'src/infrastructure/HttpClient.ts', exports: ['HttpClient'] },
  ]);
}

/**
 * Creates a UrlFetchApp double that throws on non-2xx unless muted, matching
 * production Apps Script behavior.
 */
function createUrlFetchApp(status, body) {
  const fetchCalls = [];

  return {
    fetchCalls,
    UrlFetchApp: {
      fetch(url, options = {}) {
        fetchCalls.push({ url, options });
        if (
          (status < 200 || status >= 300) &&
          options.muteHttpExceptions !== true
        ) {
          throw new Error(`Request failed for ${url}: ${status}`);
        }

        return {
          getResponseCode() {
            return status;
          },
          getContentText() {
            return body;
          },
        };
      },
    },
  };
}

test('HttpClient.get mutes HTTP exceptions so callers can handle non-2xx statuses', () => {
  const { fetchCalls, UrlFetchApp } = createUrlFetchApp(503, 'unavailable');
  const context = loadHttpClient({ UrlFetchApp });

  const response = context.HttpClient.get('https://example.com/outage');

  assert.equal(response.getResponseCode(), 503);
  assert.equal(fetchCalls.length, 1);
  assert.equal(fetchCalls[0].options.muteHttpExceptions, true);
});

test('HttpClient.get preserves caller request options', () => {
  const { fetchCalls, UrlFetchApp } = createUrlFetchApp(200, 'ok');
  const context = loadHttpClient({ UrlFetchApp });

  context.HttpClient.get('https://example.com/resource', {
    headers: { Accept: 'application/json' },
    method: 'post',
  });

  assert.equal(fetchCalls[0].options.headers.Accept, 'application/json');
  assert.equal(fetchCalls[0].options.method, 'post');
  assert.equal(fetchCalls[0].options.muteHttpExceptions, true);
});

test('HttpClient.getWithOptionalBearerToken mutes HTTP exceptions and preserves auth', () => {
  const { fetchCalls, UrlFetchApp } = createUrlFetchApp(429, 'rate limited');
  const context = loadHttpClient({ UrlFetchApp });

  const response = context.HttpClient.getWithOptionalBearerToken(
    'https://example.com/thing',
    'secret-token',
  );

  assert.equal(response.getResponseCode(), 429);
  assert.equal(fetchCalls.length, 1);
  assert.equal(fetchCalls[0].options.muteHttpExceptions, true);
  assert.equal(
    fetchCalls[0].options.headers.Authorization,
    'Bearer secret-token',
  );
});
