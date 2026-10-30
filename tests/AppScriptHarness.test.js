const assert = require('node:assert/strict');
const test = require('node:test');

const { createClockDate, loadScripts } = require('./helpers/appScriptHarness');

test('cached sources still execute in independent sandboxes with per-load exports', () => {
  const script = {
    path: 'src/infrastructure/HttpClient.ts',
    exports: ['HttpClient'],
  };
  const first = loadScripts({ UrlFetchApp: { fetch: () => 'first' } }, [
    script,
  ]);
  first.HttpClient.requireSuccess = () => 'changed';

  const second = loadScripts({ UrlFetchApp: { fetch: () => 'second' } }, [
    script,
  ]);
  assert.equal(first.HttpClient.get('https://example.com'), 'first');
  assert.equal(second.HttpClient.get('https://example.com'), 'second');
  assert.equal(
    second.HttpClient.requireSuccess({ getResponseCode: () => 200 }),
    undefined,
  );

  const withoutExports = loadScripts({}, [{ ...script, exports: [] }]);
  assert.equal(withoutExports.HttpClient, undefined);
});

test('test clocks advance independently and preserve Date identity', () => {
  const first = createClockDate(1000);
  const second = createClockDate(2000);
  first.clock.nowMs += 500;

  assert.equal(first.ClockDate.now(), 1500);
  assert.equal(new first.ClockDate().getTime(), 1500);
  assert.equal(second.ClockDate.now(), 2000);
  assert.ok(new first.ClockDate() instanceof Date);
  assert.ok(new Date() instanceof first.ClockDate);
  assert.equal(new first.ClockDate(0).getTime(), 0);
});
