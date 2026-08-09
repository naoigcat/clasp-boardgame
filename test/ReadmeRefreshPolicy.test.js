/**
 * README refresh-policy documentation tests.
 *
 * Keeps the BoardGameGeek stale-row prose aligned with
 * GAME_REFRESH_INTERVAL_DAYS versus GAME_FAILURE_BACKOFF_DAYS so operators do
 * not expect failed rows to wait a full successful refresh window.
 */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const { loadScripts } = require('./helpers/appScriptHarness');

/** English day words used by README for the current AppConfig values. */
const DAY_WORDS = {
  1: 'one',
  7: 'seven',
};

/**
 * Loads UPDATE_QUEUE_CONFIG from the shared Apps Script harness.
 */
function loadUpdateQueueConfig() {
  return loadScripts({}, [
    { path: 'src/config/AppConfig.ts', exports: ['UPDATE_QUEUE_CONFIG'] },
  ]).UPDATE_QUEUE_CONFIG;
}

/**
 * Returns the README day word for a configured day count.
 */
function dayWord(days) {
  const word = DAY_WORDS[days];
  assert.ok(
    word,
    `add a README day-word mapping for ${days} before changing AppConfig`,
  );
  return word;
}

/**
 * Collapses README wrapping so prose assertions ignore line breaks.
 */
function normalizeReadme(readme) {
  return readme.replace(/\s+/g, ' ');
}

test('README documents the short failure backoff separately from the success refresh window', () => {
  const readme = normalizeReadme(
    fs.readFileSync(path.join(__dirname, '..', 'README.md'), 'utf8'),
  );
  const config = loadUpdateQueueConfig();
  const successDays = config.GAME_REFRESH_INTERVAL_DAYS;
  const failureDays = config.GAME_FAILURE_BACKOFF_DAYS;
  const successWord = dayWord(successDays);
  const failureWord = dayWord(failureDays);

  assert.notEqual(
    failureDays,
    successDays,
    'expected failure backoff to differ from the success refresh interval',
  );

  // The inaccurate wording that treated failures like successes.
  assert.doesNotMatch(
    readme,
    /same refresh window later makes the failed row/i,
  );

  assert.match(
    readme,
    new RegExp(`successful update is ${successWord} days? old`, 'i'),
  );
  assert.match(
    readme,
    new RegExp(`failed attempt is ${failureWord} days? old`, 'i'),
  );
  assert.match(
    readme,
    new RegExp(
      `${failureWord}-day failure backoff .* shorter than the ${successWord}-day window`,
      'i',
    ),
  );
});

test('README documents that failure-backoff retries wait for the next manual Update', () => {
  const readme = normalizeReadme(
    fs.readFileSync(path.join(__dirname, '..', 'README.md'), 'utf8'),
  );

  assert.match(
    readme,
    /Retries after that backoff run on the next manual Update/i,
  );
  assert.match(
    readme,
    /Games phase usually finishes and removes the shared trigger before the backoff elapses/i,
  );
});
