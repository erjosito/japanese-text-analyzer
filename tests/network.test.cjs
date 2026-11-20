const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');

const app = fs.readFileSync(path.join(__dirname, '..', 'js', 'app.js'), 'utf8');
const code = app.slice(app.indexOf('async function postModelJson('), app.indexOf('async function analyze('));
const url = 'https://example.openai.azure.com/openai/deployments/test/chat/completions?api-version=test';
const secret = 'secret-token-not-for-diagnostics';
const body = { messages: [{ content: 'private-photo-content' }] };

function setup(fetch, online = true) {
  let now = 1000;
  const context = vm.createContext({
    fetch,
    URL,
    TypeError,
    navigator: { onLine: online },
    Date: { now: () => now },
    t: (key, params) => `${key}${params ? ' ' + JSON.stringify(params) : ''}`,
  });
  vm.runInContext(code, context);
  return { request: () => context.postModelJson(url, secret, body), tick: () => { now += 23000; } };
}

test('failed fetch reports safe diagnostics instead of generic Failed to fetch', async () => {
  let state;
  state = setup(async () => {
    state.tick();
    throw new TypeError('Failed to fetch');
  });
  await assert.rejects(state.request(), (error) => {
    assert.match(error.message, /error.networkFailure/);
    assert.match(error.message, /example\.openai\.azure\.com/);
    assert.match(error.message, /"seconds":23/);
    assert.match(error.message, new RegExp(`"size":${JSON.stringify(body).length}`));
    assert.match(error.message, /error.noHttpResponse/);
    assert.match(error.message, /error.browserOnline/);
    assert.ok(!error.message.includes(secret));
    assert.ok(!error.message.includes('private-photo-content'));
    assert.ok(!error.message.includes('/openai/deployments'));
    return true;
  });
});

test('offline status is included without claiming it is the root cause', async () => {
  const state = setup(async () => { throw new TypeError('Failed to fetch'); }, false);
  await assert.rejects(state.request(), /error.browserOffline/);
});

test('interrupted response is distinguished from no HTTP response', async () => {
  const state = setup(async () => ({
    ok: true, status: 200,
    json: async () => { throw new TypeError('Failed to fetch'); },
  }));
  await assert.rejects(state.request(), (error) => {
    assert.match(error.message, /error.responseInterrupted/);
    assert.match(error.message, /\\"status\\":200/);
    assert.ok(!error.message.includes('error.noHttpResponse'));
    return true;
  });
});

test('HTTP errors retain the Azure status and message', async () => {
  const state = setup(async () => ({ ok: false, status: 429, text: async () => 'rate limit' }));
  await assert.rejects(state.request(), (error) => {
    assert.match(error.message, /error.http/);
    assert.match(error.message, /"status":429/);
    assert.match(error.message, /rate limit/);
    assert.ok(!error.message.includes('error.networkFailure'));
    return true;
  });
});

test('successful JSON and authenticated request behavior are unchanged', async () => {
  const state = setup(async (actualUrl, request) => {
    assert.equal(actualUrl, url);
    assert.equal(request.method, 'POST');
    assert.equal(request.headers.Authorization, `Bearer ${secret}`);
    assert.equal(request.headers['Content-Type'], 'application/json');
    assert.equal(request.body, JSON.stringify(body));
    return { ok: true, status: 200, json: async () => ({ success: true }) };
  });
  assert.deepEqual(await state.request(), { success: true });
});

test('malformed response JSON is not mislabeled as a network failure', async () => {
  const malformed = new SyntaxError('Invalid JSON');
  const state = setup(async () => ({ ok: true, status: 200, json: async () => { throw malformed; } }));
  await assert.rejects(state.request(), (error) => error === malformed);
});
