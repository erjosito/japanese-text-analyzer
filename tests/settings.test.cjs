const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');

const app = fs.readFileSync(path.join(__dirname, '..', 'js', 'app.js'), 'utf8');
const settingsCode = app.slice(0, app.indexOf('// ---------- DOM refs ----------'));
const previousEndpoint = 'https://sommerlernplan-ai-a8fbd8e1.openai.azure.com';
const replacementEndpoint = 'https://foundry-japanese-text-analyzer-a8fbd8.openai.azure.com';

function load(raw) {
  const storage = new Map(raw === undefined ? [] : [['jta_settings_v1', raw]]);
  const context = vm.createContext({
    window: { JTA_I18N: { apply() {} } },
    localStorage: {
      getItem: (key) => storage.get(key) ?? null,
      setItem: (key, value) => storage.set(key, value),
    },
  });
  vm.runInContext(settingsCode, context);
  return {
    settings: JSON.parse(vm.runInContext('JSON.stringify(settings)', context)),
    save: () => vm.runInContext('saveSettings(settings)', context),
    storage,
  };
}

test('new installations use the dedicated Foundry endpoint', () => {
  const { settings } = load();
  assert.equal(settings.endpoint, replacementEndpoint);
  assert.equal(settings.deployment, 'japanese-text-analyzer');
});

test('retired default endpoint migrates without changing user preferences or authentication', () => {
  const saved = {
    endpoint: previousEndpoint,
    deployment: 'japanese-text-analyzer',
    apiVersion: '2024-10-21',
    tenantId: 'custom-tenant',
    clientId: 'custom-client',
    furigana: false,
    romaji: true,
  };
  const state = load(JSON.stringify(saved));
  assert.deepEqual(state.settings, { ...saved, endpoint: replacementEndpoint });
  state.save();
  assert.equal(JSON.parse(state.storage.get('jta_settings_v1')).endpoint, replacementEndpoint);
});

test('retired endpoint normalization handles trailing slashes, whitespace, and hostname case', () => {
  const { settings } = load(JSON.stringify({ endpoint: ` ${previousEndpoint.toUpperCase()}/// ` }));
  assert.equal(settings.endpoint, replacementEndpoint);
});

test('custom endpoints and custom deployments are preserved', () => {
  for (const saved of [
    { endpoint: 'https://custom.openai.azure.com', deployment: 'japanese-text-analyzer' },
    { endpoint: previousEndpoint, deployment: 'custom-deployment' },
    { endpoint: `${previousEndpoint}.example.com`, deployment: 'japanese-text-analyzer' },
  ]) {
    const { settings } = load(JSON.stringify(saved));
    assert.equal(settings.endpoint, saved.endpoint);
    assert.equal(settings.deployment, saved.deployment);
  }
});

test('already migrated settings stay unchanged', () => {
  const saved = load().settings;
  assert.deepEqual(load(JSON.stringify(saved)).settings, saved);
});

test('malformed stored JSON retains the existing default-settings behavior', () => {
  assert.deepEqual(load('{broken').settings, load().settings);
});
