const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

function loadStoreWithPath(packagesPath) {
  const storeModulePath = require.resolve('../src/packages/store');
  delete require.cache[storeModulePath];
  process.env.PACKAGES_PATH = packagesPath;
  return require('../src/packages/store');
}

test('packages store supports cache invalidation and force reload', () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-packages-store-'));
  const packagesPath = path.join(tmpDir, 'packages.json');
  fs.writeFileSync(packagesPath, `${JSON.stringify([{ id: 'a', text: 'Alt' }], null, 2)}\n`, 'utf-8');

  try {
    const store = loadStoreWithPath(packagesPath);
    const first = store.readPackages();
    assert.equal(first[0].text, 'Alt');

    fs.writeFileSync(packagesPath, `${JSON.stringify([{ id: 'a', text: 'Neu' }], null, 2)}\n`, 'utf-8');

    store.invalidatePackagesCache();
    const second = store.readPackages({ forceReload: true });
    assert.equal(second[0].text, 'Neu');
  } finally {
    delete process.env.PACKAGES_PATH;
  }
});

test('packages store normalizes object text to string', () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-packages-store-'));
  const packagesPath = path.join(tmpDir, 'packages.json');
  fs.writeFileSync(packagesPath, `${JSON.stringify([{
    id: 'x1',
    name: 'X',
    text: { richText: [{ text: 'Zeile A' }, { text: '\nZeile B' }] },
  }], null, 2)}\n`, 'utf-8');
  try {
    const store = loadStoreWithPath(packagesPath);
    const items = store.readPackages({ forceReload: true });
    assert.equal(items[0].text, 'Zeile A\nZeile B');
  } finally {
    delete process.env.PACKAGES_PATH;
  }
});
