const test = require('node:test');
const assert = require('node:assert/strict');
const { createImportPackagesHandler } = require('../src/packages/importRoute');

function createResponseCapture() {
  return {
    statusCode: 200,
    headers: {},
    payload: null,
    set(name, value) {
      this.headers[name] = value;
      return this;
    },
    status(code) {
      this.statusCode = code;
      return this;
    },
    json(data) {
      this.payload = data;
      return this;
    },
  };
}

test('/api/packages/import returns fresh packages from loader result', async () => {
  const callLog = [];
  const handler = createImportPackagesHandler({
    getConfig: () => ({ excelPath: 'x.xlsx' }),
    resolveExcelPath: (p) => `abs/${p}`,
    invalidatePackagesCache: () => callLog.push('invalidate'),
    importPackagesFromExcel: async () => [{ id: 'n', text: 'TEST123 neu' }],
    writePackages: (rows) => callLog.push(`write:${rows.length}`),
    readPackages: () => [{ id: 'n', text: 'TEST123 neu' }],
  });

  const res = createResponseCapture();
  await handler({}, res);

  assert.equal(res.statusCode, 200);
  assert.equal(res.headers['Cache-Control'], 'no-store');
  assert.deepEqual(res.payload, {
    ok: true,
    count: 1,
    packages: [{ id: 'n', text: 'TEST123 neu' }],
  });
  assert.deepEqual(callLog, ['invalidate', 'write:1', 'invalidate']);
});

test('/api/packages/import returns 400 on import error', async () => {
  const handler = createImportPackagesHandler({
    getConfig: () => ({ excelPath: 'x.xlsx' }),
    resolveExcelPath: (p) => `abs/${p}`,
    invalidatePackagesCache: () => {},
    importPackagesFromExcel: async () => { throw new Error('kaputt'); },
    writePackages: () => {},
    readPackages: () => [],
  });

  const res = createResponseCapture();
  await handler({}, res);

  assert.equal(res.statusCode, 400);
  assert.deepEqual(res.payload, {
    ok: false,
    message: 'kaputt',
  });
});
