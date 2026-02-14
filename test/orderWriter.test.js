const test = require('node:test');
const assert = require('node:assert/strict');
const { writeOrderBlock } = require('../src/orderWriter');

test('writeOrderBlock rejects com backend on non-windows', async () => {
  if (process.platform === 'win32') {
    return;
  }

  await assert.rejects(
    writeOrderBlock({ backend: 'com', config: {}, rootDir: '.', excelPath: 'x.xlsx', order: {} }),
    /nur auf Windows/,
  );
});

test('writeOrderBlock accepts comExceljs alias for exceljs backend', async () => {
  await assert.rejects(
    writeOrderBlock({ backend: 'comExceljs', config: {}, rootDir: '.', excelPath: 'missing.xlsx', order: {} }),
    /no such file|ENOENT|File not found/i,
  );
});
