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
