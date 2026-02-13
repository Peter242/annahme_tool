const test = require('node:test');
const assert = require('node:assert/strict');
const { writeComTestCell } = require('../src/writers/comTestWriter');

test('writeComTestCell rejects on non-windows', async () => {
  if (process.platform === 'win32') {
    return;
  }

  await assert.rejects(
    writeComTestCell({ rootDir: '.', excelPath: 'x.xlsx', cellPath: '2026!Z1' }),
    /nur auf Windows/,
  );
});
