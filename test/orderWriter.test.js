const test = require('node:test');
const assert = require('node:assert/strict');
const { writeOrderBlock } = require('../src/orderWriter');
const {
  validateComRows,
  buildComHeaderRowPreview,
  buildComSampleRowPreview,
} = require('../src/writers/comWriter');

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

test('com preflight row builder creates 10 columns for header and sample rows', () => {
  const header = buildComHeaderRowPreview({
    auftragsnotiz: 'Notiz',
    kopfBemerkung: 'Hinweis Kopf',
    pbTyp: 'PB',
  });
  const sample = buildComSampleRowPreview({
    probenbezeichnung: 'Probe 1',
    matrixTyp: 'Boden',
    tiefeOderVolumen: '0-20 cm',
    parameterTextPreview: 'Text D',
  }, 26204);

  assert.equal(Array.isArray(header), true);
  assert.equal(Array.isArray(sample), true);
  assert.equal(header.length, 10);
  assert.equal(sample.length, 10);
  assert.equal(header[3], '');
  assert.equal(sample[9], '');
  assert.doesNotThrow(() => validateComRows([header, sample], 'test'));
});

test('com preflight sample preview renders probeJ from present fields only', () => {
  const gewichtOnly = buildComSampleRowPreview({ gewicht: 2 }, 26204);
  const gewichtGeruch = buildComSampleRowPreview({ gewicht: 2, geruch: 'muffig' }, 26205);
  const bemerkungOnly = buildComSampleRowPreview({ bemerkung: 'wenig Material' }, 26206);

  assert.equal(gewichtOnly[9], 'Gewicht: 2 kg');
  assert.equal(gewichtGeruch[9], 'Gewicht: 2 kg; Geruch: muffig');
  assert.equal(bemerkungOnly[9], 'wenig Material');
});

test('com preflight validation fails on malformed row shape', () => {
  assert.throws(
    () => validateComRows(['A|B|C'], 'test'),
    /expectedCols=10/i,
  );
});
