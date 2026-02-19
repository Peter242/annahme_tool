const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const { getSheetState, buildTodayPrefix, extractOrderCore } = require('../src/sheetState');

test('extractOrderCore handles suffixes and rejects invalid values', () => {
  const cases = [
    { input: '140226801', expectedCore: '140226801', expectedSeq: 1 },
    { input: '140226801-1', expectedCore: '140226801', expectedSeq: 1 },
    { input: '140226801-B', expectedCore: '140226801', expectedSeq: 1 },
    { input: '140226802', expectedCore: '140226802', expectedSeq: 2 },
    { input: 'foo140226803', expectedCore: null, expectedSeq: null },
    { input: '14022680', expectedCore: null, expectedSeq: null },
    { input: '140226901', expectedCore: null, expectedSeq: null },
  ];

  for (const c of cases) {
    const core = extractOrderCore(c.input);
    assert.equal(core, c.expectedCore);
    if (core) {
      assert.equal(Number.parseInt(core.slice(-2), 10), c.expectedSeq);
    }
  }
});

test('getSheetState returns lastUsedRow, maxLabNumber and maxOrderSeqToday', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');

  const now = new Date('2026-02-13T10:00:00Z');
  const todayPrefix = buildTodayPrefix(now);

  sheet.getCell('A1').value = 'Titel';
  sheet.getCell('J2').value = 'x';
  sheet.getCell('A3').value = `${todayPrefix}01 Header`;
  sheet.getCell('A4').value = '26203';
  sheet.getCell('A5').value = '262031 Zusatz';
  sheet.getCell('A6').value = `${todayPrefix}09 mehr`;
  sheet.getCell('A7').value = '130226801';
  sheet.getCell('A8').value = 'ABC';
  sheet.getCell('A9').value = '1234';

  const state = getSheetState(sheet, now);

  assert.deepEqual(state, {
    lastUsedRow: 9,
    maxLabNumber: 26203,
    maxOrderSeqToday: 9,
  });
});

test('getSheetState ignores non matching rows and returns 0 defaults', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');

  sheet.getCell('A1').value = '';
  sheet.getCell('B2').value = null;
  sheet.getCell('C3').value = ' ';

  const state = getSheetState(sheet, new Date('2026-02-13T10:00:00Z'));

  assert.deepEqual(state, {
    lastUsedRow: 0,
    maxLabNumber: 0,
    maxOrderSeqToday: 0,
  });
});

test('getSheetState reads richText values in column A', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');

  sheet.getCell('A1').value = {
    richText: [{ text: '26204' }, { text: ' Probe' }],
  };

  const state = getSheetState(sheet, new Date('2026-02-13T10:00:00Z'));

  assert.deepEqual(state, {
    lastUsedRow: 1,
    maxLabNumber: 0,
    maxOrderSeqToday: 0,
  });
});

test('getSheetState parses leading lab digits with suffixes and parses order core suffixes', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');
  const now = new Date('2026-02-13T10:00:00Z');
  const todayPrefix = buildTodayPrefix(now);

  sheet.getCell('A1').value = '26203A';
  sheet.getCell('A2').value = '26203-1';
  sheet.getCell('A3').value = '26203';
  sheet.getCell('A4').value = `${todayPrefix}01A`;
  sheet.getCell('A5').value = `${todayPrefix}01-1`;
  sheet.getCell('A6').value = `${todayPrefix}02 Zusatz`;

  const state = getSheetState(sheet, now);

  assert.deepEqual(state, {
    lastUsedRow: 6,
    maxLabNumber: 26203,
    maxOrderSeqToday: 2,
  });
});

test('getSheetState parses lastLabNo from mixed suffix inputs', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');

  sheet.getCell('A1').value = '10040B';
  sheet.getCell('A2').value = '10040-1';
  sheet.getCell('A3').value = '10039';
  sheet.getCell('A4').value = 'dito';
  sheet.getCell('A5').value = '';
  sheet.getCell('A6').value = 'ABC';

  const state = getSheetState(sheet, new Date('2026-02-15T10:00:00Z'));
  assert.equal(state.maxLabNumber, 10040);
  assert.equal(state.maxLabNumber + 1, 10041);
});

test('getSheetState ignores previous-day order numbers but parses suffixes for today', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');
  const now = new Date('2026-02-14T10:00:00Z');
  const todayPrefix = buildTodayPrefix(now);

  sheet.getCell('A1').value = '130226805';
  sheet.getCell('A2').value = `${todayPrefix}01-1`;
  sheet.getCell('A3').value = `${todayPrefix}01-B`;

  const state = getSheetState(sheet, now);
  assert.equal(state.maxOrderSeqToday, 1);
});

test('getSheetState computes next order sequence robustly with suffix variants', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');
  const now = new Date('2026-02-14T10:00:00Z');
  const todayPrefix = buildTodayPrefix(now);

  sheet.getCell('A1').value = `${todayPrefix}01-1`;
  sheet.getCell('A2').value = `${todayPrefix}01-B`;
  sheet.getCell('A3').value = `${todayPrefix}01A`;
  sheet.getCell('A4').value = `${todayPrefix}02`;
  sheet.getCell('A5').value = '130226805';

  const state = getSheetState(sheet, now);
  assert.equal(state.maxOrderSeqToday, 2);
});

test('getSheetState treats duplicated core variants as one sequence value', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');
  const now = new Date('2026-02-14T10:00:00Z');

  sheet.getCell('A1').value = '140226801';
  sheet.getCell('A2').value = '140226801-1';
  sheet.getCell('A3').value = '140226801-B';

  const state = getSheetState(sheet, now);
  assert.equal(state.maxOrderSeqToday, 1);
});

test('getSheetState derives maxOrderSeqToday from numeric base with suffixes', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');
  const now = new Date('2026-02-15T10:00:00Z');

  sheet.getCell('A1').value = '150226801-1';
  sheet.getCell('A2').value = '150226801-B';
  sheet.getCell('A3').value = '140226899';

  const state = getSheetState(sheet, now);
  assert.equal(state.maxOrderSeqToday, 1);
});

test('getSheetState lab scan ignores order numbers and non-lab tokens', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');

  sheet.getCell('A3').value = '150226801';
  sheet.getCell('A4').value = '10039';
  sheet.getCell('A5').value = '10040B';
  sheet.getCell('A6').value = 'dito';
  sheet.getCell('A7').value = '150226802-1';

  const state = getSheetState(sheet, new Date('2026-02-15T10:00:00Z'));
  assert.equal(state.maxLabNumber, 10040);
  assert.equal(state.maxLabNumber + 1, 10041);
});

test('getSheetState ignores foreign order numbers and 8-digit foreign lab numbers', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');
  const now = new Date('2026-02-15T10:00:00Z');
  const todayPrefix = buildTodayPrefix(now);

  sheet.getCell('A1').value = `${todayPrefix}01`;
  sheet.getCell('A2').value = '150226701';
  sheet.getCell('A3').value = '25160467';
  sheet.getCell('A4').value = '25160462';
  sheet.getCell('A5').value = '10040B';
  sheet.getCell('A6').value = '10039';

  const state = getSheetState(sheet, now);
  assert.equal(state.maxOrderSeqToday, 1);
  assert.equal(state.maxLabNumber, 10040);
});
