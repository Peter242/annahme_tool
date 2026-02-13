const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const { getSheetState, buildTodayPrefix } = require('../src/sheetState');

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
    maxLabNumber: 262031,
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
    maxLabNumber: 26204,
    maxOrderSeqToday: 0,
  });
});

test('getSheetState counts suffix variants for lab and order numbers', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');
  const now = new Date('2026-02-13T10:00:00Z');
  const todayPrefix = buildTodayPrefix(now);

  sheet.getCell('A1').value = '26203A';
  sheet.getCell('A2').value = '26203-1';
  sheet.getCell('A3').value = '26203 B';
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
