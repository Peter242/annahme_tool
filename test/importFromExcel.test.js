const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');
const fs = require('fs');
const os = require('os');
const ExcelJS = require('exceljs');
const { importPackagesFromExcel } = require('../src/packages/importFromExcel');

async function createWorkbook(filePath, sheetName, rows) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(sheetName);
  rows.forEach((row, index) => {
    sheet.getCell(index + 1, 1).value = row[0];
    sheet.getCell(index + 1, 2).value = row[1];
    sheet.getCell(index + 1, 3).value = row[2];
  });
  await workbook.xlsx.writeFile(filePath);
}

test('importPackagesFromExcel reads rows and creates stable ids', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-import-'));
  const excelPath = path.join(tmpDir, 'packages.xlsx');

  await createWorkbook(excelPath, 'Vorlagen', [
    ['Name', 'Code', 'Parametertext'],
    ['Nitrat', 'NO3', 'Text A'],
    ['Nitrat 2', 'NO3', 'Text B'],
    ['Sulfat', '', 'Text C'],
    ['', 'SKIP', 'Text D'],
    ['Leertext', 'SKIP2', ''],
  ]);

  const result = await importPackagesFromExcel(excelPath);

  assert.deepEqual(result, [
    {
      id: 'no3',
      name: 'Nitrat',
      code: 'NO3',
      text: 'Text A',
      displayName: 'Nitrat',
      shortText: '',
      row: 2,
    },
    {
      id: 'no3_2',
      name: 'Nitrat 2',
      code: 'NO3',
      text: 'Text B',
      displayName: 'Nitrat 2',
      shortText: '',
      row: 3,
    },
    {
      id: 'sulfat',
      name: 'Sulfat',
      code: '',
      text: 'Text C',
      displayName: 'Sulfat',
      shortText: '',
      row: 4,
    },
    {
      id: 'skip',
      name: '',
      code: 'SKIP',
      text: 'Text D',
      displayName: 'Text D',
      shortText: '',
      row: 5,
    },
  ]);
});

test('importPackagesFromExcel throws when sheet is missing', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-import-'));
  const excelPath = path.join(tmpDir, 'packages.xlsx');

  await createWorkbook(excelPath, 'AndereTabelle', [['Name', 'Code', 'Parametertext']]);

  await assert.rejects(
    importPackagesFromExcel(excelPath),
    /Sheet Vorlagen nicht gefunden/,
  );
});

test('importPackagesFromExcel reads updated excel content on every import', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-import-'));
  const excelPath = path.join(tmpDir, 'packages.xlsx');

  await createWorkbook(excelPath, 'Vorlagen', [
    ['Name', 'Code', 'Parametertext'],
    ['Nitrat', 'NO3', 'Text Alt'],
  ]);
  const first = await importPackagesFromExcel(excelPath);
  assert.equal(first[0].text, 'Text Alt');

  await createWorkbook(excelPath, 'Vorlagen', [
    ['Name', 'Code', 'Parametertext'],
    ['Nitrat', 'NO3', 'Text Neu TEST123'],
  ]);
  const second = await importPackagesFromExcel(excelPath);
  assert.equal(second[0].text, 'Text Neu TEST123');
});
