const path = require('path');
const ExcelJS = require('exceljs');

const DEFAULT_SHEET_NAME = 'Vorlagen';

function cellToString(cellOrValue) {
  if (cellOrValue === null || cellOrValue === undefined) {
    return '';
  }

  if (typeof cellOrValue === 'string') {
    return cellOrValue;
  }

  const value = typeof cellOrValue === 'object' && 'value' in cellOrValue ? cellOrValue.value : cellOrValue;

  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'string') {
    return value;
  }

  if (typeof value === 'object' && Array.isArray(value.richText)) {
    return value.richText.map((part) => (part && part.text ? String(part.text) : '')).join('');
  }

  if (typeof value === 'object' && value.text) {
    return String(value.text);
  }

  if (typeof cellOrValue === 'object' && cellOrValue.v !== undefined && cellOrValue.v !== null) {
    return String(cellOrValue.v);
  }

  return String(value).trim();
}

function toBaseId(raw) {
  const normalized = cellToString(raw).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  return normalized || 'pkg';
}

function createStableId(baseId, taken) {
  if (!taken.has(baseId)) {
    taken.add(baseId);
    return baseId;
  }

  let suffix = 2;
  while (taken.has(`${baseId}_${suffix}`)) {
    suffix += 1;
  }

  const candidate = `${baseId}_${suffix}`;
  taken.add(candidate);
  return candidate;
}

async function importPackagesFromExcel(excelPath, sheetName = DEFAULT_SHEET_NAME) {
  const workbook = new ExcelJS.Workbook();
  const absoluteExcelPath = path.resolve(excelPath);
  await workbook.xlsx.readFile(absoluteExcelPath);

  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw new Error(`Sheet ${sheetName} nicht gefunden`);
  }

  const packages = [];
  const takenIds = new Set();

  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const name = cellToString(row.getCell(1)).trim();
    const code = cellToString(row.getCell(2)).trim();
    const text = cellToString(row.getCell(3)).trim();

    if (!name || !text) {
      continue;
    }

    const baseId = toBaseId(code || name);
    const id = createStableId(baseId, takenIds);

    packages.push({
      id,
      name,
      code,
      text,
      row: rowNumber,
    });
  }

  return packages;
}

module.exports = {
  importPackagesFromExcel,
};
