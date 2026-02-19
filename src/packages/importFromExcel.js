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

function splitTextLines(text) {
  return String(text || '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line !== '');
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

    if (!text) {
      continue;
    }

    const baseId = toBaseId(code || name);
    const id = createStableId(baseId, takenIds);
    const lines = splitTextLines(text);
    const firstLine = lines[0] || '';
    const secondLine = lines[1] || '';
    const displayName = String(name || firstLine || code || id).trim();
    const shortText = String(secondLine || '').trim();

    packages.push({
      id,
      name,
      code,
      text,
      displayName,
      shortText,
      row: rowNumber,
    });
  }

  return packages.sort((a, b) => String(a.displayName || a.name || '').localeCompare(String(b.displayName || b.name || ''), 'de'));
}

module.exports = {
  importPackagesFromExcel,
};
