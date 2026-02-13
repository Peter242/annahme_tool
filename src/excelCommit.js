const path = require('path');
const ExcelJS = require('exceljs');
const { getSheetState, buildTodayPrefix } = require('./sheetState');
const { readPackages } = require('./packages/store');

function pad2(value) {
  return String(value).padStart(2, '0');
}

function resolveYearSheetName(config, now) {
  const configured = (config.yearSheetName || '').trim();
  return configured || String(now.getFullYear());
}

function resolveAbsoluteExcelPath(excelPath, rootDir) {
  return path.isAbsolute(excelPath) ? excelPath : path.join(rootDir, excelPath);
}

function buildHeaderCellI(order) {
  const lines = [
    order.auftraggeberKurz || order.kunde || '',
    order.ansprechpartner || '',
    order.projektnummer || '',
    order.projektname || order.projekt || '',
    order.probenahmedatum || order.probenEingangDatum || '',
  ].filter((line) => String(line).trim() !== '');

  return lines.join('\n');
}

function buildHeaderCellJ(order, termin) {
  const sampleNotArrived = order.probeNochNichtDa || order.sampleNotArrived === true;
  const rawTermin = order.terminDatum || termin || '';
  const lines = [
    order.erfasstKuerzel ? `Erfasst: ${order.erfasstKuerzel}` : '',
    rawTermin ? `Termin: ${formatGermanTermin(rawTermin)}` : '',
    order.eilig ? 'Eilauftrag' : '',
    order.email ? `Email: ${order.email}` : '',
    sampleNotArrived ? 'Probe noch nicht da' : '',
  ].filter((line) => String(line).trim() !== '');

  return lines.join('\n');
}

function parseYmdDate(value) {
  if (typeof value !== 'string') {
    return null;
  }

  const trimmed = value.trim();
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trimmed);
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);
  const date = new Date(year, month, day);
  if (!Number.isFinite(date.getTime())) {
    return null;
  }

  if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) {
    return null;
  }

  return date;
}

function formatGermanTermin(value) {
  const parsed = parseYmdDate(value);
  if (!parsed) {
    return String(value);
  }

  const weekdayShort = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'][parsed.getDay()];
  const dd = String(parsed.getDate()).padStart(2, '0');
  const mm = String(parsed.getMonth() + 1).padStart(2, '0');
  const yyyy = parsed.getFullYear();
  return `${weekdayShort} ${dd}.${mm}.${yyyy}`;
}

function resolveParameterText(probe, packageById) {
  if (probe.packageId && packageById.has(probe.packageId)) {
    const template = packageById.get(probe.packageId);
    if (template && template.text) {
      return String(template.text).trim();
    }
  }

  return probe.parameterTextPreview ? String(probe.parameterTextPreview).trim() : '';
}

function resolveColumnG(probe) {
  if (probe.tiefeVolumen !== undefined && probe.tiefeVolumen !== null) {
    return String(probe.tiefeVolumen).trim();
  }

  if (typeof probe.volumen === 'number') {
    return String(probe.volumen);
  }

  return '';
}

function resolveColumnH(probe) {
  if (probe.materialGebinde) {
    return String(probe.materialGebinde).trim();
  }

  const material = probe.material ? String(probe.material).trim() : '';
  const gebinde = probe.gebinde ? String(probe.gebinde).trim() : '';
  const combined = `${material} ${gebinde}`.trim();

  return combined || (probe.matrixTyp ? String(probe.matrixTyp).trim() : '');
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function lineCountFromText(text) {
  if (!text) {
    return 1;
  }
  const normalized = String(text);
  const newlines = (normalized.match(/\n/g) || []).length;
  return 1 + newlines;
}

function buildProbeCellJ(probe) {
  const lines = [];

  if (probe.gewicht !== undefined && probe.gewicht !== null && String(probe.gewicht).trim() !== '') {
    const einheit = probe.gewichtEinheit ? ` ${String(probe.gewichtEinheit).trim()}` : '';
    lines.push(`Gewicht: ${String(probe.gewicht).trim()}${einheit}`);
  }

  if (probe.geruchAuffaelligkeit) {
    lines.push(`Geruch: ${String(probe.geruchAuffaelligkeit).trim()}`);
  }

  if (probe.bemerkung) {
    lines.push(`Bemerkung: ${String(probe.bemerkung).trim()}`);
  }

  return lines.join('\n');
}

async function appendOrderBlockToYearSheet(params) {
  const { config, rootDir, excelPath, order, termin, now = new Date(), packages } = params;
  const workbook = new ExcelJS.Workbook();
  const absoluteExcelPath = resolveAbsoluteExcelPath(excelPath, rootDir);
  await workbook.xlsx.readFile(absoluteExcelPath);

  const yearSheetName = resolveYearSheetName(config, now);
  const sheet = workbook.getWorksheet(yearSheetName);
  if (!sheet) {
    throw new Error(`Jahresblatt ${yearSheetName} nicht gefunden`);
  }

  const state = getSheetState(sheet, now);
  const appendRow = state.lastUsedRow + 2;

  const todayPrefix = buildTodayPrefix(now);
  const nextSeq = state.maxOrderSeqToday + 1;
  if (nextSeq > 99) {
    throw new Error(`Maximale Tagessequenz erreicht fuer Prefix ${todayPrefix}`);
  }

  const orderNo = `${todayPrefix}${pad2(nextSeq)}`;
  const startLabNo = state.maxLabNumber + 1;

  const packageSource = Array.isArray(packages) ? packages : readPackages();
  const packageById = new Map(packageSource.map((pkg) => [pkg.id, pkg]));

  sheet.getCell(appendRow, 1).value = orderNo;
  sheet.getCell(appendRow, 2).value = 'y';
  sheet.getCell(appendRow, 3).value = 'y';
  sheet.getCell(appendRow, 4).value = order.auftragsnotiz ? String(order.auftragsnotiz) : '';
  sheet.getCell(appendRow, 5).value = order.pbTyp || 'PB';
  sheet.getCell(appendRow, 9).value = buildHeaderCellI(order);
  sheet.getCell(appendRow, 10).value = buildHeaderCellJ(order, termin);

  order.proben.forEach((probe, index) => {
    const rowNumber = appendRow + 1 + index;
    const labNumber = startLabNo + index;
    const parameterText = resolveParameterText(probe, packageById);
    const probeCellJ = buildProbeCellJ(probe);
    const lineCount = Math.max(lineCountFromText(parameterText), lineCountFromText(probeCellJ));
    const rowHeight = clamp(lineCount * 15 + 5, 15, 300);
    const row = sheet.getRow(rowNumber);

    sheet.getCell(rowNumber, 1).value = labNumber;
    sheet.getCell(rowNumber, 4).value = parameterText;
    sheet.getCell(rowNumber, 4).alignment = { wrapText: true };
    sheet.getCell(rowNumber, 6).value = probe.probenbezeichnung ? String(probe.probenbezeichnung) : '';
    sheet.getCell(rowNumber, 7).value = resolveColumnG(probe);
    sheet.getCell(rowNumber, 8).value = resolveColumnH(probe);
    sheet.getCell(rowNumber, 10).value = probeCellJ;
    sheet.getCell(rowNumber, 10).alignment = { wrapText: true };
    row.height = rowHeight;
  });

  await workbook.xlsx.writeFile(absoluteExcelPath);

  return {
    orderNo,
    appendRow,
    startLabNo,
    yearSheetName,
  };
}

module.exports = {
  appendOrderBlockToYearSheet,
  resolveYearSheetName,
};
