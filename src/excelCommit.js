const path = require('path');
const ExcelJS = require('exceljs');
const { getSheetState, buildTodayPrefix } = require('./sheetState');
const { readPackages } = require('./packages/store');
const { renderColumnHFromProbe, renderContainersSummary } = require('./containers');
const { buildProbeJ } = require('./probeJ');

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

function rowHasContentAToJ(sheet, rowNumber) {
  for (let column = 1; column <= 10; column += 1) {
    const raw = sheet.getCell(rowNumber, column).value;
    if (raw !== null && raw !== undefined && String(raw).trim() !== '') {
      return true;
    }
  }
  return false;
}

function isSameCalendarDay(dateA, dateB) {
  if (!(dateA instanceof Date) || !(dateB instanceof Date)) {
    return false;
  }
  return dateA.getFullYear() === dateB.getFullYear()
    && dateA.getMonth() === dateB.getMonth()
    && dateA.getDate() === dateB.getDate();
}

function buildHeaderCellI(order, now = new Date()) {
  const kunde = order.auftraggeberKurz || order.kunde || '';
  const projekt = order.projektName || order.projektname || order.projekt || '';
  const projektNrLine = order.projektnummer ? `Projekt Nr: ${order.projektnummer}` : '';
  const projektLine = projekt ? `Projekt: ${projekt}` : '';
  const rawProbenahme = order.probenahmedatum || '';
  const probenahme = formatGermanDateOnly(rawProbenahme);
  const probenahmeLine = probenahme ? `Probenahme: ${probenahme}` : '';
  const eingangParsed = parseFlexibleDate(order.probenEingangDatum || '');
  const eingangLine = (eingangParsed && !isSameCalendarDay(eingangParsed, now))
    ? `Eingangsdatum: ${formatGermanDateOnly(eingangParsed)}`
    : '';
  const transportRaw = String(order.probentransport || '').trim().toUpperCase();
  const transportLine = (transportRaw === 'CUA' || transportRaw === 'AG') ? `Transport: ${transportRaw}` : '';
  const lines = [
    kunde,
    order.ansprechpartner || '',
    projektNrLine,
    projektLine,
    probenahmeLine,
    eingangLine,
    transportLine,
  ].filter((line) => String(line).trim() !== '');

  return lines.join('\n');
}

function shouldWriteAddressBlock(config) {
  return config?.excelWriteAddressBlock !== false;
}

function buildHeaderCellJ(order, termin, config = {}) {
  const rawTermin = order.terminDatum || termin || '';
  const kuerzel = order.kuerzel || order.erfasstKuerzel || '';
  const formattedTermin = formatGermanTermin(rawTermin);
  const adresseBlock = String(order.adresseBlock || '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line !== '')
    .join('\n');
  const kopfBemerkung = String(order.kopfBemerkung || '').trim();
  const firstLineParts = [];
  if (String(kuerzel).trim()) {
    firstLineParts.push(String(kuerzel).trim());
  }
  if (order.eilig) {
    firstLineParts.push('EILIG');
  }
  if (formattedTermin) {
    firstLineParts.push(`Termin: ${formattedTermin}`);
  }
  const lines = [];
  const terminLine = firstLineParts.join(' ').trim();
  if (terminLine) {
    lines.push(terminLine);
  }
  if (kopfBemerkung) {
    lines.push(kopfBemerkung);
  }
  if (shouldWriteAddressBlock(config) && adresseBlock) {
    lines.push(adresseBlock);
  }

  return lines.join('\n');
}

function parseFlexibleDate(value) {
  if (typeof value !== 'string' && !(value instanceof Date)) {
    return null;
  }

  if (value instanceof Date) {
    return Number.isFinite(value.getTime()) ? value : null;
  }

  const trimmed = value.trim();
  const ymd = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trimmed);
  if (ymd) {
    const year = Number(ymd[1]);
    const month = Number(ymd[2]) - 1;
    const day = Number(ymd[3]);
    const date = new Date(year, month, day);
    if (Number.isFinite(date.getTime()) && date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
      return date;
    }
  }

  const dmy = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(trimmed);
  if (dmy) {
    const day = Number(dmy[1]);
    const month = Number(dmy[2]) - 1;
    const year = Number(dmy[3]);
    const date = new Date(year, month, day);
    if (Number.isFinite(date.getTime()) && date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
      return date;
    }
  }

  const parsed = new Date(trimmed);
  return Number.isFinite(parsed.getTime()) ? parsed : null;
}

function formatGermanDateOnly(value) {
  const parsed = parseFlexibleDate(value);
  if (!parsed) {
    return '';
  }
  const dd = String(parsed.getDate()).padStart(2, '0');
  const mm = String(parsed.getMonth() + 1).padStart(2, '0');
  const yyyy = parsed.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

function formatGermanTermin(value) {
  const parsed = parseFlexibleDate(value);
  if (!parsed) {
    return '';
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
  if (probe.tiefeOderVolumen !== undefined && probe.tiefeOderVolumen !== null && String(probe.tiefeOderVolumen).trim() !== '') {
    return String(probe.tiefeOderVolumen).trim();
  }

  if (probe.tiefeVolumen !== undefined && probe.tiefeVolumen !== null) {
    return String(probe.tiefeVolumen).trim();
  }

  if (typeof probe.volumen === 'number') {
    return String(probe.volumen);
  }

  return '';
}

function resolveColumnH(probe, config) {
  if (probe.materialGebinde) {
    return String(probe.materialGebinde).trim();
  }
  return renderColumnHFromProbe(probe, { config });
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

function buildHeaderRowValues(params = {}) {
  const {
    orderNo = 'ORDER_NO',
    order = {},
    termin = null,
    config = {},
    now = new Date(),
  } = params;
  const quickConfig = {
    quickContainerPlastic: Array.isArray(config.quickContainerPlastic) ? config.quickContainerPlastic : undefined,
    quickContainerGlass: Array.isArray(config.quickContainerGlass) ? config.quickContainerGlass : undefined,
  };
  const headerSummary = order.sameContainersForAll
    ? renderContainersSummary(order.headerContainers, { config: quickConfig })
    : '';

  return [
    orderNo,
    'y',
    'y',
    '',
    order.pbTyp || 'PB',
    '',
    '',
    headerSummary,
    buildHeaderCellI(order, now),
    buildHeaderCellJ(order, termin, config),
  ];
}

async function appendOrderBlockToYearSheet(params) {
  const { config, rootDir, excelPath, order, termin, cacheHint = null, now = new Date(), packages } = params;
  const probes = Array.isArray(order.proben) ? order.proben : [];
  const workbook = new ExcelJS.Workbook();
  const absoluteExcelPath = resolveAbsoluteExcelPath(excelPath, rootDir);
  await workbook.xlsx.readFile(absoluteExcelPath);

  const yearSheetName = resolveYearSheetName(config, now);
  const sheet = workbook.getWorksheet(yearSheetName);
  if (!sheet) {
    throw new Error(`Jahresblatt ${yearSheetName} nicht gefunden`);
  }

  const todayPrefix = buildTodayPrefix(now);
  let state = null;
  let appendRow = null;
  let nextSeq = null;
  let maxOrderSeqToday = null;
  let startLabNo = null;

  if (cacheHint && Number.isInteger(cacheHint.appendRow) && cacheHint.appendRow > 0) {
    const hintAppendRow = cacheHint.appendRow;
    const rowBusy = rowHasContentAToJ(sheet, hintAppendRow);
    const prefixMatch = String(cacheHint.todayPrefix || '') === todayPrefix;
    if (!rowBusy && prefixMatch) {
      appendRow = hintAppendRow;
      maxOrderSeqToday = Number.isInteger(cacheHint.maxOrderSeqToday) ? cacheHint.maxOrderSeqToday : 0;
      nextSeq = Number.isInteger(cacheHint.nextSeq) ? cacheHint.nextSeq : (maxOrderSeqToday + 1);
      startLabNo = Number.isInteger(cacheHint.startLabNo) ? cacheHint.startLabNo : 10000;
    }
  }

  if (!Number.isInteger(appendRow) || !Number.isInteger(nextSeq) || !Number.isInteger(startLabNo)) {
    state = getSheetState(sheet, now);
    appendRow = state.lastUsedRow + 2;
    maxOrderSeqToday = state.maxOrderSeqToday;
    nextSeq = state.maxOrderSeqToday + 1;
    startLabNo = Math.max(state.maxLabNumber, 9999) + 1;
  }

  if (nextSeq > 99) {
    throw new Error(`Maximale Tagessequenz erreicht fuer Prefix ${todayPrefix}`);
  }

  const orderNo = `${todayPrefix}${pad2(nextSeq)}`;
  const sampleNos = [];

  const packageSource = Array.isArray(packages) ? packages : readPackages();
  const packageById = new Map(packageSource.map((pkg) => [pkg.id, pkg]));

  sheet.getCell(appendRow, 1).value = orderNo;
  sheet.getCell(appendRow, 2).value = 'y';
  sheet.getCell(appendRow, 3).value = 'y';
  sheet.getCell(appendRow, 4).value = '';
  sheet.getCell(appendRow, 5).value = order.pbTyp || 'PB';
  const quickConfig = {
    quickContainerPlastic: Array.isArray(config.quickContainerPlastic) ? config.quickContainerPlastic : undefined,
    quickContainerGlass: Array.isArray(config.quickContainerGlass) ? config.quickContainerGlass : undefined,
  };
  const headerSummary = order.sameContainersForAll
    ? renderContainersSummary(order.headerContainers, { config: quickConfig })
    : '';
  sheet.getCell(appendRow, 8).value = headerSummary;
  sheet.getCell(appendRow, 9).value = buildHeaderCellI(order, now);
  sheet.getCell(appendRow, 10).value = buildHeaderCellJ(order, termin, config);

  probes.forEach((probe, index) => {
    const rowNumber = appendRow + 1 + index;
    const labNumber = startLabNo + index;
    sampleNos.push(labNumber);
    const parameterText = resolveParameterText(probe, packageById);
    const probeCellJ = buildProbeJ(probe);
    const lineCount = Math.max(lineCountFromText(parameterText), lineCountFromText(probeCellJ));
    const rowHeight = clamp(lineCount * 15 + 5, 15, 300);
    const row = sheet.getRow(rowNumber);

    sheet.getCell(rowNumber, 1).value = labNumber;
    sheet.getCell(rowNumber, 4).value = parameterText;
    sheet.getCell(rowNumber, 4).alignment = { wrapText: true };
    sheet.getCell(rowNumber, 6).value = probe.probenbezeichnung ? String(probe.probenbezeichnung) : '';
    sheet.getCell(rowNumber, 7).value = resolveColumnG(probe);
    sheet.getCell(rowNumber, 8).value = resolveColumnH(probe, quickConfig);
    sheet.getCell(rowNumber, 10).value = probeCellJ;
    sheet.getCell(rowNumber, 10).alignment = { wrapText: true };
    row.height = rowHeight;
  });

  await workbook.xlsx.writeFile(absoluteExcelPath);

  const firstSampleNo = sampleNos.length > 0 ? sampleNos[0] : null;
  const lastSampleNo = sampleNos.length > 0 ? sampleNos[sampleNos.length - 1] : null;
  const endRow = appendRow + probes.length;
  const endRowRange = `A${appendRow}:J${endRow}`;

  return {
    ok: true,
    saved: true,
    writer: 'exceljs',
    todayPrefix,
    maxOrderSeqToday,
    nextSeq,
    computedOrderNo: orderNo,
    orderNo,
    auftragsnummer: orderNo,
    sampleNos,
    ersteProbennr: firstSampleNo,
    letzteProbennr: lastSampleNo,
    endRowRange,
    appendRow,
    startLabNo,
    yearSheetName,
  };
}

module.exports = {
  appendOrderBlockToYearSheet,
  resolveYearSheetName,
  buildHeaderRowValues,
};
