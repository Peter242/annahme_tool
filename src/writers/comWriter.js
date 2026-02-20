const path = require('path');
const { renderColumnHFromProbe, renderContainersSummary } = require('../containers');
const { getComWorkerClient } = require('./comWorkerClient');
const { buildProbeJ } = require('../probeJ');

const EXCEL_OPEN_REQUIRED_MESSAGE = 'Fehler: Annahme muss ge\u00f6ffnet sein. Bitte \u00f6ffnen und erneut versuchen';
const FORBIDDEN_ATTACH_ONLY_MESSAGE = 'FORBIDDEN: attempted to start Excel or open workbook';

function resolveAbsoluteExcelPath(excelPath, rootDir) {
  return path.isAbsolute(excelPath) ? excelPath : path.join(rootDir, excelPath);
}

function resolveYearSheetName(config, now) {
  const configured = (config.yearSheetName || '').trim();
  return configured || String(now.getFullYear());
}

function validateComRows(rows, label = 'rows') {
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    if (!Array.isArray(row) || row.length !== 10) {
      const rowType = Array.isArray(row) ? 'array' : typeof row;
      const preview = JSON.stringify(row).slice(0, 200);
      throw new Error(
        `COM row validation failed (${label}): rowIndex=${i} expectedCols=10 actualCols=${Array.isArray(row) ? row.length : 1} rowType=${rowType} row=${preview}`,
      );
    }
  }
}

function buildComHeaderRowPreview(order = {}, config = {}) {
  const quickConfig = {
    quickContainerPlastic: Array.isArray(config.quickContainerPlastic) ? config.quickContainerPlastic : undefined,
    quickContainerGlass: Array.isArray(config.quickContainerGlass) ? config.quickContainerGlass : undefined,
  };
  return [
    'ORDER_NO',
    'y',
    'y',
    '',
    String(order.pbTyp || 'PB'),
    '',
    '',
    order.sameContainersForAll ? renderContainersSummary(order.headerContainers, { config: quickConfig }) : '',
    'HEADER_I',
    'HEADER_J',
  ];
}

function buildComSampleRowPreview(sample = {}, labNo = 'LAB_NO', config = {}) {
  const quickConfig = {
    quickContainerPlastic: Array.isArray(config.quickContainerPlastic) ? config.quickContainerPlastic : undefined,
    quickContainerGlass: Array.isArray(config.quickContainerGlass) ? config.quickContainerGlass : undefined,
  };
  const colH = String(sample.materialGebinde || '').trim() || renderColumnHFromProbe(sample, { config: quickConfig });
  return [
    String(labNo),
    '',
    '',
    String(sample.parameterTextPreview || ''),
    '',
    String(sample.probenbezeichnung || ''),
    String(sample.tiefeOderVolumen || sample.tiefeVolumen || sample.volumen || ''),
    colH,
    '',
    buildProbeJ(sample),
  ];
}

async function writeOrderBlockWithCom(params) {
  if (process.platform !== 'win32') {
    throw new Error('COM writer ist nur auf Windows verfuegbar');
  }

  const { config, rootDir, excelPath, order, termin, cacheHint = null, now = new Date() } = params;
  if (config.allowAutoOpenExcel === true) {
    throw new Error(FORBIDDEN_ATTACH_ONLY_MESSAGE);
  }
  const absoluteExcelPath = resolveAbsoluteExcelPath(excelPath, rootDir);
  const t0 = Date.now();
  const normalizedProbes = (Array.isArray(order?.proben) ? order.proben : []).map((probe) => ({
    ...probe,
    probeJ: buildProbeJ(probe),
  }));
  const payload = {
    excelPath: absoluteExcelPath,
    workbookFullName: absoluteExcelPath,
    yearSheetName: resolveYearSheetName(config, now),
    excelWriteAddressBlock: config.excelWriteAddressBlock !== false,
    allowAutoOpenExcel: false,
    now: now.toISOString(),
    termin: termin || null,
    cacheHint,
    order: {
      ...(order || {}),
      proben: normalizedProbes,
    },
  };

  // Hard preflight validation before invoking PowerShell COM writer.
  const preflightRows = [
    buildComHeaderRowPreview(order, config),
    ...(normalizedProbes.map((sample, idx) => (
      buildComSampleRowPreview(sample, `LAB_${idx + 1}`, config)
    ))),
  ];
  validateComRows(preflightRows, 'preflight');
  const payloadBuildMs = Date.now() - t0;

  const client = getComWorkerClient(rootDir);
  const tRequestStart = Date.now();
  const parsed = await client.request(payload, {
    timeoutMs: 20000,
    retryOnFailure: true,
  });
  const workerRoundtripMs = Date.now() - tRequestStart;

  if (!parsed || parsed.ok !== true) {
    if (String(parsed?.error || '').trim() === EXCEL_OPEN_REQUIRED_MESSAGE) {
      throw new Error(EXCEL_OPEN_REQUIRED_MESSAGE);
    }
    const messageParts = [];
    messageParts.push(parsed?.error || 'Unbekannter COM-Fehler');
    if (parsed?.where) messageParts.push(`where=${parsed.where}`);
    if (parsed?.detail) messageParts.push(`detail=${parsed.detail}`);
    if (parsed?.line) messageParts.push(`line=${parsed.line}`);
    if (parsed?.code) messageParts.push(`code=${String(parsed.code).trim()}`);
    throw new Error(messageParts.join(' | '));
  }

  if (typeof parsed.saved !== 'boolean') {
    parsed.saved = false;
  }
  const timingMs = parsed && typeof parsed.timingMs === 'object' ? parsed.timingMs : {};
  parsed.timingMs = {
    buildPayloadMs: payloadBuildMs,
    ...timingMs,
    workerRoundtripMs,
    writerProcessMs: workerRoundtripMs,
  };
  console.log(`[commit:timing] ${JSON.stringify(parsed.timingMs)}`);

  return parsed;
}

module.exports = {
  writeOrderBlockWithCom,
  validateComRows,
  buildComHeaderRowPreview,
  buildComSampleRowPreview,
};
