const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');
const { renderColumnHFromProbe, renderContainersSummary } = require('../containers');

const EXCEL_OPEN_REQUIRED_MESSAGE = 'Fehler: Annahme muss ge\u00f6ffnet sein. Bitte \u00f6ffnen und erneut versuchen';

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
    'PROBE_J',
  ];
}

async function writeOrderBlockWithCom(params) {
  if (process.platform !== 'win32') {
    throw new Error('COM writer ist nur auf Windows verfuegbar');
  }

  const { config, rootDir, excelPath, order, termin, cacheHint = null, now = new Date() } = params;
  const absoluteExcelPath = resolveAbsoluteExcelPath(excelPath, rootDir);
  const payloadPath = path.join(os.tmpdir(), `annahme-writer-${Date.now()}-${Math.random().toString(16).slice(2)}.json`);
  const scriptPath = path.join(rootDir, 'scripts', 'writer.ps1');

  const payload = {
    excelPath: absoluteExcelPath,
    workbookFullName: absoluteExcelPath,
    yearSheetName: resolveYearSheetName(config, now),
    allowAutoOpenExcel: config.allowAutoOpenExcel === true,
    now: now.toISOString(),
    termin: termin || null,
    cacheHint,
    order,
  };

  // Hard preflight validation before invoking PowerShell COM writer.
  const preflightRows = [
    buildComHeaderRowPreview(order, config),
    ...((Array.isArray(order?.proben) ? order.proben : []).map((sample, idx) => (
      buildComSampleRowPreview(sample, `LAB_${idx + 1}`, config)
    ))),
  ];
  validateComRows(preflightRows, 'preflight');

  fs.writeFileSync(payloadPath, JSON.stringify(payload, null, 2), 'utf-8');
  try {
    const result = spawnSync('powershell.exe', [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      scriptPath,
      '-PayloadPath',
      payloadPath,
    ], {
      encoding: 'utf-8',
    });

    if (result.error) {
      throw new Error(`PowerShell konnte nicht gestartet werden: ${result.error.message}`);
    }

    const output = String(result.stdout || '').trim();
    const stderr = String(result.stderr || '').trim();
    console.log(`[commit:ps:exit] status=${String(result.status)} signal=${String(result.signal || '')}`);
    if (output) {
      console.log(`[commit:ps:stdout]\n${output}`);
    }
    if (stderr) {
      console.error(`[commit:ps:stderr]\n${stderr}`);
    }
    const lastLine = output.split(/\r?\n/).filter(Boolean).pop() || '{}';
    let parsed = null;

    try {
      parsed = JSON.parse(lastLine);
    } catch (error) {
      throw new Error(`PowerShell Antwort ungueltig: ${lastLine || stderr || error.message}`);
    }

    if (!parsed.ok || result.status !== 0) {
      if (String(parsed.error || '').trim() === EXCEL_OPEN_REQUIRED_MESSAGE) {
        throw new Error(EXCEL_OPEN_REQUIRED_MESSAGE);
      }
      const messageParts = [];
      messageParts.push(parsed.error || stderr || `Unbekannter COM-Fehler (exit=${String(result.status)})`);
      if (parsed.where) messageParts.push(`where=${parsed.where}`);
      if (parsed.detail) messageParts.push(`detail=${parsed.detail}`);
      if (parsed.line) messageParts.push(`line=${parsed.line}`);
      if (parsed.code) messageParts.push(`code=${String(parsed.code).trim()}`);
      throw new Error(messageParts.join(' | '));
    }

    if (typeof parsed.saved !== 'boolean') {
      parsed.saved = false;
    }

    return parsed;
  } finally {
    try {
      fs.unlinkSync(payloadPath);
    } catch (_error) {
      // ignore cleanup errors
    }
  }
}

module.exports = {
  writeOrderBlockWithCom,
  validateComRows,
  buildComHeaderRowPreview,
  buildComSampleRowPreview,
};
