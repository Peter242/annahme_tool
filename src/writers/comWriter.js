const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');

function resolveAbsoluteExcelPath(excelPath, rootDir) {
  return path.isAbsolute(excelPath) ? excelPath : path.join(rootDir, excelPath);
}

function resolveYearSheetName(config, now) {
  const configured = (config.yearSheetName || '').trim();
  return configured || String(now.getFullYear());
}

async function writeOrderBlockWithCom(params) {
  if (process.platform !== 'win32') {
    throw new Error('COM writer ist nur auf Windows verfuegbar');
  }

  const { config, rootDir, excelPath, order, termin, now = new Date() } = params;
  const absoluteExcelPath = resolveAbsoluteExcelPath(excelPath, rootDir);
  const payloadPath = path.join(os.tmpdir(), `annahme-writer-${Date.now()}-${Math.random().toString(16).slice(2)}.json`);
  const scriptPath = path.join(rootDir, 'scripts', 'writer.ps1');

  const payload = {
    excelPath: absoluteExcelPath,
    workbookFullName: absoluteExcelPath,
    yearSheetName: resolveYearSheetName(config, now),
    now: now.toISOString(),
    termin: termin || null,
    order,
  };

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
      const message = parsed.error || stderr || `Unbekannter COM-Fehler (exit=${String(result.status)})`;
      throw new Error(message);
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
};
