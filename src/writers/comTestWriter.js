const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');

function resolveAbsoluteExcelPath(excelPath, rootDir) {
  return path.isAbsolute(excelPath) ? excelPath : path.join(rootDir, excelPath);
}

async function writeComTestCell(params) {
  if (process.platform !== 'win32') {
    throw new Error('COM test ist nur auf Windows verfuegbar');
  }

  const { rootDir, excelPath, cellPath } = params;
  const absoluteExcelPath = resolveAbsoluteExcelPath(excelPath, rootDir);
  const payloadPath = path.join(os.tmpdir(), `annahme-com-test-${Date.now()}-${Math.random().toString(16).slice(2)}.json`);
  const scriptPath = path.join(rootDir, 'scripts', 'com_test.ps1');

  const payload = {
    excelPath: absoluteExcelPath,
    cellPath,
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
    const lastLine = output.split(/\r?\n/).filter(Boolean).pop() || '{}';

    let parsed;
    try {
      parsed = JSON.parse(lastLine);
    } catch (error) {
      throw new Error(`PowerShell Antwort ungueltig: ${lastLine || stderr || error.message}`);
    }

    if (!parsed.ok || result.status !== 0) {
      throw new Error(parsed.error || 'Unbekannter COM-Testfehler');
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
  writeComTestCell,
};
