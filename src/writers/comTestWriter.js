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
  const scriptPath = path.join(rootDir, 'scripts', 'com_test.ps1');

  const payload = {
    excelPath: absoluteExcelPath,
    workbookFullName: absoluteExcelPath,
    cellPath,
    value: params.value,
  };

  const result = spawnSync('powershell.exe', [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    scriptPath,
  ], {
    input: JSON.stringify(payload),
    encoding: 'utf-8',
  });

  if (result.error) {
    throw new Error(`PowerShell konnte nicht gestartet werden: ${result.error.message}`);
  }

  const output = String(result.stdout || '').trim();
  const stderr = String(result.stderr || '').trim();
  console.log(`[com-test:ps:exit] status=${String(result.status)} signal=${String(result.signal || '')}`);
  if (output) {
    console.log(`[com-test:ps:stdout]\n${output}`);
  }
  if (stderr) {
    console.error(`[com-test:ps:stderr]\n${stderr}`);
  }
  const lastLine = output.split(/\r?\n/).filter(Boolean).pop() || '{}';

  let parsed;
  try {
    parsed = JSON.parse(lastLine);
  } catch (error) {
    throw new Error(`PowerShell Antwort ungueltig: ${lastLine || stderr || error.message}`);
  }

  if (!parsed.ok || result.status !== 0) {
    const message = parsed.error || stderr || `Unbekannter COM-Testfehler (exit=${String(result.status)})`;
    throw new Error(message);
  }

  if (typeof parsed.saved !== 'boolean') {
    parsed.saved = false;
  }

  return parsed;
}

module.exports = {
  writeComTestCell,
};
