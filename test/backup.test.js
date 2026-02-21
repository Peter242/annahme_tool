const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { ensureBackupBeforeCommit, createManualBackup } = require('../src/backup');

function createTempRoot() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-backup-test-'));
}

function baseConfig(overrides = {}) {
  return {
    backupEnabled: true,
    backupPolicy: 'daily',
    backupIntervalMinutes: 60,
    backupRetentionDays: 14,
    backupZip: false,
    backupDir: './backups',
    ...overrides,
  };
}

test('createManualBackup creates backup even when backupEnabled is false', () => {
  const rootDir = createTempRoot();
  const excelRelPath = './data/lab.xlsx';
  const excelAbsPath = path.join(rootDir, 'data', 'lab.xlsx');
  fs.mkdirSync(path.dirname(excelAbsPath), { recursive: true });
  fs.writeFileSync(excelAbsPath, 'dummy');

  const result = createManualBackup({
    config: baseConfig({ backupEnabled: false, backupDir: './manual-backups' }),
    excelPath: excelRelPath,
    rootDir,
  }, { force: true });

  assert.equal(result.created, true);
  assert.equal(result.reason, 'manual');
  assert.ok(typeof result.fileName === 'string' && result.fileName.length > 0);
  assert.ok(typeof result.absoluteBackupPath === 'string' && result.absoluteBackupPath.length > 0);
  assert.equal(fs.existsSync(result.absoluteBackupPath), true);
});

test('ensureBackupBeforeCommit still rotates by policy', () => {
  const rootDir = createTempRoot();
  const excelRelPath = './data/lab.xlsx';
  const excelAbsPath = path.join(rootDir, 'data', 'lab.xlsx');
  fs.mkdirSync(path.dirname(excelAbsPath), { recursive: true });
  fs.writeFileSync(excelAbsPath, 'dummy');

  const config = baseConfig();
  const first = ensureBackupBeforeCommit({ config, excelPath: excelRelPath, rootDir });
  const second = ensureBackupBeforeCommit({ config, excelPath: excelRelPath, rootDir });

  assert.equal(first.created, true);
  assert.equal(second.created, false);
  assert.equal(second.reason, 'rotation_skip');
});

test('cleanup deletes only entries inside configured backupDir', () => {
  const rootDir = createTempRoot();
  const excelRelPath = './data/lab.xlsx';
  const excelAbsPath = path.join(rootDir, 'data', 'lab.xlsx');
  fs.mkdirSync(path.dirname(excelAbsPath), { recursive: true });
  fs.writeFileSync(excelAbsPath, 'dummy');

  const configuredDir = path.join(rootDir, 'my-backups');
  const outsideDir = path.join(rootDir, 'other-dir');
  fs.mkdirSync(configuredDir, { recursive: true });
  fs.mkdirSync(outsideDir, { recursive: true });
  const oldBackupName = 'Annahme_backup_20000101_000000.xlsx';
  fs.writeFileSync(path.join(configuredDir, oldBackupName), 'old');
  fs.writeFileSync(path.join(outsideDir, oldBackupName), 'outside');

  const result = createManualBackup({
    config: baseConfig({ backupDir: './my-backups', backupRetentionDays: 0 }),
    excelPath: excelRelPath,
    rootDir,
  }, { force: true });

  assert.equal(result.created, true);
  assert.equal(fs.existsSync(path.join(configuredDir, oldBackupName)), false);
  assert.equal(fs.existsSync(path.join(outsideDir, oldBackupName)), true);
});

test('createManualBackup fails with clean error when excel file is missing', () => {
  const rootDir = createTempRoot();
  const result = createManualBackup({
    config: baseConfig(),
    excelPath: './data/missing.xlsx',
    rootDir,
  }, { force: true });

  assert.equal(result.created, false);
  assert.equal(result.reason, 'excel_missing');
  assert.ok(typeof result.message === 'string' && result.message.length > 0);
});
