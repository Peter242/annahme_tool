const fs = require('fs');
const path = require('path');

const BACKUP_PREFIX = 'Annahme_backup_';
const BACKUP_RE = /^Annahme_backup_(\d{8})_(\d{6})\.(xlsx|zip)$/;

function pad2(value) {
  return String(value).padStart(2, '0');
}

function formatDateCompact(date) {
  return `${date.getFullYear()}${pad2(date.getMonth() + 1)}${pad2(date.getDate())}`;
}

function formatTimestampForFile(date) {
  return `${formatDateCompact(date)}_${pad2(date.getHours())}${pad2(date.getMinutes())}${pad2(date.getSeconds())}`;
}

function parseBackupTimeFromName(fileName) {
  const match = BACKUP_RE.exec(fileName);
  if (!match) {
    return null;
  }

  const [, ymd, hms] = match;
  const year = Number(ymd.slice(0, 4));
  const month = Number(ymd.slice(4, 6)) - 1;
  const day = Number(ymd.slice(6, 8));
  const hour = Number(hms.slice(0, 2));
  const minute = Number(hms.slice(2, 4));
  const second = Number(hms.slice(4, 6));
  return new Date(year, month, day, hour, minute, second);
}

function getBackupEntries(backupDir) {
  if (!fs.existsSync(backupDir)) {
    return [];
  }

  return fs
    .readdirSync(backupDir)
    .map((name) => {
      const createdAt = parseBackupTimeFromName(name);
      if (!createdAt) {
        return null;
      }
      return {
        name,
        filePath: path.join(backupDir, name),
        createdAt,
      };
    })
    .filter(Boolean)
    .sort((a, b) => b.createdAt.getTime() - a.createdAt.getTime());
}

function shouldCreateBackup(policy, now, lastBackupTime, intervalMinutes) {
  if (!lastBackupTime) {
    return true;
  }

  if (policy === 'daily') {
    return formatDateCompact(now) !== formatDateCompact(lastBackupTime);
  }

  const diffMs = now.getTime() - lastBackupTime.getTime();
  const intervalMs = intervalMinutes * 60 * 1000;
  return diffMs >= intervalMs;
}

function cleanupBackups(backupDir, retentionDays, now) {
  if (!fs.existsSync(backupDir)) {
    return [];
  }

  const cutoff = now.getTime() - retentionDays * 24 * 60 * 60 * 1000;
  const deleted = [];
  const entries = getBackupEntries(backupDir);

  entries.forEach((entry) => {
    if (entry.createdAt.getTime() < cutoff) {
      fs.unlinkSync(entry.filePath);
      deleted.push(entry.name);
    }
  });

  return deleted;
}

function resolveBackupDir(rootDir, configuredBackupDir) {
  const rawBackupDir = typeof configuredBackupDir === 'string' ? configuredBackupDir.trim() : '';
  const effectiveBackupDir = rawBackupDir || './backups';
  if (path.isAbsolute(effectiveBackupDir)) {
    return path.normalize(effectiveBackupDir);
  }
  return path.resolve(rootDir, effectiveBackupDir);
}

function ensureBackupDirWritable(backupDir) {
  try {
    fs.mkdirSync(backupDir, { recursive: true });
    fs.accessSync(backupDir, fs.constants.W_OK);
    return { ok: true, message: '' };
  } catch (error) {
    return {
      ok: false,
      message: error instanceof Error ? error.message : String(error),
    };
  }
}

function runBackup(params, options = {}) {
  const { config, excelPath, rootDir } = params;
  const now = new Date();
  const isManual = options.manual === true;
  const force = options.force === true;
  const backupDir = resolveBackupDir(rootDir, config.backupDir);
  const absoluteExcelPath = path.isAbsolute(excelPath) ? excelPath : path.join(rootDir, excelPath);

  if (!isManual && !config.backupEnabled) {
    let deleted = [];
    try {
      deleted = cleanupBackups(backupDir, config.backupRetentionDays, now);
    } catch (_error) {
      deleted = [];
    }
    return {
      created: false,
      reason: 'backup_disabled',
      fileName: null,
      absoluteBackupPath: null,
      cleanupDeleted: deleted,
    };
  }

  if (!fs.existsSync(absoluteExcelPath)) {
    return {
      created: false,
      reason: 'excel_missing',
      fileName: null,
      absoluteBackupPath: null,
      cleanupDeleted: [],
      message: `Excel-Datei nicht gefunden: ${absoluteExcelPath}`,
    };
  }

  const writableCheck = ensureBackupDirWritable(backupDir);
  if (!writableCheck.ok) {
    return {
      created: false,
      reason: 'backup_dir_unwritable',
      fileName: null,
      absoluteBackupPath: null,
      cleanupDeleted: [],
      message: writableCheck.message,
    };
  }

  let deleted = [];
  try {
    deleted = cleanupBackups(backupDir, config.backupRetentionDays, now);
  } catch (error) {
    return {
      created: false,
      reason: 'backup_dir_unwritable',
      fileName: null,
      absoluteBackupPath: null,
      cleanupDeleted: [],
      message: error instanceof Error ? error.message : String(error),
    };
  }

  const backupEntries = getBackupEntries(backupDir);
  const lastBackupTime = backupEntries.length > 0 ? backupEntries[0].createdAt : null;
  const shouldCreate = isManual || force || shouldCreateBackup(
    config.backupPolicy,
    now,
    lastBackupTime,
    config.backupIntervalMinutes,
  );

  if (!shouldCreate) {
    return {
      created: false,
      reason: 'rotation_skip',
      fileName: null,
      absoluteBackupPath: null,
      cleanupDeleted: deleted,
    };
  }

  const extension = config.backupZip ? 'zip' : 'xlsx';
  const fileName = `${BACKUP_PREFIX}${formatTimestampForFile(now)}.${extension}`;
  const destination = path.join(backupDir, fileName);
  try {
    fs.copyFileSync(absoluteExcelPath, destination);
  } catch (error) {
    return {
      created: false,
      reason: 'backup_copy_failed',
      fileName: null,
      absoluteBackupPath: null,
      cleanupDeleted: deleted,
      message: error instanceof Error ? error.message : String(error),
    };
  }

  return {
    created: true,
    reason: isManual ? 'manual' : 'created',
    fileName,
    absoluteBackupPath: destination,
    cleanupDeleted: deleted,
  };
}

function ensureBackupBeforeCommit(params) {
  return runBackup(params, { manual: false, force: false });
}

function createManualBackup(params, options = {}) {
  return runBackup(params, { manual: true, force: options.force === true });
}

module.exports = {
  ensureBackupBeforeCommit,
  createManualBackup,
  resolveBackupDir,
  ensureBackupDirWritable,
};
