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

function ensureBackupBeforeCommit(params) {
  const { config, excelPath, rootDir } = params;
  const now = new Date();
  const backupDir = path.join(rootDir, 'backups');
  const deleted = cleanupBackups(backupDir, config.backupRetentionDays, now);

  if (!config.backupEnabled) {
    return {
      created: false,
      reason: 'backup_disabled',
      fileName: null,
      cleanupDeleted: deleted,
    };
  }

  const absoluteExcelPath = path.isAbsolute(excelPath) ? excelPath : path.join(rootDir, excelPath);
  if (!fs.existsSync(absoluteExcelPath)) {
    return {
      created: false,
      reason: 'excel_missing',
      fileName: null,
      cleanupDeleted: deleted,
    };
  }

  fs.mkdirSync(backupDir, { recursive: true });
  const backupEntries = getBackupEntries(backupDir);
  const lastBackupTime = backupEntries.length > 0 ? backupEntries[0].createdAt : null;
  const shouldCreate = shouldCreateBackup(
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
      cleanupDeleted: deleted,
    };
  }

  const extension = config.backupZip ? 'zip' : 'xlsx';
  const fileName = `${BACKUP_PREFIX}${formatTimestampForFile(now)}.${extension}`;
  const destination = path.join(backupDir, fileName);
  fs.copyFileSync(absoluteExcelPath, destination);

  return {
    created: true,
    reason: 'created',
    fileName,
    cleanupDeleted: deleted,
  };
}

module.exports = {
  ensureBackupBeforeCommit,
};
