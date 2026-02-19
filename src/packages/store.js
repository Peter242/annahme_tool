const fs = require('fs');
const path = require('path');

const PACKAGES_PATH = process.env.PACKAGES_PATH
  ? path.resolve(process.env.PACKAGES_PATH)
  : path.join(__dirname, '..', '..', 'data', 'packages.json');
let packagesCache = null;
let packagesCacheMtimeMs = -1;

function stringifyCellLike(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string') {
    return value;
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  if (Array.isArray(value)) {
    return value.map((part) => stringifyCellLike(part)).join('');
  }
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => String(part?.text || '')).join('');
    }
    if (value.text !== undefined && value.text !== null) {
      return String(value.text);
    }
    if (value.result !== undefined && value.result !== null) {
      return String(value.result);
    }
    if (value.v !== undefined && value.v !== null) {
      return String(value.v);
    }
  }
  return String(value);
}

function normalizePackageEntry(entry) {
  const value = entry && typeof entry === 'object' ? entry : {};
  return {
    ...value,
    id: value.id === undefined || value.id === null ? '' : String(value.id),
    name: value.name === undefined || value.name === null ? '' : String(value.name),
    code: value.code === undefined || value.code === null ? '' : String(value.code),
    text: stringifyCellLike(value.text),
    row: Number.isFinite(Number(value.row)) ? Number(value.row) : null,
    displayName: value.displayName === undefined || value.displayName === null ? '' : String(value.displayName),
    shortText: value.shortText === undefined || value.shortText === null ? '' : String(value.shortText),
  };
}

function normalizePackages(parsed) {
  return (Array.isArray(parsed) ? parsed : []).map((entry) => normalizePackageEntry(entry));
}

function readPackagesFromDisk() {
  if (!fs.existsSync(PACKAGES_PATH)) {
    return [];
  }

  const raw = fs.readFileSync(PACKAGES_PATH, 'utf-8');
  if (!raw.trim()) {
    return [];
  }

  const parsed = JSON.parse(raw);
  return normalizePackages(parsed);
}

function readPackages(options = {}) {
  const { forceReload = false } = options;
  if (forceReload) {
    packagesCache = null;
    packagesCacheMtimeMs = -1;
  }

  let statMtimeMs = -1;
  try {
    const stat = fs.statSync(PACKAGES_PATH);
    statMtimeMs = Number(stat.mtimeMs);
  } catch (_error) {
    statMtimeMs = -1;
  }

  if (packagesCache && statMtimeMs === packagesCacheMtimeMs) {
    return packagesCache;
  }

  const parsed = readPackagesFromDisk();
  packagesCache = parsed;
  packagesCacheMtimeMs = statMtimeMs;
  return parsed;
}

function writePackages(packages) {
  const normalized = normalizePackages(packages);
  const dir = path.dirname(PACKAGES_PATH);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(PACKAGES_PATH, `${JSON.stringify(normalized, null, 2)}\n`, 'utf-8');
  packagesCache = normalized;
  try {
    const stat = fs.statSync(PACKAGES_PATH);
    packagesCacheMtimeMs = Number(stat.mtimeMs);
  } catch (_error) {
    packagesCacheMtimeMs = -1;
  }
}

function invalidatePackagesCache() {
  packagesCache = null;
  packagesCacheMtimeMs = -1;
}

module.exports = {
  readPackages,
  writePackages,
  invalidatePackagesCache,
};
