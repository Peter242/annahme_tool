const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawnSync } = require('child_process');
const express = require('express');
const ExcelJS = require('exceljs');
const { z } = require('zod');
const { makeOrderNumber, nextLabNumbers } = require('./src/numbering');
const { ensureBackupBeforeCommit, createManualBackup, resolveBackupDir, ensureBackupDirWritable } = require('./src/backup');
const { importPackagesFromExcel } = require('./src/packages/importFromExcel');
const { readPackages, writePackages, invalidatePackagesCache } = require('./src/packages/store');
const { createImportPackagesHandler } = require('./src/packages/importRoute');
const { getSheetState, buildTodayPrefix } = require('./src/sheetState');
const { writeOrderBlock } = require('./src/orderWriter');
const { writeComTestCell } = require('./src/writers/comTestWriter');
const { getComWorkerClient } = require('./src/writers/comWorkerClient');
const { calculateTermin } = require('./src/termin');
const { buildParameterTextFromSelection } = require('./src/parameterTextBuilder');
const { mapTogglesToSelection } = require('./src/singleParamsMapper');
const {
  QUICK_CONTAINER_DEFAULTS,
  normalizeQuickContainerConfig,
  normalizeContainers,
  normalizeContainerItems,
  renderContainers,
  renderContainersSummary,
  renderColumnHFromProbe,
} = require('./src/containers');

const configSchema = z
  .object({
    port: z.number().int().min(1).max(65535),
    mode: z.enum(['single', 'writer', 'client']),
    writerBackend: z.enum(['exceljs', 'com', 'comExceljs']),
    excelPath: z.string().trim().min(1),
    yearSheetName: z.string(),
    excelWriteAddressBlock: z.boolean(),
    allowAutoOpenExcel: z.boolean(),
    writerHost: z.string().trim(),
    writerToken: z.string(),
    backupEnabled: z.boolean(),
    backupPolicy: z.enum(['daily', 'interval']),
    backupIntervalMinutes: z.number().int().positive(),
    backupRetentionDays: z.number().int().nonnegative(),
    backupZip: z.boolean(),
    backupDir: z.string().trim().min(1),
    uiShowPackagePreview: z.boolean(),
    uiKuerzelPreset: z.array(z.string()),
    uiRequiredFields: z.array(z.string()),
    uiRequireAtLeastOneSample: z.boolean(),
    uiWarnOnly: z.boolean(),
    uiBlockOnMissing: z.boolean(),
    uiDefaultEilig: z.enum(['ja', 'nein']),
    quickContainerPlastic: z.array(z.string()),
    quickContainerGlass: z.array(z.string()),
  })
  .superRefine((cfg, ctx) => {
    if (cfg.mode === 'writer' && !cfg.writerToken.trim()) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ['writerToken'],
        message: 'writerToken ist im Modus writer erforderlich',
      });
    }
  });

const defaultConfig = {
  port: 3000,
  mode: 'single',
  writerBackend: 'com',
  excelPath: './data/lab.xlsx',
  yearSheetName: '',
  excelWriteAddressBlock: true,
  allowAutoOpenExcel: false,
  writerHost: 'http://localhost:3000',
  writerToken: 'dev-writer-token',
  backupEnabled: true,
  backupPolicy: 'daily',
  backupIntervalMinutes: 60,
  backupRetentionDays: 14,
  backupZip: false,
  backupDir: './backups',
  uiShowPackagePreview: true,
  uiKuerzelPreset: ['AD', 'DV', 'LB', 'DH', 'SE', 'JO', 'RS', 'KH'],
  uiRequiredFields: [
    'kunde',
    'projektName',
    'projektnummer',
    'ansprechpartner',
    'email',
    'kuerzel',
    'proben[0].probenbezeichnung',
    'proben[0].packageId',
  ],
  uiRequireAtLeastOneSample: true,
  uiWarnOnly: true,
  uiBlockOnMissing: false,
  uiDefaultEilig: 'ja',
  quickContainerPlastic: [...QUICK_CONTAINER_DEFAULTS.plastic],
  quickContainerGlass: [...QUICK_CONTAINER_DEFAULTS.glass],
};

const configPath = path.join(__dirname, 'config.json');
const singleParameterCatalogPath = path.join(__dirname, 'data', 'single_parameter_catalog.json');
const builderPackagesPath = path.join(__dirname, 'data', 'builder_packages.json');
const reportDbDir = path.join(__dirname, 'data', 'reportdb');
const reportDbEventsPath = path.join(reportDbDir, 'orders_events.jsonl');
const reportDbIndexPath = path.join(reportDbDir, 'orders_index.csv');
const REPORTDB_CSV_FIELDS = [
  'committedAt',
  'commitId',
  'orderNo',
  'projectNo',
  'projectName',
  'clientName',
  'clientEmail',
  'samplingDate',
  'samplers',
  'transport',
  'addressCompany',
  'addressStreet',
  'addressZip',
  'addressCity',
  'labNoFirst',
  'labNoLast',
  'labNoCount',
  'packageSummary',
  'deleted',
];
const REPORTDB_CSV_HEADER_LINE = REPORTDB_CSV_FIELDS.join(';');
let singleParameterCatalogCache = null;
let singleParameterCatalogUpdatedAt = null;
let builderPackagesCache = null;

function loadSingleParameterCatalog() {
  if (!singleParameterCatalogCache) {
    const raw = fs.readFileSync(singleParameterCatalogPath, 'utf-8');
    singleParameterCatalogCache = JSON.parse(raw);
    singleParameterCatalogUpdatedAt = String(singleParameterCatalogCache?.updatedAt || '').trim() || null;
  }
  return singleParameterCatalogCache;
}

function defaultBuilderPackagesStore() {
  return {
    version: 1,
    updatedAt: '',
    packages: [],
  };
}

function normalizeBuilderPackage(entry) {
  if (!entry || typeof entry !== 'object') return null;
  const id = String(entry.id || '').trim();
  const name = String(entry.name || '').trim();
  if (!id || !name) return null;
  const createdAt = String(entry.createdAt || '').trim();
  const updatedAt = String(entry.updatedAt || '').trim();
  const definition = entry.definition && typeof entry.definition === 'object'
    ? entry.definition
    : { toggles: {} };
  return {
    id,
    type: 'builder',
    name,
    createdAt: createdAt || updatedAt || new Date().toISOString(),
    updatedAt: updatedAt || createdAt || new Date().toISOString(),
    definition,
  };
}

function loadBuilderPackagesStore() {
  if (builderPackagesCache) return builderPackagesCache;
  if (!fs.existsSync(builderPackagesPath)) {
    const initial = defaultBuilderPackagesStore();
    fs.writeFileSync(builderPackagesPath, `${JSON.stringify(initial, null, 2)}\n`, 'utf-8');
    builderPackagesCache = initial;
    return builderPackagesCache;
  }
  const raw = fs.readFileSync(builderPackagesPath, 'utf-8');
  const parsed = JSON.parse(raw);
  const source = parsed && typeof parsed === 'object' ? parsed : defaultBuilderPackagesStore();
  builderPackagesCache = {
    version: Number(source.version || 1),
    updatedAt: String(source.updatedAt || '').trim(),
    packages: (Array.isArray(source.packages) ? source.packages : [])
      .map(normalizeBuilderPackage)
      .filter(Boolean),
  };
  return builderPackagesCache;
}

function saveBuilderPackagesStore(store) {
  const source = store && typeof store === 'object' ? store : defaultBuilderPackagesStore();
  const normalized = {
    version: Number(source.version || 1),
    updatedAt: String(source.updatedAt || '').trim() || new Date().toISOString(),
    packages: (Array.isArray(source.packages) ? source.packages : [])
      .map(normalizeBuilderPackage)
      .filter(Boolean),
  };
  fs.writeFileSync(builderPackagesPath, `${JSON.stringify(normalized, null, 2)}\n`, 'utf-8');
  builderPackagesCache = normalized;
  return normalized;
}

function loadConfig() {
  let rawConfig = { ...defaultConfig };

  if (fs.existsSync(configPath)) {
    const fileContent = fs.readFileSync(configPath, 'utf-8');
    rawConfig = { ...defaultConfig, ...JSON.parse(fileContent) };
  }
  const normalizedQuickConfig = normalizeQuickContainerConfig(rawConfig);
  rawConfig.quickContainerPlastic = normalizedQuickConfig.plastic;
  rawConfig.quickContainerGlass = normalizedQuickConfig.glass;

  const parsed = configSchema.safeParse(rawConfig);
  if (!parsed.success) {
    console.error('Ungueltige Konfiguration in config.json/defaults');
    console.error(parsed.error.format());
    process.exit(1);
  }

  return parsed.data;
}

function saveConfig(config) {
  fs.writeFileSync(configPath, `${JSON.stringify(config, null, 2)}\n`, 'utf-8');
}

function toPublicConfig(config) {
  const { writerToken, ...publicConfig } = config;
  return publicConfig;
}

const app = express();
let runtimeConfig = loadConfig();
const port = runtimeConfig.port;
const COMMIT_REQUEST_TTL_MS = 10 * 60 * 1000;
const COMMIT_REQUEST_MAX_ENTRIES = 200;
const commitRequestStore = new Map();
const sheetStateCachePath = path.join(__dirname, 'data', 'sheetStateCache.json');
const customerStorePath = path.join(__dirname, 'data', 'customers.json');
let sheetStateCache = null;
let customerProfilesCache = [];
customerProfilesCache = readCustomerProfilesFromDisk();

function resolveExcelPath(excelPath) {
  return path.isAbsolute(excelPath) ? excelPath : path.join(__dirname, excelPath);
}

function isWindowsDrivePath(value) {
  return typeof value === 'string' && /^[a-zA-Z]:[\\/]/.test(value.trim());
}

function resolveCommitWriterBackend(config) {
  const configuredBackend = String(config.writerBackend || '').trim().toLowerCase();
  if (configuredBackend === 'comexceljs' || configuredBackend === 'exceljs') {
    return 'exceljs';
  }

  if (config.mode === 'single') {
    if (process.platform === 'win32' && isWindowsDrivePath(config.excelPath)) {
      return 'com';
    }
    return 'exceljs';
  }

  if (configuredBackend === 'com') {
    return 'com';
  }

  return 'exceljs';
}

function getConfig() {
  return runtimeConfig;
}

const EXCEL_NOT_OPEN_USER_MESSAGE = 'Fehler: Annahme muss ge\u00f6ffnet sein. Bitte \u00f6ffnen und erneut versuchen';
const WORKBOOK_READONLY_USER_MESSAGE = 'Annahme.xlsx ist schreibgeschützt oder von einem anderen Benutzer gesperrt. Bitte in Excel schreibbar öffnen oder Sperre lösen.';

function extractWriterDebug(errorMessage) {
  const text = String(errorMessage || '');
  const debug = {};
  let cleaned = text;
  const keys = ['where', 'detail', 'line', 'code'];
  keys.forEach((key) => {
    const regex = new RegExp(`(?:^|\\|)\\s*${key}=([^|]+)`, 'i');
    const match = cleaned.match(regex);
    if (!match) return;
    const rawValue = String(match[1] || '').trim();
    if (rawValue) {
      if (key === 'line') {
        const parsedLine = Number.parseInt(rawValue, 10);
        debug.line = Number.isFinite(parsedLine) ? parsedLine : rawValue;
      } else {
        debug[key] = rawValue;
      }
    }
    cleaned = cleaned.replace(regex, '');
  });
  cleaned = cleaned.replace(/\s*\|\s*/g, ' | ').replace(/^\s*\|\s*|\s*\|\s*$/g, '').trim();
  return {
    userMessage: cleaned || text || 'Unbekannter Writer-Fehler',
    debug: Object.keys(debug).length > 0 ? debug : undefined,
  };
}

function isExcelNotOpenMessage(message) {
  const text = String(message || '');
  return text.includes('Annahme.xlsx muss offen sein')
    || text.includes('Annahme muss ge\u00f6ffnet sein')
    || text.includes('Annahme muss geöffnet sein');
}

function normalizeQuickListPayload(values) {
  const seen = new Set();
  const result = [];
  for (const raw of Array.isArray(values) ? values : []) {
    const normalized = String(raw || '').trim().replace(/\s+/g, ' ');
    if (!normalized || seen.has(normalized)) {
      continue;
    }
    seen.add(normalized);
    result.push(normalized);
  }
  return result;
}

function normalizeKuerzelPresetPayload(values) {
  const seen = new Set();
  const result = [];
  for (const raw of Array.isArray(values) ? values : []) {
    const normalized = String(raw || '').trim().replace(/\s+/g, ' ').toUpperCase();
    if (!normalized || seen.has(normalized)) {
      continue;
    }
    seen.add(normalized);
    result.push(normalized);
  }
  return result;
}

function readSheetStateCacheFromDisk() {
  try {
    if (!fs.existsSync(sheetStateCachePath)) {
      return null;
    }
    const raw = fs.readFileSync(sheetStateCachePath, 'utf-8');
    if (!raw.trim()) {
      return null;
    }
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') {
      return null;
    }
    return parsed;
  } catch (_error) {
    return null;
  }
}

function persistSheetStateCache(cache, force = false) {
  // Runtime cache writes are intentionally disabled to avoid nodemon restart loops
  // from generated files in the project tree. Cache remains in-memory for runtime use.
  void cache;
  void force;
  return false;
}

function getExcelFileMeta(absoluteExcelPath) {
  try {
    const stat = fs.statSync(absoluteExcelPath);
    return {
      fileMtimeMs: Number(stat.mtimeMs),
      excelFileSize: Number(stat.size),
      lastWriteTime: stat.mtime.toISOString(),
    };
  } catch (_error) {
    return {
      fileMtimeMs: -1,
      excelFileSize: -1,
      lastWriteTime: '',
    };
  }
}

function resolveYearPrefixFromSheetName(yearSheetName, now = new Date()) {
  const match = String(yearSheetName || '').trim().match(/^(\d{4})$/);
  const year = match ? Number.parseInt(match[1], 10) : now.getFullYear();
  return String(year % 100).padStart(2, '0');
}

function cellValueToString(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
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
  }

  return String(value);
}

function computeColAHash50(sheet) {
  const lines = [];
  for (let row = 1; row <= 50; row += 1) {
    const raw = sheet.getCell(row, 1).value;
    lines.push(cellValueToString(raw).trim());
  }
  return crypto.createHash('sha1').update(lines.join('\n')).digest('hex');
}

async function probeSheetCacheState(context) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(context.absoluteExcelPath);
  const sheet = workbook.getWorksheet(context.yearSheetName);
  if (!sheet) {
    throw new Error(`Jahresblatt ${context.yearSheetName} nicht gefunden`);
  }

  return {
    sheet,
    colAHash50: computeColAHash50(sheet),
  };
}

function isSheetStateCacheValid(cache, context) {
  if (!cache || typeof cache !== 'object') {
    return false;
  }
  if (cache.excelPath !== context.absoluteExcelPath) {
    return false;
  }
  if (cache.yearSheetName !== context.yearSheetName) {
    return false;
  }
  if (Number(cache.fileMtimeMs) !== Number(context.fileMtimeMs)) {
    return false;
  }
  if (Number(cache.excelFileSize) !== Number(context.excelFileSize)) {
    return false;
  }
  if (String(cache.lastWriteTime || '') !== String(context.lastWriteTime || '')) {
    return false;
  }
  if (String(cache.yearPrefix || '') !== String(context.yearPrefix || '')) {
    return false;
  }
  return true;
}

function normalizeOrderSeqByPrefix(orderSeqByPrefix) {
  const out = {};
  if (!orderSeqByPrefix || typeof orderSeqByPrefix !== 'object') {
    return out;
  }
  for (const [prefix, seq] of Object.entries(orderSeqByPrefix)) {
    const parsed = Number.parseInt(String(seq), 10);
    if (Number.isFinite(parsed) && parsed >= 0) {
      out[prefix] = parsed;
    }
  }
  return out;
}

async function buildFullScanSheetStateCache(context, now = new Date()) {
  const probe = await probeSheetCacheState(context);
  const { sheet, colAHash50 } = probe;
  const state = getSheetState(sheet, now);
  const todayPrefix = buildTodayPrefix(now);
  return {
    version: 2,
    excelPath: context.absoluteExcelPath,
    yearSheetName: context.yearSheetName,
    fileMtimeMs: context.fileMtimeMs,
    excelFileSize: context.excelFileSize,
    lastWriteTime: context.lastWriteTime,
    yearPrefix: context.yearPrefix,
    colAHash50,
    // lastUsedRow is stored as the trailing blank separator row.
    lastUsedRow: Number(state.lastUsedRow) + 1,
    lastLabNo: Number(state.maxLabNumber),
    orderSeqByPrefix: {
      [todayPrefix]: Number(state.maxOrderSeqToday),
    },
    updatedAt: new Date().toISOString(),
  };
}

async function ensureSheetStateCache(config, now = new Date()) {
  const absoluteExcelPath = resolveExcelPath(config.excelPath);
  const yearSheetName = getYearSheetName(config);
  const fileMeta = getExcelFileMeta(absoluteExcelPath);
  const context = {
    absoluteExcelPath,
    yearSheetName,
    fileMtimeMs: fileMeta.fileMtimeMs,
    excelFileSize: fileMeta.excelFileSize,
    lastWriteTime: fileMeta.lastWriteTime,
    yearPrefix: resolveYearPrefixFromSheetName(yearSheetName, now),
  };

  if (sheetStateCache && isSheetStateCacheValid(sheetStateCache, context)) {
    return sheetStateCache;
  }

  const fromDisk = readSheetStateCacheFromDisk();
  if (isSheetStateCacheValid(fromDisk, context)) {
    sheetStateCache = {
      ...fromDisk,
      orderSeqByPrefix: normalizeOrderSeqByPrefix(fromDisk.orderSeqByPrefix),
    };
    return sheetStateCache;
  }

  const rebuilt = await buildFullScanSheetStateCache(context, now);
  sheetStateCache = rebuilt;
  persistSheetStateCache(rebuilt, true);
  return sheetStateCache;
}

function pruneCommitRequestStore(now = Date.now()) {
  for (const [requestId, entry] of commitRequestStore.entries()) {
    if (now - entry.ts > COMMIT_REQUEST_TTL_MS) {
      commitRequestStore.delete(requestId);
    }
  }

  while (commitRequestStore.size > COMMIT_REQUEST_MAX_ENTRIES) {
    const firstKey = commitRequestStore.keys().next().value;
    if (!firstKey) {
      break;
    }
    commitRequestStore.delete(firstKey);
  }
}

function readClientRequestId(value) {
  if (typeof value !== 'string') {
    return '';
  }
  return value.trim();
}

function parseLabNoNumeric(labNoRaw) {
  const m = String(labNoRaw || '').match(/^(\d+)/);
  if (!m) return null;
  const digits = m[1];
  if (digits.length === 8) return null;
  const n = Number(digits);
  return Number.isFinite(n) ? n : null;
}

function computeLabNoRange(samples) {
  const numericNos = (Array.isArray(samples) ? samples : [])
    .map((sample) => {
      if (sample && typeof sample === 'object') {
        return parseLabNoNumeric(sample.labNo ?? sample.labNoRaw ?? '');
      }
      return parseLabNoNumeric(sample);
    })
    .filter((value) => value !== null);
  if (numericNos.length === 0) {
    return { first: '', last: '', count: 0 };
  }
  const min = Math.min(...numericNos);
  const max = Math.max(...numericNos);
  return {
    first: String(min),
    last: String(max),
    count: numericNos.length,
  };
}

function csvEscape(value) {
  if (value === null || value === undefined) return '';
  const text = String(value).replace(/\r\n?/g, ' ').replace(/\n/g, ' ');
  const escaped = text.replace(/"/g, '""');
  if (/[;"\n]/.test(text)) {
    return `"${escaped}"`;
  }
  return escaped;
}

function ensureCsvHeader(filePath, headerLine) {
  if (fs.existsSync(filePath)) return;
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, `${headerLine}\r\n`, 'utf-8');
}

function ensureReportDbFiles() {
  fs.mkdirSync(reportDbDir, { recursive: true });
  if (!fs.existsSync(reportDbEventsPath)) {
    fs.writeFileSync(reportDbEventsPath, '', 'utf-8');
  }
  ensureCsvHeader(reportDbIndexPath, REPORTDB_CSV_HEADER_LINE);
}

function parseCsvLine(line) {
  const out = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      const next = line[i + 1];
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (ch === ';' && !inQuotes) {
      out.push(current);
      current = '';
      continue;
    }
    current += ch;
  }
  out.push(current);
  return out;
}

function readReportDbIndexRows() {
  ensureReportDbFiles();
  const raw = fs.readFileSync(reportDbIndexPath, 'utf-8');
  const lines = raw.split(/\r?\n/).filter((line) => line !== '');
  if (lines.length <= 1) return [];
  const header = parseCsvLine(lines[0]);
  const rows = [];
  for (const line of lines.slice(1)) {
    const cols = parseCsvLine(line);
    const row = {};
    header.forEach((name, idx) => {
      row[name] = String(cols[idx] ?? '');
    });
    rows.push(row);
  }
  return rows;
}

function writeReportDbIndexRows(rows) {
  ensureReportDbFiles();
  const lines = [REPORTDB_CSV_HEADER_LINE];
  for (const row of Array.isArray(rows) ? rows : []) {
    const values = REPORTDB_CSV_FIELDS.map((field) => csvEscape(row?.[field] ?? ''));
    lines.push(values.join(';'));
  }
  fs.writeFileSync(reportDbIndexPath, `${lines.join('\r\n')}\r\n`, 'utf-8');
}

function appendReportDbEvent(event) {
  ensureReportDbFiles();
  fs.appendFileSync(reportDbEventsPath, `${JSON.stringify(event)}\n`, 'utf-8');
}

function createReportDbCommitId() {
  if (typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeSamplerCsvValue(samplers) {
  const list = (Array.isArray(samplers) ? samplers : [])
    .map((item) => String(item || '').trim())
    .filter(Boolean);
  if (list.some((item) => item.toLowerCase() === 'kunde')) {
    return 'Kunde';
  }
  return list.join('+');
}

function parseAddressForReportDb(addressBlock) {
  const lines = String(addressBlock || '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);
  const out = {
    company: '',
    street: '',
    zip: '',
    city: '',
  };
  if (lines.length === 0) return out;
  out.company = lines[0];
  let idx = 1;
  if (lines[idx] && /^zh\s+/i.test(lines[idx])) {
    idx += 1;
  }
  out.street = String(lines[idx] || '');
  const zipCityLine = String(lines[idx + 1] || '');
  const m = zipCityLine.match(/^(\S+)\s+(.+)$/);
  if (m) {
    out.zip = String(m[1] || '').trim();
    out.city = String(m[2] || '').trim();
  } else if (zipCityLine) {
    out.city = zipCityLine;
  }
  return out;
}

function buildPackageDisplayMapForReportDb() {
  const map = new Map();
  try {
    const packages = readPackages();
    for (const pkg of Array.isArray(packages) ? packages : []) {
      const id = String(pkg?.id || '').trim();
      if (!id) continue;
      const display = String(pkg?.displayName || pkg?.name || id).trim();
      map.set(id, display || id);
    }
  } catch (_error) {
    // ignore package map errors, fallback to IDs below
  }
  try {
    const store = loadBuilderPackagesStore();
    const builderPackages = Array.isArray(store?.packages) ? store.packages : [];
    for (const pkg of builderPackages) {
      const rawId = String(pkg?.id || '').trim();
      if (!rawId) continue;
      const id = `builder:${rawId}`;
      const display = String(pkg?.name || rawId).trim();
      map.set(id, display ? `[B] ${display}` : id);
    }
  } catch (_error) {
    // ignore builder package map errors, fallback to IDs below
  }
  return map;
}

function buildPackageSummaryForReportDb(probes, packageDisplayMap) {
  const unique = [];
  const seen = new Set();
  for (const probe of Array.isArray(probes) ? probes : []) {
    const packageId = String(probe?.packageId || '').trim();
    const paketKey = String(probe?.paketKey || '').trim();
    const token = packageId || paketKey;
    if (!token) continue;
    const display = String(packageDisplayMap.get(token) || token).trim();
    if (!display || seen.has(display)) continue;
    seen.add(display);
    unique.push(display);
  }
  if (unique.length <= 3) {
    return unique.join('|');
  }
  return `${unique.slice(0, 3).join('|')}|+${unique.length - 3}`;
}

function upsertReportDbIndexRow(nextRow) {
  const rows = readReportDbIndexRows();
  const orderNo = String(nextRow?.orderNo || '').trim();
  if (!orderNo) {
    throw new Error('orderNo fehlt für ReportDB Upsert');
  }
  const index = rows.findIndex((row) => String(row.orderNo || '').trim() === orderNo);
  const normalized = {};
  REPORTDB_CSV_FIELDS.forEach((field) => {
    normalized[field] = String(nextRow?.[field] ?? '');
  });
  if (index >= 0) {
    rows[index] = normalized;
  } else {
    rows.push(normalized);
  }
  writeReportDbIndexRows(rows);
}

function markReportDbOrderDeleted(orderNo, committedAt, commitId) {
  const normalizedOrderNo = String(orderNo || '').trim();
  if (!normalizedOrderNo) {
    throw new Error('orderNo fehlt');
  }
  const rows = readReportDbIndexRows();
  const idx = rows.findIndex((row) => String(row.orderNo || '').trim() === normalizedOrderNo);
  if (idx >= 0) {
    rows[idx] = {
      ...rows[idx],
      committedAt: String(committedAt || ''),
      commitId: String(commitId || ''),
      deleted: '1',
    };
  } else {
    const row = {};
    REPORTDB_CSV_FIELDS.forEach((field) => {
      row[field] = '';
    });
    row.committedAt = String(committedAt || '');
    row.commitId = String(commitId || '');
    row.orderNo = normalizedOrderNo;
    row.deleted = '1';
    rows.push(row);
  }
  writeReportDbIndexRows(rows);
}

function buildReportDbUpsertRow({ order, orderNo, sampleNos, committedAt, commitId }) {
  const sourceOrder = order && typeof order === 'object' ? order : {};
  const probes = Array.isArray(sourceOrder.proben) ? sourceOrder.proben : [];
  const packageDisplayMap = buildPackageDisplayMapForReportDb();
  const packageSummary = buildPackageSummaryForReportDb(probes, packageDisplayMap);
  const address = parseAddressForReportDb(sourceOrder.adresseBlock);
  const sampleRange = computeLabNoRange(probes.map((probe, idx) => ({
    labNo: sampleNos[idx] ?? probe?.labNo ?? '',
    labNoRaw: probe?.labNoRaw ?? '',
  })));
  return {
    committedAt: String(committedAt || ''),
    commitId: String(commitId || ''),
    orderNo: String(orderNo || ''),
    projectNo: String(sourceOrder.projektnummer || ''),
    projectName: String(sourceOrder.projektName || ''),
    clientName: String(sourceOrder.kunde || ''),
    clientEmail: String(sourceOrder.email || ''),
    samplingDate: String(sourceOrder.probenahmedatum || ''),
    samplers: normalizeSamplerCsvValue(sourceOrder.samplers),
    transport: String(sourceOrder.probentransport || ''),
    addressCompany: address.company,
    addressStreet: address.street,
    addressZip: address.zip,
    addressCity: address.city,
    labNoFirst: sampleRange.first,
    labNoLast: sampleRange.last,
    labNoCount: String(sampleRange.count),
    packageSummary,
    deleted: '0',
  };
}

function saveReportDbUpsertForCommit({ order, orderNo, sampleNos, committedAt }) {
  const normalizedOrderNo = String(orderNo || '').trim();
  if (!normalizedOrderNo) {
    throw new Error('orderNo fehlt für ReportDB Commit');
  }
  const commitId = createReportDbCommitId();
  const row = buildReportDbUpsertRow({
    order,
    orderNo: normalizedOrderNo,
    sampleNos: Array.isArray(sampleNos) ? sampleNos : [],
    committedAt,
    commitId,
  });
  appendReportDbEvent({
    schemaVersion: 1,
    eventType: 'upsert',
    commitId,
    committedAt,
    orderNo: normalizedOrderNo,
    data: row,
  });
  upsertReportDbIndexRow(row);
  return { commitId, row };
}

function normalizeCustomerName(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/,+$/g, '')
    .trim();
}

function normalizeCustomerKey(value) {
  return normalizeCustomerName(value).toLocaleLowerCase('de-DE').replace(/\s+/g, ' ');
}

function splitLines(text) {
  return String(text || '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line !== '');
}

function normalizeAdresseBlock(value) {
  return splitLines(value).join('\n');
}

function sanitizeCustomerProfile(input) {
  const kunde = normalizeCustomerName(input?.kunde);
  const key = normalizeCustomerKey(input?.key || kunde);
  return {
    key,
    kunde,
    ansprechpartner: String(input?.ansprechpartner || '').trim(),
    email: String(input?.email || '').trim(),
    adresseBlock: normalizeAdresseBlock(input?.adresseBlock),
    kopfBemerkung: String(input?.kopfBemerkung || '').trim(),
    usageCount: Number.parseInt(String(input?.usageCount || 0), 10) || 0,
    lastUsed: String(input?.lastUsed || '').trim(),
    updatedAt: String(input?.updatedAt || '').trim(),
  };
}

function mergeCustomerProfilesByKey(profiles) {
  const byKey = new Map();
  for (const rawEntry of Array.isArray(profiles) ? profiles : []) {
    const entry = sanitizeCustomerProfile(rawEntry);
    if (!entry.key || !entry.kunde) continue;
    const existing = byKey.get(entry.key);
    if (!existing) {
      byKey.set(entry.key, entry);
      continue;
    }
    byKey.set(entry.key, {
      ...existing,
      kunde: existing.kunde || entry.kunde,
      ansprechpartner: existing.ansprechpartner || entry.ansprechpartner,
      email: existing.email || entry.email,
      adresseBlock: existing.adresseBlock || entry.adresseBlock,
      kopfBemerkung: existing.kopfBemerkung || entry.kopfBemerkung,
      usageCount: Math.max(Number(existing.usageCount || 0), Number(entry.usageCount || 0)),
      lastUsed: String(existing.lastUsed || '').trim() || String(entry.lastUsed || '').trim(),
      updatedAt: String(existing.updatedAt || '').trim() || String(entry.updatedAt || '').trim(),
    });
  }
  return Array.from(byKey.values());
}

function readCustomerProfilesFromDisk() {
  try {
    if (!fs.existsSync(customerStorePath)) {
      return [];
    }
    const raw = fs.readFileSync(customerStorePath, 'utf-8');
    if (!raw.trim()) {
      return [];
    }
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') {
      return [];
    }
    if (Array.isArray(parsed)) {
      return mergeCustomerProfilesByKey(parsed
        .filter((entry) => entry && typeof entry === 'object')
        .map((entry) => sanitizeCustomerProfile(entry))
        .filter((entry) => entry.kunde !== ''));
    }
    return mergeCustomerProfilesByKey(Object.entries(parsed)
      .map(([key, entry]) => sanitizeCustomerProfile({
        key,
        ...(entry && typeof entry === 'object' ? entry : {}),
      }))
      .filter((entry) => entry.kunde !== ''));
  } catch (_error) {
    return [];
  }
}

function writeCustomerProfilesToDisk(profiles) {
  const dir = path.dirname(customerStorePath);
  fs.mkdirSync(dir, { recursive: true });
  const payload = {};
  for (const rawEntry of Array.isArray(profiles) ? profiles : []) {
    const entry = sanitizeCustomerProfile(rawEntry);
    if (!entry.key || !entry.kunde) continue;
    payload[entry.key] = {
      key: entry.key,
      kunde: entry.kunde,
      ansprechpartner: entry.ansprechpartner,
      email: entry.email,
      adresseBlock: entry.adresseBlock,
      kopfBemerkung: entry.kopfBemerkung,
      usageCount: entry.usageCount,
      lastUsed: entry.lastUsed,
      updatedAt: entry.updatedAt,
    };
  }
  fs.writeFileSync(customerStorePath, `${JSON.stringify(payload, null, 2)}\n`, 'utf-8');
}

function listCustomerProfilesAlpha() {
  return [...customerProfilesCache].sort((a, b) => {
    const usageDiff = Number(b.usageCount || 0) - Number(a.usageCount || 0);
    if (usageDiff !== 0) {
      return usageDiff;
    }
    return String(a.kunde || '').localeCompare(String(b.kunde || ''), 'de');
  });
}

function deleteCustomerProfileById(id) {
  const key = normalizeCustomerKey(id);
  if (!key) {
    return false;
  }
  const before = customerProfilesCache.length;
  customerProfilesCache = customerProfilesCache.filter((entry) => entry.key !== key);
  if (customerProfilesCache.length === before) {
    return false;
  }
  writeCustomerProfilesToDisk(customerProfilesCache);
  return true;
}

function upsertCustomerProfile(fields, options = {}) {
  const { now = new Date(), incrementUsage = false, persist = true } = options;
  const kunde = normalizeCustomerName(fields?.kunde);
  if (!kunde) {
    return null;
  }
  const key = normalizeCustomerKey(kunde);
  const index = customerProfilesCache.findIndex((entry) => entry.key === key);
  const current = index >= 0
    ? sanitizeCustomerProfile(customerProfilesCache[index])
    : sanitizeCustomerProfile({ key, kunde });

  current.key = key;
  current.kunde = kunde;
  current.updatedAt = now.toISOString();
  if (incrementUsage) {
    current.usageCount = Number.parseInt(String(current.usageCount || 0), 10) + 1;
    current.lastUsed = current.updatedAt;
  }

  const assignIfNonEmpty = (fieldName, value, normalizer = null) => {
    const normalized = normalizer ? normalizer(value) : String(value || '').trim();
    if (normalized) {
      current[fieldName] = normalized;
    }
  };
  assignIfNonEmpty('ansprechpartner', fields?.ansprechpartner);
  assignIfNonEmpty('email', fields?.email);
  assignIfNonEmpty('adresseBlock', fields?.adresseBlock, normalizeAdresseBlock);
  assignIfNonEmpty('kopfBemerkung', fields?.kopfBemerkung);

  if (index >= 0) {
    customerProfilesCache[index] = current;
  } else {
    customerProfilesCache.push(current);
  }

  if (persist) {
    writeCustomerProfilesToDisk(customerProfilesCache);
  }
  return current;
}

function parseCustomerFromHeaderCellI(value) {
  const lines = splitLines(value);
  if (lines.length === 0) {
    return null;
  }
  const out = {
    kunde: lines[0],
    ansprechpartner: '',
  };
  if (lines.length > 1 && !/^projekt\b/i.test(lines[1])) {
    out.ansprechpartner = lines[1];
  }
  return out.kunde ? out : null;
}

function parseAdresseBlockFromHeaderCellJ(value) {
  const lines = splitLines(value);
  if (lines.length <= 1) {
    return '';
  }
  const rest = lines.slice(1);
  if (rest.length < 2) {
    return '';
  }
  const lastLine = rest[rest.length - 1];
  if (!/[A-ZÄÖÜ]{2,}/.test(lastLine)) {
    return '';
  }
  return rest.join('\n');
}

async function refreshCustomersFromExcel(config, now = new Date()) {
  const absoluteExcelPath = resolveExcelPath(config.excelPath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(absoluteExcelPath);
  const yearSheetName = getYearSheetName(config);
  const sheet = workbook.getWorksheet(yearSheetName);
  if (!sheet) {
    throw new Error(`Jahresblatt ${yearSheetName} nicht gefunden`);
  }

  let scannedHeaders = 0;
  let upserts = 0;
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const colA = cellValueToString(row.getCell(1).value).trim();
    const colB = cellValueToString(row.getCell(2).value).trim().toLowerCase();
    const colC = cellValueToString(row.getCell(3).value).trim().toLowerCase();
    const looksHeader = (colB === 'y' && colC === 'y') || /^\d{6}8\d{2}$/.test(colA);
    if (!looksHeader) continue;
    scannedHeaders += 1;
    const colI = cellValueToString(row.getCell(9).value).trim();
    const colJ = cellValueToString(row.getCell(10).value).trim();
    const parsed = parseCustomerFromHeaderCellI(colI);
    if (!parsed || !parsed.kunde) continue;
    const updated = upsertCustomerProfile({
      kunde: parsed.kunde,
      ansprechpartner: parsed.ansprechpartner,
      adresseBlock: parseAdresseBlockFromHeaderCellJ(colJ),
    }, { now, incrementUsage: false, persist: false });
    if (updated) {
      upserts += 1;
    }
  }
  writeCustomerProfilesToDisk(customerProfilesCache);
  return {
    sheetName: yearSheetName,
    scannedHeaders,
    upserts,
    customerCount: customerProfilesCache.length,
  };
}

function upsertCustomerProfileFromOrder(order, now = new Date()) {
  return upsertCustomerProfile({
    kunde: order.kunde,
    ansprechpartner: order.ansprechpartner,
    email: order.email,
    adresseBlock: order.adresseBlock,
    kopfBemerkung: order.kopfBemerkung || order.auftragsnotiz,
  }, { now, incrementUsage: true, persist: true });
}

function getYearSheetName(config) {
  const configured = (config.yearSheetName || '').trim();
  return configured || String(new Date().getFullYear());
}

function normalizePackageLookupToken(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

async function applyPaketKeyTextsToOrder(order, absoluteExcelPath) {
  const probes = Array.isArray(order.proben) ? order.proben : [];
  const probesWithPaketKey = probes.filter((probe) => String(probe?.paketKey || '').trim() !== '');
  if (probesWithPaketKey.length === 0) {
    return { order, warnings: [] };
  }

  const warnings = [];
  let packages = [];
  try {
    packages = await importPackagesFromExcel(absoluteExcelPath, 'Vorlagen');
  } catch (error) {
    for (const probe of probesWithPaketKey) {
      const paketKey = String(probe.paketKey || '').trim();
      probe.parameterTextPreview = `UNBEKANNTES PAKET: ${paketKey}`;
      warnings.push(`Paket ${paketKey} konnte nicht aus Vorlagen geladen werden (${error.message})`);
    }
    return { order, warnings };
  }

  const lookup = new Map();
  for (const pkg of packages) {
    const candidates = [pkg.code, pkg.name, pkg.id, `${pkg.name}/${pkg.code}`, `${pkg.code}/${pkg.name}`];
    for (const candidate of candidates) {
      const key = normalizePackageLookupToken(candidate);
      if (key && !lookup.has(key)) {
        lookup.set(key, pkg);
      }
    }
  }

  for (const probe of probesWithPaketKey) {
    const paketKey = String(probe.paketKey || '').trim();
    const found = lookup.get(normalizePackageLookupToken(paketKey));
    if (found && String(found.text || '').trim() !== '') {
      probe.parameterTextPreview = String(found.text);
    } else {
      probe.parameterTextPreview = `UNBEKANNTES PAKET: ${paketKey}`;
      warnings.push(`Paket nicht gefunden: ${paketKey}`);
    }
  }

  return { order, warnings };
}

function buildOrderCommitExample() {
  return {
    kunde: 'Musterkunde GmbH',
    projektName: 'Projekt Muster Name',
    projektnummer: 'P-2026-001',
    ansprechpartner: 'Max Mustermann',
    email: 'max@example.com',
    kopfBemerkung: 'Allgemeiner Hinweis zur Kopfzeile',
    adresseBlock: 'Musterkunde GmbH\nzH Max Mustermann\nMusterstrasse 1\n12345 MUSTERSTADT',
    kuerzel: 'MM',
    eilig: false,
    probenahmedatum: '2026-02-13',
    samplers: ['SB', 'JO'],
    sameContainersForAll: false,
    headerContainers: {
      mode: 'perOrder',
      items: [],
    },
    probenEingangDatum: '2026-02-14',
    proben: [
      {
        probenbezeichnung: 'Probe 1',
        material: 'Boden',
        gewicht: 1.2,
        containers: {
          mode: 'perSample',
          items: ['K:30mL+HCl', 'K:30mL+HCl', 'G:1L'],
        },
        paketKey: 'DepV/DepV DK0',
      },
    ],
  };
}

function buildOrderSchemaInfo() {
  return {
    fields: {
      kunde: 'string (required)',
      projektName: 'string (required, legacy fallback: projekt/projektname)',
      projektnummer: 'string (required)',
      ansprechpartner: 'string (optional)',
      email: 'string (optional)',
      kopfBemerkung: 'string (optional)',
      adresseBlock: 'string (optional, mehrzeilig fuer Kopfzelle J)',
      kuerzel: 'string (optional)',
      eilig: 'boolean (optional)',
      probenahmedatum: 'string YYYY-MM-DD (optional, Ausgabe als Probenahme in Kopf-Spalte I)',
      samplers: 'string[] (optional, Werte: SB|LW|AR|JO|AD|Kunde; Kunde ist exklusiv)',
      sameContainersForAll: 'boolean (optional, wenn true nutzt Kopf-Gebinde fuer alle Proben)',
      headerContainers: 'object (optional, siehe containers schema)',
      probenEingangDatum: 'string YYYY-MM-DD (optional)',
      probeNochNichtDa: 'boolean (optional)',
      sampleNotArrived: 'boolean (optional)',
      proben: 'array of sample objects (optional)',
    },
    sampleFields: {
      probenbezeichnung: 'string (optional)',
      material: 'string (optional, frei editierbar)',
      matrixTyp: 'string (optional, legacy fallback)',
      gewicht: 'number > 0 (optional)',
      gewichtEinheit: 'string (optional)',
      volumen: 'number > 0 (optional, legacy)',
      tiefeVolumen: 'string|number (optional, legacy)',
      tiefeOderVolumen: 'string (optional)',
      geruch: 'string (optional)',
      packageId: 'string (optional)',
      paketKey: 'string (optional, lookup in sheet Vorlagen)',
      parameterTextPreview: 'string (optional)',
      geruchAuffaelligkeit: 'string (optional)',
      bemerkung: 'string (optional)',
      materialGebinde: 'string (optional)',
      material: 'string (optional)',
      gebinde: 'string (optional)',
      gebindeItems: 'string[] (optional)',
      gebindeKonservierung: 'string[] (optional)',
      gebindeSonstiges: 'string (optional)',
      gebindeSummary: 'string (optional)',
      containers: 'object (optional) { mode: perSample|perOrder, items: token[] }',
    },
    quickContainerDefaults: QUICK_CONTAINER_DEFAULTS,
  };
}

const level1Fields = [
  'excelPath',
  'yearSheetName',
  'excelWriteAddressBlock',
  'allowAutoOpenExcel',
  'backupEnabled',
  'backupPolicy',
  'backupIntervalMinutes',
  'backupRetentionDays',
  'backupZip',
  'backupDir',
  'uiShowPackagePreview',
  'uiKuerzelPreset',
  'uiRequiredFields',
  'uiRequireAtLeastOneSample',
  'uiWarnOnly',
  'uiBlockOnMissing',
  'uiDefaultEilig',
  'quickContainerPlastic',
  'quickContainerGlass',
];
const level2Fields = ['mode', 'writerHost', 'writerToken', 'writerBackend'];
const allEditableFields = [...level1Fields, ...level2Fields];

const configUpdateSchema = z.object({
  excelPath: z.string().trim().min(1).optional(),
  yearSheetName: z.string().optional(),
  excelWriteAddressBlock: z.boolean().optional(),
  allowAutoOpenExcel: z.boolean().optional(),
  backupEnabled: z.boolean().optional(),
  backupPolicy: z.enum(['daily', 'interval']).optional(),
  backupIntervalMinutes: z.number().int().positive().optional(),
  backupRetentionDays: z.number().int().nonnegative().optional(),
  backupZip: z.boolean().optional(),
  backupDir: z.string().trim().min(1).optional(),
  uiShowPackagePreview: z.boolean().optional(),
  uiKuerzelPreset: z.array(z.string()).optional(),
  uiRequiredFields: z.array(z.string()).optional(),
  uiRequireAtLeastOneSample: z.boolean().optional(),
  uiWarnOnly: z.boolean().optional(),
  uiBlockOnMissing: z.boolean().optional(),
  uiDefaultEilig: z.enum(['ja', 'nein']).optional(),
  quickContainerPlastic: z.array(z.string()).optional(),
  quickContainerGlass: z.array(z.string()).optional(),
  mode: z.enum(['single', 'writer', 'client']).optional(),
  writerBackend: z.enum(['exceljs', 'com', 'comExceljs']).optional(),
  writerHost: z.string().trim().optional(),
  writerToken: z.string().optional(),
}).strict();

const containersSchema = z.object({
  mode: z.enum(['perSample', 'perOrder']).optional(),
  items: z.array(z.string().trim().min(1)).optional().default([]),
  history: z.array(z.string().trim().min(1)).optional().default([]),
}).strict();

const sampleSchema = z
  .object({
    probenbezeichnung: z.string().trim().optional(),
    matrixTyp: z.string().trim().optional(),
    material: z.string().optional(),
    gewicht: z.number().positive().optional(),
    gewichtEinheit: z.string().trim().optional(),
    volumen: z.number().positive().optional(),
    packageId: z.string().trim().optional(),
    paketKey: z.string().trim().optional(),
    parameterTextPreview: z.string().optional(),
    singleParams: z.any().optional(),
    packageAddons: z.array(z.string().trim().min(1)).optional(),
    tiefeVolumen: z.union([z.string(), z.number()]).optional(),
    tiefeOderVolumen: z.string().trim().optional(),
    geruch: z.string().trim().optional(),
    geruchAuffaelligkeit: z.string().trim().optional(),
    geruchOption: z.string().trim().optional(),
    geruchSonstiges: z.string().trim().optional(),
    bemerkung: z.string().trim().optional(),
    materialGebinde: z.string().optional(),
    gebinde: z.string().optional(),
    gebindeItems: z.array(z.string().trim()).optional(),
    gebindeKonservierung: z.array(z.string().trim()).optional(),
    gebindeSonstiges: z.string().trim().optional(),
    gebindeSummary: z.string().trim().optional(),
    containers: containersSchema.optional(),
  });

const orderSchema = z
  .object({
    kunde: z.string().trim().optional(),
    projektName: z.string().trim().optional(),
    projektnummer: z.string().trim().optional(),
    auftragsnotiz: z.string().optional(),
    kopfBemerkung: z.string().optional(),
    pbTyp: z.enum(['PB', 'AI', 'AKN']).optional(),
    auftraggeberKurz: z.string().optional(),
    ansprechpartner: z.string().optional(),
    email: z.string().optional(),
    adresseBlock: z.string().optional(),
    probenahmedatum: z.string().optional(),
    samplers: z.array(z.enum(['SB', 'LW', 'AR', 'JO', 'AD', 'Kunde'])).optional(),
    erfasstKuerzel: z.string().optional(),
    kuerzel: z.string().optional(),
    terminDatum: z.string().optional(),
    eilig: z.boolean().optional().default(false),
    probentransport: z.enum(['CUA', 'AG']).optional(),
    sameContainersForAll: z.boolean().optional().default(false),
    headerContainers: containersSchema.optional(),
    probeNochNichtDa: z.boolean().optional().default(false),
    sampleNotArrived: z.boolean().optional().default(false),
    probenEingangDatum: z
      .string()
      .date('ProbenEingangDatum muss ein gueltiges Datum sein (YYYY-MM-DD)')
      .optional(),
    proben: z.array(sampleSchema).optional().default([]),
  });

function isCommitAllowed() {
  return getConfig().mode !== 'client';
}

function parsePreviewDate(value) {
  if (value instanceof Date) {
    return Number.isFinite(value.getTime()) ? value : null;
  }
  if (typeof value !== 'string') {
    return null;
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

function formatPreviewGermanDate(value) {
  const parsed = parsePreviewDate(value);
  if (!parsed) return '';
  const dd = String(parsed.getDate()).padStart(2, '0');
  const mm = String(parsed.getMonth() + 1).padStart(2, '0');
  const yyyy = parsed.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

function normalizeSamplerSelection(value) {
  const allowed = ['SB', 'LW', 'AR', 'JO', 'AD', 'Kunde'];
  const source = Array.isArray(value) ? value : [];
  const seen = new Set();
  const out = [];
  source.forEach((rawValue) => {
    const raw = String(rawValue || '').trim();
    if (!raw) return;
    const normalized = raw.toLowerCase() === 'kunde' ? 'Kunde' : raw.toUpperCase();
    if (!allowed.includes(normalized) || seen.has(normalized)) return;
    seen.add(normalized);
    out.push(normalized);
  });
  if (out.includes('Kunde')) return ['Kunde'];
  return allowed.filter((item) => item !== 'Kunde' && out.includes(item));
}

function formatProbenahmeLineForPreview(rawDate, samplers) {
  const probenahme = formatPreviewGermanDate(rawDate || '');
  if (!probenahme) return '';
  const normalizedSamplers = normalizeSamplerSelection(samplers);
  if (normalizedSamplers.length < 1) return `Probenahme: ${probenahme}`;
  return `Probenahme: ${probenahme} ${normalizedSamplers.join(' + ')}`;
}

function samePreviewDay(left, right) {
  const a = parsePreviewDate(left);
  const b = parsePreviewDate(right);
  if (!a || !b) return false;
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
}

function buildHeaderILinesPreview(order, now = new Date()) {
  const kunde = String(order.auftraggeberKurz || order.kunde || '').trim();
  const projekt = String(order.projektName || order.projektname || order.projekt || '').trim();
  const projektNr = String(order.projektnummer || '').trim();
  const ansprechpartner = String(order.ansprechpartner || '').trim();
  const probenahmeLine = formatProbenahmeLineForPreview(order.probenahmedatum, order.samplers);
  const probenEingang = formatPreviewGermanDate(order.probenEingangDatum || '');
  const transportRaw = String(order.probentransport || '').trim().toUpperCase();
  const lines = [];
  if (kunde) lines.push(kunde);
  if (ansprechpartner) lines.push(ansprechpartner);
  if (projektNr) lines.push(`Projekt Nr: ${projektNr}`);
  if (projekt) lines.push(`Projekt: ${projekt}`);
  if (probenahmeLine) lines.push(probenahmeLine);
  if (probenEingang && !samePreviewDay(order.probenEingangDatum, now)) {
    lines.push(`Eingangsdatum: ${probenEingang}`);
  }
  if (transportRaw === 'CUA' || transportRaw === 'AG') {
    lines.push(`Transport: ${transportRaw}`);
  }
  return lines;
}

function normalizeSingleParamsTogglesForMerge(value) {
  const source = value && typeof value === 'object' ? value : {};
  const out = {};
  Object.entries(source).forEach(([rawKey, toggle]) => {
    const key = String(rawKey || '').trim();
    if (!key || !toggle || typeof toggle !== 'object' || toggle.selected !== true) return;
    const media = {};
    const sourceMedia = toggle.media && typeof toggle.media === 'object' ? toggle.media : {};
    Object.entries(sourceMedia).forEach(([medium, enabled]) => {
      const mediumKey = String(medium || '').trim();
      if (mediumKey && enabled === true) media[mediumKey] = true;
    });
    out[key] = {
      selected: true,
      lab: String(toggle.lab || '').trim(),
      media,
      vorOrt: toggle.vorOrt === true,
    };
  });
  return out;
}

function stripExistingPktLines(text) {
  if (!text) return '';
  const lines = String(text).split(/\r?\n/);
  const out = lines.filter((line) => !String(line).trimStart().toUpperCase().startsWith('PKT:'));
  return out.join('\n').trim();
}

function buildPktAddonSuffixFromToggles(catalog, toggles) {
  if (!toggles || typeof toggles !== 'object') return '';
  const params = Array.isArray(catalog?.parameters) ? catalog.parameters : [];
  const byKey = new Map(params.map((p) => [String(p?.key || '').trim(), p]));
  const labels = [];
  Object.entries(toggles).forEach(([rawKey, toggle]) => {
    const key = String(rawKey || '').trim();
    if (!key || !toggle || toggle.selected !== true) return;
    const param = byKey.get(key);
    const labelLong = String(param?.labelLong || '').trim();
    labels.push(labelLong || key);
  });
  const unique = Array.from(new Set(labels));
  unique.sort((a, b) => a.localeCompare(b, 'de', { sensitivity: 'base' }));
  if (unique.length < 1) return '';
  return `+${unique.join(', +')}`;
}

function extractLegacyManualAddonLines(addonText) {
  const source = String(addonText || '').replace(/\r\n?/g, '\n');
  if (!source) return [];
  return source
    .split('\n')
    .map((line) => String(line || '').trim())
    .filter((line) => line && /^[^:]+:/.test(line))
    .filter((line) => !/^(PKT|PV|VOR\s*ORT)\s*:/i.test(line));
}

function buildLegacyManualBlock(packageDisplayName, addonSuffix, addonText) {
  const packageName = String(packageDisplayName || '').trim();
  const suffix = String(addonSuffix || '').trim();
  const manualLines = extractLegacyManualAddonLines(addonText);
  const pktLine = `PKT: ${[packageName, suffix].filter(Boolean).join(' ')}`.trim();
  return [
    'MANUELL (Legacy Paket): Addons bitte in Excel in die Paketzeilen übernehmen.',
    pktLine,
    ...manualLines,
  ].join('\n');
}

function getProbeSingleParamToggles(probe) {
  const singleParams = probe?.singleParams;
  if (!singleParams || typeof singleParams !== 'object') return {};
  const toggles = singleParams.toggles && typeof singleParams.toggles === 'object'
    ? singleParams.toggles
    : {};
  return normalizeSingleParamsTogglesForMerge(toggles);
}

function isBuilderPackageId(value) {
  return String(value || '').trim().toLowerCase().startsWith('builder:');
}

function normalizeBuilderPackageId(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  return raw.replace(/^builder:/i, '').trim();
}

function mergeBuilderPackageAddonToggles(baseToggles, addonToggles) {
  const base = normalizeSingleParamsTogglesForMerge(baseToggles);
  const addons = normalizeSingleParamsTogglesForMerge(addonToggles);
  const merged = {};
  const keys = new Set([...Object.keys(base), ...Object.keys(addons)]);
  keys.forEach((key) => {
    const baseToggle = base[key];
    const addonToggle = addons[key];
    if (baseToggle && !addonToggle) {
      merged[key] = baseToggle;
      return;
    }
    if (!baseToggle && addonToggle) {
      merged[key] = addonToggle;
      return;
    }
    if (!baseToggle || !addonToggle) return;
    const media = {};
    const baseMedia = baseToggle.media && typeof baseToggle.media === 'object' ? baseToggle.media : {};
    const addonMedia = addonToggle.media && typeof addonToggle.media === 'object' ? addonToggle.media : {};
    Object.entries(baseMedia).forEach(([medium, enabled]) => {
      if (enabled === true) media[medium] = true;
    });
    Object.entries(addonMedia).forEach(([medium, enabled]) => {
      if (enabled === true) media[medium] = true;
    });
    const vorOrt = baseToggle.vorOrt === true || addonToggle.vorOrt === true;
    merged[key] = {
      selected: true,
      lab: String(baseToggle.lab || '').trim() || String(addonToggle.lab || '').trim(),
      media: vorOrt ? {} : media,
      vorOrt,
    };
  });
  return merged;
}

function resolvePackageTextForProbe({ probe, packageById, builderPackageById, catalog }) {
  const packageId = String(probe?.packageId || '').trim();
  if (!packageId) {
    return {
      text: String(probe?.parameterTextPreview || '').trim(),
      legacyMergeStatus: null,
      legacyMergeReason: null,
    };
  }

  if (isBuilderPackageId(packageId)) {
    const builderId = normalizeBuilderPackageId(packageId);
    const builderPackage = builderPackageById.get(builderId) || null;
    if (!builderPackage || !catalog) {
      return {
        text: String(probe?.parameterTextPreview || '').trim(),
        legacyMergeStatus: null,
        legacyMergeReason: null,
      };
    }
    const definition = builderPackage.definition && typeof builderPackage.definition === 'object'
      ? builderPackage.definition
      : {};
    const baseToggles = definition && typeof definition.toggles === 'object' ? definition.toggles : {};
    const addonToggles = getProbeSingleParamToggles(probe);
    const mergedToggles = mergeBuilderPackageAddonToggles(baseToggles, addonToggles);
    const addonSuffix = buildPktAddonSuffixFromToggles(catalog, addonToggles);
    const selection = mapTogglesToSelection({
      catalog,
      toggles: mergedToggles,
    });
    const builtText = buildParameterTextFromSelection(selection);
    const packageName = String(builderPackage.name || builderId || packageId).trim();
    const pktHeader = packageName ? `PKT: ${packageName}${addonSuffix}` : '';
    if (!pktHeader) {
      return { text: builtText, legacyMergeStatus: null, legacyMergeReason: null };
    }
    if (!builtText) {
      return { text: pktHeader, legacyMergeStatus: null, legacyMergeReason: null };
    }
    return { text: `${pktHeader}\n${builtText}`, legacyMergeStatus: null, legacyMergeReason: null };
  }

  const packageTemplate = packageById.get(packageId) || null;
  const packageDisplayName = packageTemplate
    ? String(packageTemplate.displayName || packageTemplate.name || packageTemplate.id || packageId).trim()
    : String(packageId).trim();
  const rawBaseText = packageTemplate ? String(packageTemplate.text || '').trim() : String(probe?.parameterTextPreview || '').trim();
  const baseText = stripExistingPktLines(rawBaseText);
  const addonToggles = getProbeSingleParamToggles(probe);
  const addonSuffix = buildPktAddonSuffixFromToggles(catalog, addonToggles);
  let addonText = '';
  if (probe?.singleParams) {
    if (catalog && probe.singleParams && typeof probe.singleParams === 'object' && probe.singleParams.toggles) {
      const addonSelection = mapTogglesToSelection({
        catalog,
        toggles: probe.singleParams.toggles,
      });
      addonText = buildParameterTextFromSelection(addonSelection);
    } else {
      addonText = buildParameterTextFromSelection(probe.singleParams);
    }
  }
  const hasAddons = Boolean(String(addonSuffix || '').trim());
  const pktHeader = packageDisplayName ? `PKT: ${packageDisplayName}` : '';
  const baseSections = [pktHeader, baseText].filter(Boolean);
  const bodyText = baseSections.join('\n');
  const manualBlock = hasAddons
    ? buildLegacyManualBlock(packageDisplayName, addonSuffix, addonText)
    : '';
  const text = manualBlock
    ? [bodyText, manualBlock].filter(Boolean).join('\n\n')
    : bodyText;
  return {
    text,
    legacyMergeStatus: hasAddons ? 'appended' : null,
    legacyMergeReason: hasAddons ? 'legacy_manual_block' : null,
  };
}

function buildOrderPreview(order, options = {}) {
  const { lastLabNo = 26203, now = new Date() } = options;
  const warnings = [];
  const packages = readPackages();
  const packageById = new Map(packages.map((pkg) => [pkg.id, pkg]));
  let builderPackageById = new Map();
  try {
    const builderStore = loadBuilderPackagesStore();
    const builderPackages = Array.isArray(builderStore?.packages) ? builderStore.packages : [];
    builderPackageById = new Map(builderPackages.map((pkg) => [String(pkg?.id || '').trim(), pkg]));
  } catch (_error) {
    builderPackageById = new Map();
  }
  let singleParamCatalog = null;
  try {
    singleParamCatalog = loadSingleParameterCatalog();
  } catch (_error) {
    singleParamCatalog = null;
  }
  const probes = Array.isArray(order.proben) ? order.proben : [];
  const vorschau = {
    ...order,
    headerILines: buildHeaderILinesPreview(order, now),
    proben: probes.map((probe) => {
      const resolved = resolvePackageTextForProbe({
        probe,
        packageById,
        builderPackageById,
        catalog: singleParamCatalog,
      });
      return {
        ...probe,
        parameterTextPreview: resolved.text,
        renderedTextD: resolved.text,
        legacyMergeStatus: resolved.legacyMergeStatus || undefined,
        legacyMergeReason: resolved.legacyMergeReason || undefined,
      };
    }),
  };

  const sampleNotArrived = order.probeNochNichtDa || order.sampleNotArrived === true;
  const termin = sampleNotArrived ? null : calculateTermin(order.probenEingangDatum, order.eilig);
  if (!sampleNotArrived && !termin) {
    warnings.push('Termin konnte nicht berechnet werden, weil ProbenEingangDatum fehlt oder ungueltig ist');
  }

  const xy = 1;
  const lastLab = Number.isFinite(Number(lastLabNo)) ? Number(lastLabNo) : 26203;
  const orderNumberPreview = order.probenEingangDatum ? makeOrderNumber(order.probenEingangDatum, xy) : null;
  const labNumberPreview = nextLabNumbers(lastLab, probes.length);

  return {
    ok: true,
    vorschau,
    headerILines: vorschau.headerILines,
    headerIText: vorschau.headerILines.join('\n'),
    termin,
    orderNumberPreview,
    labNumberPreview,
    warnings,
  };
}

async function buildCommitPreviewState(order, config, now = new Date()) {
  const commitWarnings = [];
  const absoluteExcelPath = resolveExcelPath(config.excelPath);
  const enrichedOrderResult = await applyPaketKeyTextsToOrder(order, absoluteExcelPath);
  const orderForWrite = enrichedOrderResult.order;
  const packages = readPackages();
  const packageById = new Map(packages.map((pkg) => [pkg.id, pkg]));
  let builderPackageById = new Map();
  try {
    const builderStore = loadBuilderPackagesStore();
    const builderPackages = Array.isArray(builderStore?.packages) ? builderStore.packages : [];
    builderPackageById = new Map(builderPackages.map((pkg) => [String(pkg?.id || '').trim(), pkg]));
  } catch (_error) {
    builderPackageById = new Map();
  }
  let singleParamCatalog = null;
  try {
    singleParamCatalog = loadSingleParameterCatalog();
  } catch (_error) {
    singleParamCatalog = null;
  }
  const normalizedOrderForWrite = {
    ...orderForWrite,
    proben: (Array.isArray(orderForWrite.proben) ? orderForWrite.proben : []).map((probe) => {
      if (!probe || typeof probe !== 'object') {
        return probe;
      }
      const hasPackage = String(probe.packageId || '').trim() || String(probe.paketKey || '').trim();
      if (hasPackage) {
        const resolved = resolvePackageTextForProbe({
          probe,
          packageById,
          builderPackageById,
          catalog: singleParamCatalog,
        });
        return {
          ...probe,
          parameterTextPreview: resolved.text,
          legacyMergeStatus: resolved.legacyMergeStatus || undefined,
          legacyMergeReason: resolved.legacyMergeReason || undefined,
        };
      }
      if (!probe.singleParams) {
        return { ...probe };
      }
      let built = '';
      if (singleParamCatalog && probe.singleParams && typeof probe.singleParams === 'object' && probe.singleParams.toggles) {
        const selection = mapTogglesToSelection({
          catalog: singleParamCatalog,
          toggles: probe.singleParams.toggles,
        });
        built = buildParameterTextFromSelection(selection);
      } else {
        built = buildParameterTextFromSelection(probe.singleParams);
      }
      if (!built) {
        return { ...probe };
      }
      return {
        ...probe,
        parameterTextPreview: built,
      };
    }),
  };
  commitWarnings.push(...enrichedOrderResult.warnings);

  let cache = null;
  try {
    cache = await ensureSheetStateCache(config, now);
  } catch (error) {
    console.warn(`[sheet-cache] fallback without cache due to scan error: ${error.message}`);
    cache = null;
  }

  const preview = buildOrderPreview(normalizedOrderForWrite, {
    now,
    lastLabNo: cache ? Number.parseInt(String(cache.lastLabNo || 0), 10) : undefined,
  });
  commitWarnings.push(...(Array.isArray(preview.warnings) ? preview.warnings : []));
  const todayPrefix = buildTodayPrefix(now);
  const maxOrderSeqToday = cache
    ? Number(cache.orderSeqByPrefix?.[todayPrefix] || 0)
    : 0;
  const nextSeq = maxOrderSeqToday + 1;
  const computedOrderNo = `${todayPrefix}${String(nextSeq).padStart(2, '0')}`;
  const appendRow = cache
    ? Number.parseInt(String(cache.lastUsedRow || 1), 10) + 1
    : null;
  const startLabNo = cache
    ? Math.max(Number.parseInt(String(cache.lastLabNo || 0), 10), 9999) + 1
    : null;
  const cacheHint = cache
    ? {
      appendRow,
      startLabNo,
      todayPrefix,
      maxOrderSeqToday,
      nextSeq,
      computedOrderNo,
    }
    : null;

  return {
    absoluteExcelPath,
    orderForWrite: normalizedOrderForWrite,
    commitWarnings,
    cache,
    preview,
    todayPrefix,
    maxOrderSeqToday,
    nextSeq,
    computedOrderNo,
    appendRow,
    startLabNo,
    cacheHint,
  };
}

function parseOrderOrRespond(req, res) {
  const quickConfig = normalizeQuickContainerConfig(getConfig());
  const incoming = req.body && typeof req.body === 'object' ? { ...req.body } : {};
  const projektNameFromClient = String(incoming.projektName || '').trim();
  const projektnameFromClient = String(incoming.projektname || '').trim();
  const projektLegacy = String(incoming.projekt || '').trim();
  const normalizedProjektName = projektNameFromClient || projektnameFromClient || projektLegacy;
  if (normalizedProjektName) {
    incoming.projektName = normalizedProjektName;
  }
  delete incoming.projektname;
  delete incoming.projekt;

  if (incoming.probenEingangDatum === null || incoming.probenEingangDatum === '') {
    incoming.probenEingangDatum = undefined;
  }
  if (!Array.isArray(incoming.samplers)) {
    incoming.samplers = undefined;
  } else {
    incoming.samplers = normalizeSamplerSelection(incoming.samplers);
  }
  if (!String(incoming.kopfBemerkung || '').trim() && String(incoming.auftragsnotiz || '').trim()) {
    incoming.kopfBemerkung = String(incoming.auftragsnotiz).trim();
  }
  if (!String(incoming.auftragsnotiz || '').trim() && String(incoming.kopfBemerkung || '').trim()) {
    incoming.auftragsnotiz = String(incoming.kopfBemerkung).trim();
  }
  if (incoming.adresseBlock === null || incoming.adresseBlock === undefined) {
    incoming.adresseBlock = undefined;
  } else {
    incoming.adresseBlock = normalizeAdresseBlock(incoming.adresseBlock);
  }
  if (Array.isArray(incoming.proben)) {
    incoming.proben = incoming.proben.map((probe) => {
      if (!probe || typeof probe !== 'object') {
        return probe;
      }
      const nextProbe = { ...probe };
      const material = String(nextProbe.material || '').trim();
      const matrixTyp = String(nextProbe.matrixTyp || '').trim();
      nextProbe.material = material || matrixTyp || '';
      if (nextProbe.containers && typeof nextProbe.containers === 'object') {
        nextProbe.containers = normalizeContainers(nextProbe.containers, { modeDefault: 'perSample' });
      }
      return nextProbe;
    });
  }

  const parsed = orderSchema.safeParse(incoming);
  if (!parsed.success) {
    res.status(400).json({
      ok: false,
      message: 'Validierung fehlgeschlagen',
      errors: parsed.error.flatten(),
      expectedExample: buildOrderCommitExample(),
    });
    return null;
  }

  const normalizedOrder = {
    ...parsed.data,
    sameContainersForAll: parsed.data.sameContainersForAll === true,
    headerContainers: normalizeContainers(parsed.data.headerContainers, { modeDefault: 'perOrder' }),
    proben: (Array.isArray(parsed.data.proben) ? parsed.data.proben : []).map((probe) => {
      const material = String(probe.material || '').trim();
      const matrixTyp = String(probe.matrixTyp || '').trim();
      const normalizedProbe = {
        ...probe,
        material: material || matrixTyp || '',
        containers: normalizeContainers(probe.containers, { modeDefault: 'perSample' }),
      };
      const onlyMaterial = parsed.data.sameContainersForAll === true;
      const renderedColH = renderColumnHFromProbe(normalizedProbe, {
        onlyMaterial,
        config: quickConfig,
      });
      return {
        ...normalizedProbe,
        gebindeSummary: renderContainersSummary(normalizedProbe.containers, { config: quickConfig }) || undefined,
        materialGebinde: renderedColH || undefined,
      };
    }),
  };

  if (normalizedOrder.sameContainersForAll) {
    normalizedOrder.headerGebindeSummary = renderContainersSummary(normalizedOrder.headerContainers, { config: quickConfig }) || '';
  } else {
    normalizedOrder.headerGebindeSummary = '';
  }

  return normalizedOrder;
}

app.use(express.json({
  type: ['application/json', 'application/*+json'],
}));
app.use(express.static(path.join(__dirname, 'public'), {
  setHeaders(res, filePath) {
    const ext = path.extname(filePath).toLowerCase();
    if (ext === '.html') {
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      return;
    }
    if (ext === '.css') {
      res.setHeader('Content-Type', 'text/css; charset=utf-8');
      return;
    }
    if (ext === '.js') {
      res.setHeader('Content-Type', 'application/javascript; charset=utf-8');
    }
  },
}));
app.get('/packages', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'packages.html'));
});
app.get('/settings', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'settings.html'));
});
app.get('/settings/', (req, res) => {
  res.redirect(302, '/settings');
});
app.get('/single-parameters', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'single-parameters.html'));
});
app.get('/builder-packages', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'builder-packages.html'));
});
app.get('/settings/packages', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'packages.html'));
});

app.get('/api/customers', (req, res) => {
  const customers = listCustomerProfilesAlpha();

  return res.json({
    ok: true,
    customers,
  });
});

app.delete('/api/customers/:id', (req, res) => {
  const id = decodeURIComponent(String(req.params.id || '').trim());
  if (!id) {
    return res.status(400).json({
      ok: false,
      message: 'Kunden-ID fehlt',
    });
  }
  const deleted = deleteCustomerProfileById(id);
  if (!deleted) {
    return res.status(404).json({
      ok: false,
      message: 'Kunde nicht gefunden',
    });
  }
  return res.json({
    ok: true,
    deleted: true,
    customerCount: customerProfilesCache.length,
  });
});

app.post('/api/customers/refresh-from-excel', async (req, res) => {
  try {
    const result = await refreshCustomersFromExcel(getConfig(), new Date());
    return res.json({
      ok: true,
      ...result,
    });
  } catch (error) {
    return res.status(400).json({
      ok: false,
      message: `Kunden aus Excel aktualisieren fehlgeschlagen: ${error.message}`,
    });
  }
});

app.post('/api/state/reset', (req, res) => {
  try {
    if (fs.existsSync(sheetStateCachePath)) {
      fs.unlinkSync(sheetStateCachePath);
    }
    sheetStateCache = null;
    return res.json({
      ok: true,
      cacheReset: true,
    });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      message: `Cache reset fehlgeschlagen: ${error.message}`,
    });
  }
});

app.get('/api/config', (req, res) => {
  const config = getConfig();
  return res.json({
    ok: true,
    canWriteConfig: true,
    config: toPublicConfig(config),
    mode: config.mode,
    excelPath: config.excelPath,
    commitAllowed: isCommitAllowed(),
  });
});

app.post('/api/config', (req, res) => {
  const parsedPayload = configUpdateSchema.safeParse(req.body);
  if (!parsedPayload.success) {
    return res.status(400).json({
      ok: false,
      message: 'Ungueltiges Config-Payload',
      errors: parsedPayload.error.flatten(),
    });
  }

  const payload = { ...parsedPayload.data };
  if (Object.prototype.hasOwnProperty.call(payload, 'quickContainerPlastic')) {
    payload.quickContainerPlastic = normalizeQuickListPayload(payload.quickContainerPlastic);
  }
  if (Object.prototype.hasOwnProperty.call(payload, 'quickContainerGlass')) {
    payload.quickContainerGlass = normalizeQuickListPayload(payload.quickContainerGlass);
  }
  if (Object.prototype.hasOwnProperty.call(payload, 'uiKuerzelPreset')) {
    payload.uiKuerzelPreset = normalizeKuerzelPresetPayload(payload.uiKuerzelPreset);
  }
  const keys = Object.keys(payload);
  if (keys.length === 0) {
    return res.status(400).json({
      ok: false,
      message: 'Kein aenderbares Feld im Payload',
    });
  }

  const invalidKey = keys.find((key) => !allEditableFields.includes(key));
  if (invalidKey) {
    return res.status(400).json({
      ok: false,
      message: `Feld nicht erlaubt: ${invalidKey}`,
    });
  }

  const hasLevel2Change = keys.some((key) => level2Fields.includes(key));
  if (hasLevel2Change) {
    const providedAdminKey = req.get('x-admin-key') || '';
    const expectedAdminKey = process.env.ANNAHME_ADMIN_KEY || '';
    if (!expectedAdminKey || providedAdminKey !== expectedAdminKey) {
      return res.status(403).json({
        ok: false,
        message: 'Admin Key fehlt oder ist ungueltig fuer Advanced Felder',
      });
    }
  }

  const currentConfig = getConfig();
  const mergedConfig = { ...currentConfig, ...payload };
  const parsedMerged = configSchema.safeParse(mergedConfig);
  if (!parsedMerged.success) {
    return res.status(400).json({
      ok: false,
      message: 'Konfiguration ungueltig',
      errors: parsedMerged.error.flatten(),
    });
  }

  saveConfig(parsedMerged.data);
  runtimeConfig = loadConfig();

  return res.json({
    ok: true,
    restartRequired: false,
    config: toPublicConfig(runtimeConfig),
  });
});

app.get('/api/config/ui-kuerzel-preset', (req, res) => {
  const config = getConfig();
  return res.json({
    ok: true,
    uiKuerzelPreset: Array.isArray(config.uiKuerzelPreset)
      ? config.uiKuerzelPreset
      : ['AD', 'DV', 'LB', 'DH', 'SE', 'JO', 'RS', 'KH'],
  });
});

app.post('/api/config/ui-kuerzel-preset', (req, res) => {
  const schema = z.object({
    uiKuerzelPreset: z.array(z.string()),
  }).strict();
  const parsed = schema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({
      ok: false,
      message: 'Ungueltiges Payload fuer uiKuerzelPreset',
      errors: parsed.error.flatten(),
    });
  }

  const normalized = normalizeKuerzelPresetPayload(parsed.data.uiKuerzelPreset);
  const mergedConfig = {
    ...getConfig(),
    uiKuerzelPreset: normalized,
  };
  const parsedMerged = configSchema.safeParse(mergedConfig);
  if (!parsedMerged.success) {
    return res.status(400).json({
      ok: false,
      message: 'Konfiguration ungueltig',
      errors: parsedMerged.error.flatten(),
    });
  }

  saveConfig(parsedMerged.data);
  runtimeConfig = loadConfig();
  return res.json({
    ok: true,
    uiKuerzelPreset: runtimeConfig.uiKuerzelPreset,
    config: toPublicConfig(runtimeConfig),
  });
});

app.get('/api/config/validate', async (req, res) => {
  const querySchema = z.object({
    excelPath: z.string().trim().optional(),
  });
  const parsedQuery = querySchema.safeParse(req.query);
  if (!parsedQuery.success) {
    return res.status(400).json({
      ok: false,
      message: 'Ungueltige Query Parameter',
      errors: parsedQuery.error.flatten(),
    });
  }

  const config = getConfig();
  const excelPath = parsedQuery.data.excelPath || config.excelPath;
  const absoluteExcelPath = resolveExcelPath(excelPath);
  const warnings = [];
  const errors = [];

  if (!fs.existsSync(absoluteExcelPath)) {
    errors.push(`Excel-Datei nicht gefunden: ${absoluteExcelPath}`);
    return res.status(400).json({
      ok: false,
      excelPath,
      absoluteExcelPath,
      warnings,
      errors,
    });
  }

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(absoluteExcelPath);

    const templateSheet = workbook.getWorksheet('Vorlagen');
    if (!templateSheet) {
      errors.push('Sheet Vorlagen nicht gefunden');
    } else {
      const packages = await importPackagesFromExcel(absoluteExcelPath, 'Vorlagen');
      if (packages.length < 1) {
        errors.push('Sheet Vorlagen enthaelt keine gueltigen Pakete');
      }
    }

    const yearSheetName = getYearSheetName(config);
    const yearSheet = workbook.getWorksheet(yearSheetName);
    if (!yearSheet) {
      warnings.push(`Jahresblatt ${yearSheetName} nicht gefunden`);
    }
  } catch (error) {
    errors.push(`Excel-Datei kann nicht gelesen werden: ${error.message}`);
  }

  return res.status(errors.length > 0 ? 400 : 200).json({
    ok: errors.length === 0,
    excelPath,
    absoluteExcelPath,
    warnings,
    errors,
  });
});

app.get('/api/backups/validate', (req, res) => {
  const querySchema = z.object({
    dir: z.string().optional(),
  });
  const parsedQuery = querySchema.safeParse(req.query);
  if (!parsedQuery.success) {
    return res.status(400).json({
      ok: false,
      writable: false,
      absolutePath: null,
      message: 'Ungueltige Query Parameter',
      errors: parsedQuery.error.flatten(),
    });
  }

  const config = getConfig();
  const dirInput = typeof parsedQuery.data.dir === 'string' ? parsedQuery.data.dir : config.backupDir;
  const rawDir = String(dirInput || '').trim();
  if (!rawDir) {
    return res.status(200).json({
      ok: false,
      writable: false,
      absolutePath: null,
      message: 'Bitte einen Backup-Ordner angeben.',
    });
  }

  const absolutePath = resolveBackupDir(__dirname, rawDir);
  const writableCheck = ensureBackupDirWritable(absolutePath);
  return res.status(200).json({
    ok: writableCheck.ok,
    writable: writableCheck.ok,
    absolutePath,
    message: writableCheck.ok ? 'Backup-Ordner ist vorhanden und beschreibbar.' : `Backup-Ordner ist nicht beschreibbar: ${writableCheck.message}`,
  });
});

app.post('/api/backups/create', (req, res) => {
  const schema = z.object({
    force: z.boolean().optional(),
  }).strict();
  const parsedBody = schema.safeParse(req.body || {});
  if (!parsedBody.success) {
    return res.status(400).json({
      ok: false,
      reason: 'invalid_payload',
      message: 'Ungueltiges Payload',
      errors: parsedBody.error.flatten(),
    });
  }

  const config = getConfig();
  const backup = createManualBackup({
    config,
    excelPath: config.excelPath,
    rootDir: __dirname,
  }, {
    force: parsedBody.data.force === true,
  });

  if (!backup.created) {
    return res.status(400).json({
      ok: false,
      reason: backup.reason || 'backup_failed',
      message: backup.message || 'Backup konnte nicht erstellt werden.',
      cleanupDeleted: Array.isArray(backup.cleanupDeleted) ? backup.cleanupDeleted : [],
    });
  }

  return res.json({
    ok: true,
    created: true,
    reason: backup.reason,
    fileName: backup.fileName,
    absoluteBackupPath: backup.absoluteBackupPath,
    cleanupDeleted: Array.isArray(backup.cleanupDeleted) ? backup.cleanupDeleted : [],
  });
});

app.get('/api/system/pick-backup-dir', (_req, res) => {
  if (process.platform !== 'win32') {
    return res.status(400).json({
      ok: false,
      message: 'Nur unter Windows verfuegbar.',
    });
  }

  const scriptPath = path.join(__dirname, 'scripts', 'pick_folder.ps1');
  const result = spawnSync('powershell.exe', [
    '-NoLogo',
    '-NoProfile',
    '-NonInteractive',
    '-ExecutionPolicy',
    'Bypass',
    '-STA',
    '-File',
    scriptPath,
  ], {
    windowsHide: true,
    encoding: 'utf8',
  });

  if (result.error || result.status !== 0) {
    const detail = result.error
      ? result.error.message
      : (String(result.stderr || '').trim() || `exit status ${result.status}`);
    return res.status(500).json({
      ok: false,
      message: `Ordnerauswahl fehlgeschlagen: ${detail}`,
    });
  }

  const selectedPath = String(result.stdout || '').trim();
  if (!selectedPath) {
    return res.status(200).json({
      ok: true,
      canceled: true,
      selectedPath: null,
    });
  }

  return res.status(200).json({
    ok: true,
    canceled: false,
    selectedPath,
  });
});

app.get('/api/single-parameter-catalog', (_req, res) => {
  try {
    const catalog = loadSingleParameterCatalog();
    const updatedAt = String(catalog?.updatedAt || '').trim() || singleParameterCatalogUpdatedAt || null;
    return res.status(200).json({
      ok: true,
      catalog,
      updatedAt,
    });
  } catch (_error) {
    return res.status(500).json({
      ok: false,
      message: 'Single Parameter Katalog konnte nicht geladen werden.',
    });
  }
});

app.post('/api/single-parameter-catalog', (req, res) => {
  try {
    const payload = req.body && typeof req.body === 'object' ? req.body : {};
    const catalog = payload.catalog && typeof payload.catalog === 'object' ? payload.catalog : null;
    if (!catalog) {
      return res.status(400).json({
        ok: false,
        message: 'catalog fehlt.',
      });
    }
    if (typeof catalog.version !== 'number' || Number.isNaN(catalog.version)) {
      return res.status(400).json({
        ok: false,
        message: 'catalog.version muss eine Zahl sein.',
      });
    }
    if (!Array.isArray(catalog.parameters)) {
      return res.status(400).json({
        ok: false,
        message: 'catalog.parameters muss ein Array sein.',
      });
    }
    for (const param of catalog.parameters) {
      const key = String(param?.key || '').trim();
      const label = String(param?.label || '').trim();
      if (!key || !label) {
        return res.status(400).json({
          ok: false,
          message: 'Jeder Parameter braucht mindestens key und label.',
        });
      }
      if (param.functionGroup !== undefined && param.functionGroup !== null) {
        const value = String(param.functionGroup || '').trim();
        if (!value) param.functionGroup = null;
      }
    }

    const nowIso = new Date().toISOString();
    catalog.updatedAt = nowIso;
    fs.writeFileSync(singleParameterCatalogPath, `${JSON.stringify(catalog, null, 2)}\n`, 'utf-8');
    singleParameterCatalogCache = catalog;
    singleParameterCatalogUpdatedAt = nowIso;

    return res.status(200).json({
      ok: true,
      catalog,
      updatedAt: nowIso,
    });
  } catch (_error) {
    return res.status(500).json({
      ok: false,
      message: 'Single Parameter Katalog konnte nicht gespeichert werden.',
    });
  }
});

app.get('/api/packages', (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');
    const packages = readPackages({ forceReload: true });
    return res.json(packages);
  } catch (error) {
    return res.status(500).json({
      ok: false,
      message: error.message,
    });
  }
});

app.get('/api/builder-packages', (_req, res) => {
  try {
    const store = loadBuilderPackagesStore();
    return res.status(200).json({
      ok: true,
      packages: store.packages,
      updatedAt: store.updatedAt || null,
    });
  } catch (_error) {
    return res.status(500).json({
      ok: false,
      message: 'Builder-Pakete konnten nicht geladen werden.',
    });
  }
});

app.post('/api/builder-packages', (req, res) => {
  try {
    const payload = req.body && typeof req.body === 'object' ? req.body : {};
    const incoming = (payload.package && typeof payload.package === 'object')
      ? payload.package
      : payload;
    if (!incoming) {
      return res.status(400).json({
        ok: false,
        message: 'package fehlt.',
      });
    }
    const nowIso = new Date().toISOString();
    const name = String(incoming.name || '').trim();
    if (!name) {
      return res.status(400).json({
        ok: false,
        message: 'name ist erforderlich.',
      });
    }
    const id = String(incoming.id || '').trim() || `bp_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
    const definition = incoming.definition && typeof incoming.definition === 'object'
      ? incoming.definition
      : { toggles: {} };

    const store = loadBuilderPackagesStore();
    const packages = Array.isArray(store.packages) ? [...store.packages] : [];
    const existingIndex = packages.findIndex((entry) => String(entry.id || '').trim() === id);
    const existing = existingIndex >= 0 ? packages[existingIndex] : null;
    const record = normalizeBuilderPackage({
      id,
      type: 'builder',
      name,
      createdAt: existing?.createdAt || String(incoming.createdAt || '').trim() || nowIso,
      updatedAt: nowIso,
      definition,
    });
    if (!record) {
      return res.status(400).json({
        ok: false,
        message: 'Builder-Paket ist ungültig.',
      });
    }
    if (existingIndex >= 0) {
      packages[existingIndex] = record;
    } else {
      packages.push(record);
    }
    const saved = saveBuilderPackagesStore({
      ...store,
      updatedAt: nowIso,
      packages,
    });
    return res.status(200).json({
      ok: true,
      package: record,
      packages: saved.packages,
      updatedAt: saved.updatedAt || nowIso,
    });
  } catch (_error) {
    return res.status(500).json({
      ok: false,
      message: 'Builder-Paket konnte nicht gespeichert werden.',
    });
  }
});

app.delete('/api/builder-packages/:id', (req, res) => {
  try {
    const id = String(req.params.id || '').trim();
    if (!id) {
      return res.status(400).json({
        ok: false,
        message: 'id fehlt.',
      });
    }
    const store = loadBuilderPackagesStore();
    const packages = Array.isArray(store.packages) ? store.packages : [];
    const nextPackages = packages.filter((entry) => String(entry.id || '').trim() !== id);
    if (nextPackages.length === packages.length) {
      return res.status(404).json({
        ok: false,
        message: 'Builder-Paket nicht gefunden.',
      });
    }
    const nowIso = new Date().toISOString();
    const saved = saveBuilderPackagesStore({
      ...store,
      updatedAt: nowIso,
      packages: nextPackages,
    });
    return res.status(200).json({
      ok: true,
      updatedAt: saved.updatedAt || nowIso,
    });
  } catch (_error) {
    return res.status(500).json({
      ok: false,
      message: 'Builder-Paket konnte nicht gelöscht werden.',
    });
  }
});

app.post('/api/packages/import', createImportPackagesHandler({
  getConfig,
  resolveExcelPath,
  invalidatePackagesCache,
  importPackagesFromExcel,
  writePackages,
  readPackages,
}));

app.post('/api/writer/login', (req, res) => {
  const config = getConfig();
  if (config.mode !== 'writer') {
    return res.status(400).json({
      ok: false,
      message: 'Writer login ist nur im Modus writer verfuegbar',
    });
  }

  const loginSchema = z.object({ token: z.string().min(1) });
  const parsed = loginSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({
      ok: false,
      message: 'Token fehlt',
    });
  }

  if (parsed.data.token !== config.writerToken) {
    return res.status(401).json({
      ok: false,
      message: 'Token ungueltig',
    });
  }

  return res.json({ ok: true });
});

app.post('/api/com-test', async (req, res) => {
  const schema = z.object({
    cellPath: z.string().trim().min(1).default('2026!Z1'),
    value: z.string().optional(),
  });
  const parsed = schema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({
      ok: false,
      message: 'cellPath fehlt oder ist ungueltig',
      errors: parsed.error.flatten(),
    });
  }

  try {
    const config = getConfig();
    const result = await writeComTestCell({
      rootDir: __dirname,
      excelPath: config.excelPath,
      cellPath: parsed.data.cellPath,
      value: parsed.data.value,
    });

    console.log(`[api/com-test] ok=true cellPath=${parsed.data.cellPath}`);
    return res.json({
      ok: true,
      writtenValue: result.writtenValue,
      saved: result.saved === true,
      readbackValue: result.readbackValue,
      mode: result.mode || null,
      workbookFullName: result.workbookFullName || null,
      workbookName: result.workbookName || null,
      excelVersion: result.excelVersion || null,
      excelHwnd: Number.isInteger(result.excelHwnd) ? result.excelHwnd : null,
    });
  } catch (error) {
    console.error(`[api/com-test] ok=false cellPath=${parsed.data.cellPath} error=${error.message}`);
    return res.status(400).json({
      ok: false,
      message: error.message,
      saved: false,
    });
  }
});

app.post('/api/com-test/umlaut', async (req, res) => {
  const umlautText = 'Umlaute: äöüÄÖÜß';
  try {
    const config = getConfig();
    const result = await writeComTestCell({
      rootDir: __dirname,
      excelPath: config.excelPath,
      cellPath: '2026!Z2',
      value: umlautText,
    });
    return res.json({
      ok: true,
      cellPath: '2026!Z2',
      writtenValue: result.writtenValue,
      readbackValue: result.readbackValue,
      saved: result.saved === true,
    });
  } catch (error) {
    return res.status(400).json({
      ok: false,
      message: error.message,
      saved: false,
    });
  }
});

app.get('/api/order/schema', (req, res) => {
  return res.json({
    ok: true,
    commitExample: buildOrderCommitExample(),
    schema: buildOrderSchemaInfo(),
    uiModel: {
      defaults: {
        eilig: false,
        material: 'Boden',
        projektName: '',
        ansprechpartner: '',
        email: '',
        adresseBlock: '',
        kopfBemerkung: '',
        kuerzel: '',
      },
      sampleDefaults: {
        gewicht: '',
        geruch: '',
        bemerkung: '',
        tiefeOderVolumen: '',
      },
    },
  });
});

app.get('/api/reportdb/index', (req, res) => {
  try {
    const rawLimit = Number.parseInt(String(req.query.limit || '50'), 10);
    const limit = Number.isFinite(rawLimit) && rawLimit > 0 ? Math.min(rawLimit, 500) : 50;
    const search = String(req.query.search || '').trim().toLowerCase();
    const includeDeletedRaw = String(req.query.includeDeleted || '0').trim().toLowerCase();
    const includeDeleted = includeDeletedRaw === '1' || includeDeletedRaw === 'true' || includeDeletedRaw === 'yes';

    let rows = readReportDbIndexRows();
    if (!includeDeleted) {
      rows = rows.filter((row) => String(row.deleted || '0') !== '1');
    }
    if (search) {
      rows = rows.filter((row) => {
        const orderNo = String(row.orderNo || '').toLowerCase();
        const clientName = String(row.clientName || '').toLowerCase();
        const projectName = String(row.projectName || '').toLowerCase();
        return orderNo.includes(search) || clientName.includes(search) || projectName.includes(search);
      });
    }
    rows.sort((a, b) => String(b.committedAt || '').localeCompare(String(a.committedAt || '')));
    return res.json({
      ok: true,
      rows: rows.slice(0, limit),
    });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      message: `ReportDB Index konnte nicht geladen werden: ${error.message}`,
    });
  }
});

app.post('/api/reportdb/delete-order', (req, res) => {
  try {
    const orderNo = String(req.body?.orderNo || '').trim();
    if (!orderNo) {
      return res.status(400).json({
        ok: false,
        message: 'orderNo ist erforderlich',
      });
    }
    const committedAt = new Date().toISOString();
    const commitId = createReportDbCommitId();
    appendReportDbEvent({
      schemaVersion: 1,
      eventType: 'delete',
      commitId,
      committedAt,
      orderNo,
    });
    markReportDbOrderDeleted(orderNo, committedAt, commitId);
    return res.json({ ok: true });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      message: `ReportDB Löschung fehlgeschlagen: ${error.message}`,
    });
  }
});

app.post('/api/order/draft', (req, res) => {
  const order = parseOrderOrRespond(req, res);
  if (!order) {
    return;
  }

  try {
    upsertCustomerProfileFromOrder(order, new Date());
  } catch (error) {
    console.warn(`[customers] draft upsert failed: ${error.message}`);
  }

  return res.json({
    ...buildOrderPreview(order),
    operation: 'draft',
  });
});

app.post('/api/order/preview', async (req, res) => {
  const order = parseOrderOrRespond(req, res);
  if (!order) {
    return;
  }

  try {
    const config = getConfig();
    const now = new Date();
    const state = await buildCommitPreviewState(order, config, now);
    const firstLab = Array.isArray(state.preview.labNumberPreview) && state.preview.labNumberPreview.length > 0
      ? state.preview.labNumberPreview[0]
      : null;

    return res.json({
      ok: true,
      ...state.preview,
      warnings: state.commitWarnings,
      computedOrderNo: state.computedOrderNo,
      labNumberPreview: firstLab,
      todayPrefix: state.todayPrefix,
      maxOrderSeqToday: state.maxOrderSeqToday,
      nextSeq: state.nextSeq,
      lastLabNo: state.cache ? Number.parseInt(String(state.cache.lastLabNo || 0), 10) : null,
      lastUsedRow: state.cache ? Number.parseInt(String(state.cache.lastUsedRow || 0), 10) : null,
      operation: 'preview',
    });
  } catch (error) {
    return res.status(400).json({
      ok: false,
      message: `Preview fehlgeschlagen: ${error.message}`,
    });
  }
});

app.post('/api/order/commit', async (req, res) => {
  const config = getConfig();
  if (config.mode === 'client') {
    return res.status(403).json({
      ok: false,
      message: 'Commit API ist im Modus client deaktiviert',
    });
  }

  if (config.mode === 'writer') {
    const isUiRequest = req.get('x-ui-request') === '1';
    const token = req.get('x-writer-token') || '';

    if (!isUiRequest) {
      return res.status(403).json({
        ok: false,
        message: 'Commit ist im Modus writer nur fuer UI Requests erlaubt',
      });
    }

    if (token !== config.writerToken) {
      return res.status(401).json({
        ok: false,
        message: 'Writer Token fehlt oder ist ungueltig',
      });
    }
  }

  const order = parseOrderOrRespond(req, res);
  if (!order) {
    return;
  }

  const clientRequestId = readClientRequestId(req.body?.clientRequestId);
  if (clientRequestId) {
    pruneCommitRequestStore();
    const existing = commitRequestStore.get(clientRequestId);
    if (existing) {
      if (existing.state === 'done' && existing.response) {
        return res.json({
          ...existing.response,
          duplicateIgnored: true,
          message: 'duplicate ignored',
          clientRequestId,
        });
      }
      return res.json({
        ok: true,
        duplicateIgnored: true,
        message: 'duplicate ignored (processing)',
        clientRequestId,
      });
    }
    commitRequestStore.set(clientRequestId, {
      ts: Date.now(),
      state: 'processing',
      response: null,
    });
  }

  const nowForCommit = new Date();
  const previewState = await buildCommitPreviewState(order, config, nowForCommit);
  const {
    absoluteExcelPath,
    orderForWrite,
    commitWarnings,
    cache,
    preview,
    todayPrefix: todayPrefixForCommit,
    maxOrderSeqToday: maxOrderSeqTodayFromCache,
    nextSeq: nextSeqFromCache,
    computedOrderNo: computedOrderNoFromCache,
    appendRow: appendRowFromCache,
    startLabNo: startLabNoFromCache,
    cacheHint,
  } = previewState;

  const backup = ensureBackupBeforeCommit({
    config,
    excelPath: config.excelPath,
    rootDir: __dirname,
  });

  let writeResult = null;
  const usedBackend = resolveCommitWriterBackend(config);
  try {
    writeResult = await writeOrderBlock({
      backend: usedBackend,
      config,
      rootDir: __dirname,
      excelPath: config.excelPath,
      order: orderForWrite,
      termin: preview.termin,
      now: nowForCommit,
      cacheHint,
    });
  } catch (error) {
    if (clientRequestId) {
      commitRequestStore.delete(clientRequestId);
    }
    const fullMessage = `Writer fehlgeschlagen (${usedBackend}): ${error.message}`;
    const parsedWriterError = extractWriterDebug(error.message);
    const writerErrorCode = String(parsedWriterError?.debug?.code || '').trim();
    const isWorkbookReadonly = writerErrorCode === 'WORKBOOK_READONLY';
    const normalizedUserMessage = isExcelNotOpenMessage(parsedWriterError.userMessage)
      ? EXCEL_NOT_OPEN_USER_MESSAGE
      : (isWorkbookReadonly ? WORKBOOK_READONLY_USER_MESSAGE : parsedWriterError.userMessage);
    const userMessage = normalizedUserMessage === EXCEL_NOT_OPEN_USER_MESSAGE
      ? EXCEL_NOT_OPEN_USER_MESSAGE
      : `Writer fehlgeschlagen (${usedBackend}): ${normalizedUserMessage}`;
    const debug = normalizedUserMessage === EXCEL_NOT_OPEN_USER_MESSAGE ? undefined : (parsedWriterError.debug || undefined);
    const statusCode = isWorkbookReadonly ? 409 : 400;
    return res.status(statusCode).json({
      ok: false,
      message: fullMessage,
      userMessage,
      debug,
      clientRequestId,
    });
  }

  const sampleNos = Array.isArray(writeResult.sampleNos)
    ? writeResult.sampleNos
    : (Number.isInteger(writeResult.startLabNo)
      ? Array.from({ length: orderForWrite.proben.length }, (_unused, idx) => writeResult.startLabNo + idx)
      : []);
  const orderNo = writeResult.orderNo || preview.orderNumberPreview || null;
  const ersteProbennr = sampleNos.length > 0 ? sampleNos[0] : null;
  const letzteProbennr = sampleNos.length > 0 ? sampleNos[sampleNos.length - 1] : null;
  const fallbackEndRow = Number.isInteger(writeResult.appendRow) ? writeResult.appendRow + orderForWrite.proben.length : null;
  const endRowRange = writeResult.endRowRange || (fallbackEndRow ? `A${writeResult.appendRow}:J${fallbackEndRow}` : null);
  const saved = writeResult.saved !== false;
  const todayPrefix = typeof writeResult.todayPrefix === 'string'
    ? writeResult.todayPrefix
    : (orderNo && String(orderNo).length >= 7 ? String(orderNo).slice(0, 7) : null);
  const maxOrderSeqToday = Number.isInteger(writeResult.maxOrderSeqToday)
    ? writeResult.maxOrderSeqToday
    : null;
  const nextSeq = Number.isInteger(writeResult.nextSeq)
    ? writeResult.nextSeq
    : (orderNo && String(orderNo).length >= 2 ? Number.parseInt(String(orderNo).slice(-2), 10) : null);
  const computedOrderNo = writeResult.computedOrderNo || orderNo || null;
  const previewBlock = {
    ...preview,
    todayPrefix,
    maxOrderSeqToday,
    nextSeq,
    computedOrderNo,
  };

  if (cache && writeResult.saved !== false) {
    const fileMetaAfterWrite = getExcelFileMeta(absoluteExcelPath);
    const resolvedAppendRow = Number.isInteger(writeResult.appendRow) ? writeResult.appendRow : appendRowFromCache;
    const resolvedEndRow = Number.isInteger(resolvedAppendRow) ? resolvedAppendRow + orderForWrite.proben.length : null;
    const resolvedLastLab = Number.isInteger(letzteProbennr) ? letzteProbennr : cache.lastLabNo;
    const resolvedTodayPrefix = todayPrefix || todayPrefixForCommit;
    const resolvedSeq = Number.isInteger(nextSeq) ? nextSeq : maxOrderSeqTodayFromCache;
    const updatedCache = {
      ...cache,
      fileMtimeMs: fileMetaAfterWrite.fileMtimeMs,
      excelFileSize: fileMetaAfterWrite.excelFileSize,
      lastWriteTime: fileMetaAfterWrite.lastWriteTime,
      lastUsedRow: Number.isInteger(resolvedEndRow) ? (resolvedEndRow + 1) : cache.lastUsedRow,
      lastLabNo: Number.isInteger(resolvedLastLab) ? resolvedLastLab : cache.lastLabNo,
      orderSeqByPrefix: {
        ...normalizeOrderSeqByPrefix(cache.orderSeqByPrefix),
        [resolvedTodayPrefix]: Math.max(
          Number(cache.orderSeqByPrefix?.[resolvedTodayPrefix] || 0),
          Number.isInteger(resolvedSeq) ? resolvedSeq : 0,
        ),
      },
      updatedAt: new Date().toISOString(),
    };
    sheetStateCache = updatedCache;
    persistSheetStateCache(updatedCache, false);
  }

  console.log(`[commit] writer=${usedBackend} order=${orderNo || 'n/a'} rows=${endRowRange || 'n/a'}`);
  let reportDbSaved = false;
  let reportDbError = '';
  try {
    saveReportDbUpsertForCommit({
      order: orderForWrite,
      orderNo,
      sampleNos,
      committedAt: nowForCommit.toISOString(),
    });
    reportDbSaved = true;
  } catch (error) {
    reportDbError = String(error.message || 'unbekannter Fehler');
    console.warn(`[reportdb] save failed for order ${orderNo || 'n/a'}: ${reportDbError}`);
  }

  const responsePayload = {
    ok: true,
    writer: usedBackend,
    saved,
    mode: writeResult.mode || null,
    workbookFullName: writeResult.workbookFullName || null,
    workbookName: writeResult.workbookName || null,
    excelVersion: writeResult.excelVersion || null,
    excelHwnd: Number.isInteger(writeResult.excelHwnd) ? writeResult.excelHwnd : null,
    readbackRows: Array.isArray(writeResult.readbackRows) ? writeResult.readbackRows : [],
    orderNo,
    auftragsnummer: orderNo,
    sampleNos,
    ersteProbennr,
    letzteProbennr,
    endRowRange,
    warnings: commitWarnings,
    ...preview,
    preview: previewBlock,
    operation: 'commit',
    backup,
    writerBackend: usedBackend,
    writerResult: writeResult,
    reportDbSaved,
    reportDbError: reportDbSaved ? undefined : reportDbError,
    clientRequestId: clientRequestId || undefined,
  };

  if (clientRequestId) {
    commitRequestStore.set(clientRequestId, {
      ts: Date.now(),
      state: 'done',
      response: responsePayload,
    });
    pruneCommitRequestStore();
  }

  try {
    upsertCustomerProfileFromOrder(orderForWrite, new Date());
  } catch (error) {
    console.warn(`[customers] upsert failed: ${error.message}`);
  }

  return res.json(responsePayload);
});

app.post('/api/order', (req, res) => {
  const order = parseOrderOrRespond(req, res);
  if (!order) {
    return;
  }

  return res.json({
    ...buildOrderPreview(order),
    operation: 'draft',
  });
});

app.listen(port, () => {
  const config = getConfig();
  console.log(`Server laeuft auf http://localhost:${port} (mode=${config.mode}, excelPath=${config.excelPath})`);
  if (resolveCommitWriterBackend(config) === 'com') {
    const workerClient = getComWorkerClient(__dirname);
    workerClient.start()
      .then(() => {
        console.log('[com-worker] started');
        const warmupPayload = {
          __warmup: true,
          excelPath: resolveExcelPath(config.excelPath),
          yearSheetName: getYearSheetName(config),
          allowAutoOpenExcel: false,
        };
        return workerClient.request(warmupPayload, {
          timeoutMs: 10000,
          retryOnFailure: false,
        });
      })
      .then(() => {
        console.log('[com-worker] warmup complete');
      })
      .catch((error) => {
        console.warn(`[com-worker] start failed: ${error.message}`);
      });
  }
  ensureSheetStateCache(config, new Date())
    .then((cache) => {
      console.log(
        `[sheet-cache] ready yearSheet=${cache.yearSheetName} lastUsedRow=${cache.lastUsedRow} lastLabNo=${cache.lastLabNo}`,
      );
    })
    .catch((error) => {
      console.warn(`[sheet-cache] init failed: ${error.message}`);
    });
});



