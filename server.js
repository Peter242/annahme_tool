const fs = require('fs');
const path = require('path');
const express = require('express');
const ExcelJS = require('exceljs');
const { z } = require('zod');
const { makeOrderNumber, nextLabNumbers } = require('./src/numbering');
const { ensureBackupBeforeCommit } = require('./src/backup');
const { importPackagesFromExcel } = require('./src/packages/importFromExcel');
const { readPackages, writePackages } = require('./src/packages/store');
const { writeOrderBlock } = require('./src/orderWriter');
const { writeComTestCell } = require('./src/writers/comTestWriter');
const { calculateTermin } = require('./src/termin');

const configSchema = z
  .object({
    port: z.number().int().min(1).max(65535),
    mode: z.enum(['single', 'writer', 'client']),
    writerBackend: z.enum(['exceljs', 'com', 'comExceljs']),
    excelPath: z.string().trim().min(1),
    yearSheetName: z.string(),
    writerHost: z.string().trim(),
    writerToken: z.string(),
    backupEnabled: z.boolean(),
    backupPolicy: z.enum(['daily', 'interval']),
    backupIntervalMinutes: z.number().int().positive(),
    backupRetentionDays: z.number().int().nonnegative(),
    backupZip: z.boolean(),
    uiShowPackagePreview: z.boolean(),
    uiDefaultEilig: z.enum(['ja', 'nein']),
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
  writerHost: 'http://localhost:3000',
  writerToken: 'dev-writer-token',
  backupEnabled: true,
  backupPolicy: 'daily',
  backupIntervalMinutes: 60,
  backupRetentionDays: 14,
  backupZip: false,
  uiShowPackagePreview: true,
  uiDefaultEilig: 'ja',
};

const configPath = path.join(__dirname, 'config.json');

function loadConfig() {
  let rawConfig = { ...defaultConfig };

  if (fs.existsSync(configPath)) {
    const fileContent = fs.readFileSync(configPath, 'utf-8');
    rawConfig = { ...defaultConfig, ...JSON.parse(fileContent) };
  }

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
    projekt: 'Projekt Muster',
    projektname: 'Projekt Muster Name',
    projektnummer: 'P-2026-001',
    ansprechpartner: 'Max Mustermann',
    email: 'max@example.com',
    kuerzel: 'MM',
    eilig: false,
    probenEingangDatum: '2026-02-14',
    proben: [
      {
        probenbezeichnung: 'Probe 1',
        matrixTyp: 'Boden',
        gewicht: 1.2,
        paketKey: 'DepV/DepV DK0',
      },
    ],
  };
}

function buildOrderSchemaInfo() {
  return {
    matrixTypEnum: ['Boden', 'Wasser', 'Luft'],
    fields: {
      kunde: 'string (required)',
      projekt: 'string (required)',
      projektnummer: 'string (required)',
      projektname: 'string (optional)',
      ansprechpartner: 'string (optional)',
      email: 'string (optional)',
      kuerzel: 'string (optional)',
      eilig: 'boolean (required)',
      probenEingangDatum: 'string YYYY-MM-DD (required unless probeNochNichtDa/sampleNotArrived=true)',
      probeNochNichtDa: 'boolean (optional)',
      sampleNotArrived: 'boolean (optional)',
      proben: 'array(min 1) of sample objects',
    },
    sampleFields: {
      probenbezeichnung: 'string (required)',
      matrixTyp: 'enum: Boden | Wasser | Luft (required)',
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
    },
  };
}

const level1Fields = [
  'excelPath',
  'yearSheetName',
  'backupEnabled',
  'backupPolicy',
  'backupIntervalMinutes',
  'backupRetentionDays',
  'backupZip',
  'uiShowPackagePreview',
  'uiDefaultEilig',
];
const level2Fields = ['mode', 'writerHost', 'writerToken', 'writerBackend'];
const allEditableFields = [...level1Fields, ...level2Fields];

const configUpdateSchema = z.object({
  excelPath: z.string().trim().min(1).optional(),
  yearSheetName: z.string().optional(),
  backupEnabled: z.boolean().optional(),
  backupPolicy: z.enum(['daily', 'interval']).optional(),
  backupIntervalMinutes: z.number().int().positive().optional(),
  backupRetentionDays: z.number().int().nonnegative().optional(),
  backupZip: z.boolean().optional(),
  uiShowPackagePreview: z.boolean().optional(),
  uiDefaultEilig: z.enum(['ja', 'nein']).optional(),
  mode: z.enum(['single', 'writer', 'client']).optional(),
  writerBackend: z.enum(['exceljs', 'com', 'comExceljs']).optional(),
  writerHost: z.string().trim().optional(),
  writerToken: z.string().optional(),
}).strict();

const sampleSchema = z
  .object({
    probenbezeichnung: z.string().trim().min(1, 'Probenbezeichnung ist erforderlich'),
    matrixTyp: z.enum(['Boden', 'Wasser', 'Luft']),
    gewicht: z.number().positive().optional(),
    gewichtEinheit: z.string().trim().optional(),
    volumen: z.number().positive().optional(),
    packageId: z.string().trim().optional(),
    paketKey: z.string().trim().optional(),
    parameterTextPreview: z.string().optional(),
    tiefeVolumen: z.union([z.string(), z.number()]).optional(),
    tiefeOderVolumen: z.string().trim().optional(),
    geruch: z.string().trim().optional(),
    geruchAuffaelligkeit: z.string().trim().optional(),
    bemerkung: z.string().trim().optional(),
    materialGebinde: z.string().optional(),
    material: z.string().optional(),
    gebinde: z.string().optional(),
  });

const orderSchema = z
  .object({
    kunde: z.string().trim().min(1, 'Kunde ist erforderlich'),
    projekt: z.string().trim().min(1, 'Projekt ist erforderlich'),
    projektnummer: z.string().trim().min(1, 'Projektnummer ist erforderlich'),
    auftragsnotiz: z.string().optional(),
    pbTyp: z.enum(['PB', 'AI', 'AKN']).optional(),
    auftraggeberKurz: z.string().optional(),
    ansprechpartner: z.string().optional(),
    email: z.string().optional(),
    projektname: z.string().optional(),
    probenahmedatum: z.string().optional(),
    erfasstKuerzel: z.string().optional(),
    kuerzel: z.string().optional(),
    terminDatum: z.string().optional(),
    eilig: z.boolean(),
    probeNochNichtDa: z.boolean().optional().default(false),
    sampleNotArrived: z.boolean().optional().default(false),
    probenEingangDatum: z
      .string()
      .date('ProbenEingangDatum muss ein gueltiges Datum sein (YYYY-MM-DD)')
      .optional(),
    proben: z.array(sampleSchema).min(1, 'Mindestens eine Probe ist erforderlich'),
  })
  .superRefine((order, ctx) => {
    const sampleNotArrived = order.probeNochNichtDa || order.sampleNotArrived === true;

    if (!sampleNotArrived && !order.probenEingangDatum) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ['probenEingangDatum'],
        message: 'ProbenEingangDatum ist erforderlich, wenn die Probe schon da ist',
      });
    }

    if (sampleNotArrived && order.probenEingangDatum) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ['probenEingangDatum'],
        message: 'ProbenEingangDatum darf nicht gesetzt sein, wenn Probe noch nicht da aktiviert ist',
      });
    }
  });

function isCommitAllowed() {
  return getConfig().mode !== 'client';
}

function buildOrderPreview(order) {
  const packages = readPackages();
  const packageById = new Map(packages.map((pkg) => [pkg.id, pkg]));
  const vorschau = {
    ...order,
    proben: order.proben.map((probe) => {
      const packageTemplate = probe.packageId ? packageById.get(probe.packageId) : null;
      const renderedTextD = packageTemplate ? packageTemplate.text : (probe.parameterTextPreview || '');
      return {
        ...probe,
        parameterTextPreview: renderedTextD,
        renderedTextD,
      };
    }),
  };

  const sampleNotArrived = order.probeNochNichtDa || order.sampleNotArrived === true;
  const termin = sampleNotArrived ? null : calculateTermin(order.probenEingangDatum, order.eilig);

  const xy = 1;
  const lastLab = 26203;
  const orderNumberPreview = order.probenEingangDatum ? makeOrderNumber(order.probenEingangDatum, xy) : null;
  const labNumberPreview = nextLabNumbers(lastLab, order.proben.length);

  return {
    ok: true,
    vorschau,
    termin,
    orderNumberPreview,
    labNumberPreview,
  };
}

function parseOrderOrRespond(req, res) {
  const parsed = orderSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({
      ok: false,
      message: 'Validierung fehlgeschlagen',
      errors: parsed.error.flatten(),
      expectedExample: buildOrderCommitExample(),
    });
    return null;
  }

  return parsed.data;
}

function logCommit(entry) {
  const logsDir = path.join(__dirname, 'logs');
  const logPath = path.join(logsDir, 'commit-log.jsonl');
  fs.mkdirSync(logsDir, { recursive: true });
  fs.appendFileSync(logPath, `${JSON.stringify(entry)}\n`, 'utf-8');
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

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

  const payload = parsedPayload.data;
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

app.get('/api/packages', (req, res) => {
  try {
    const packages = readPackages();
    return res.json(packages);
  } catch (error) {
    return res.status(500).json({
      ok: false,
      message: error.message,
    });
  }
});

app.post('/api/packages/import', async (req, res) => {
  try {
    const config = getConfig();
    const excelPath = resolveExcelPath(config.excelPath);
    const packages = await importPackagesFromExcel(excelPath, 'Vorlagen');
    writePackages(packages);
    return res.json({ count: packages.length });
  } catch (error) {
    return res.status(400).json({
      ok: false,
      message: error.message,
    });
  }
});

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
    cellPath: z.string().trim().min(1),
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
    });

    console.log(`[api/com-test] ok=true cellPath=${parsed.data.cellPath}`);
    return res.json({
      ok: true,
      writtenValue: result.writtenValue,
      saved: result.saved === true,
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

app.get('/api/order/schema', (req, res) => {
  return res.json({
    ok: true,
    commitExample: buildOrderCommitExample(),
    schema: buildOrderSchemaInfo(),
    uiModel: {
      defaults: {
        eilig: false,
        matrixTyp: 'Boden',
        projektname: '',
        ansprechpartner: '',
        email: '',
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

app.post('/api/order/draft', (req, res) => {
  const order = parseOrderOrRespond(req, res);
  if (!order) {
    return;
  }

  return res.json({
    ...buildOrderPreview(order),
    operation: 'draft',
  });
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

  const commitWarnings = [];
  const absoluteExcelPath = resolveExcelPath(config.excelPath);
  const enrichedOrderResult = await applyPaketKeyTextsToOrder(order, absoluteExcelPath);
  const orderForWrite = enrichedOrderResult.order;
  commitWarnings.push(...enrichedOrderResult.warnings);

  const backup = ensureBackupBeforeCommit({
    config,
    excelPath: config.excelPath,
    rootDir: __dirname,
  });
  const preview = buildOrderPreview(orderForWrite);

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
    });
  } catch (error) {
    return res.status(400).json({
      ok: false,
      message: `Writer fehlgeschlagen (${usedBackend}): ${error.message}`,
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

  console.log(`[commit] writer=${usedBackend} order=${orderNo || 'n/a'} rows=${endRowRange || 'n/a'}`);

  logCommit({
    timestamp: new Date().toISOString(),
    mode: config.mode,
    writerBackend: usedBackend,
    kunde: orderForWrite.kunde,
    projekt: orderForWrite.projekt,
    projektnummer: orderForWrite.projektnummer,
    probeCount: orderForWrite.proben.length,
    backup,
    warnings: commitWarnings,
    orderNumberPreview: preview.orderNumberPreview,
    termin: preview.termin,
    writerResult: writeResult,
  });

  return res.json({
    ok: true,
    writer: usedBackend,
    saved,
    orderNo,
    auftragsnummer: orderNo,
    sampleNos,
    ersteProbennr,
    letzteProbennr,
    endRowRange,
    warnings: commitWarnings,
    ...preview,
    operation: 'commit',
    backup,
    writerBackend: usedBackend,
    writerResult: writeResult,
  });
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
});
