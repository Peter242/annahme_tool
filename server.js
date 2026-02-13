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
const { calculateTermin } = require('./src/termin');

const configSchema = z
  .object({
    port: z.number().int().min(1).max(65535),
    mode: z.enum(['single', 'writer', 'client']),
    writerBackend: z.enum(['exceljs', 'com']),
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

function getConfig() {
  return runtimeConfig;
}

function getYearSheetName(config) {
  const configured = (config.yearSheetName || '').trim();
  return configured || String(new Date().getFullYear());
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
  writerBackend: z.enum(['exceljs', 'com']).optional(),
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
    parameterTextPreview: z.string().optional(),
    tiefeVolumen: z.union([z.string(), z.number()]).optional(),
    geruchAuffaelligkeit: z.string().trim().optional(),
    bemerkung: z.string().trim().optional(),
    materialGebinde: z.string().optional(),
    material: z.string().optional(),
    gebinde: z.string().optional(),
  })
  .superRefine((sample, ctx) => {
    if (sample.matrixTyp === 'Boden') {
      if (typeof sample.gewicht !== 'number') {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ['gewicht'],
          message: 'Gewicht ist fuer Matrix Typ Boden erforderlich',
        });
      }
      if (sample.volumen !== undefined) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ['volumen'],
          message: 'Volumen darf bei Boden nicht gesetzt sein',
        });
      }
    }

    if (sample.matrixTyp === 'Wasser' || sample.matrixTyp === 'Luft') {
      if (typeof sample.volumen !== 'number') {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ['volumen'],
          message: `Volumen ist fuer Matrix Typ ${sample.matrixTyp} erforderlich`,
        });
      }
      if (sample.gewicht !== undefined) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ['gewicht'],
          message: `Gewicht darf bei ${sample.matrixTyp} nicht gesetzt sein`,
        });
      }
    }
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
      const renderedTextD = packageTemplate ? packageTemplate.text : '';
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

  const backup = ensureBackupBeforeCommit({
    config,
    excelPath: config.excelPath,
    rootDir: __dirname,
  });
  const preview = buildOrderPreview(order);

  try {
    await writeOrderBlock({
      backend: config.writerBackend,
      config,
      rootDir: __dirname,
      excelPath: config.excelPath,
      order,
      termin: preview.termin,
    });
  } catch (error) {
    return res.status(400).json({
      ok: false,
      message: `Writer fehlgeschlagen (${config.writerBackend}): ${error.message}`,
    });
  }

  logCommit({
    timestamp: new Date().toISOString(),
    mode: config.mode,
    kunde: order.kunde,
    projekt: order.projekt,
    projektnummer: order.projektnummer,
    probeCount: order.proben.length,
    backup,
    orderNumberPreview: preview.orderNumberPreview,
    termin: preview.termin,
  });

  return res.json({
    ...preview,
    operation: 'commit',
    backup,
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
