const fs = require('fs');
const path = require('path');

const rootDir = path.join(__dirname, '..');
const sourcePath = path.join(rootDir, 'data', 'single_parameter_source.tsv');
const targetPath = path.join(rootDir, 'data', 'single_parameter_catalog.json');
const MEDIA_DEFAULT = ['FS', 'H2O', '2e', '10e'];
const REQUIRES_PV_DEFAULT = ['2e', '10e'];

function parseTsv(content) {
  const lines = String(content || '')
    .replace(/^\uFEFF/, '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line !== '');

  if (lines.length === 0) {
    throw new Error('TSV ist leer.');
  }

  const header = lines[0].split('\t').map((x) => x.trim());
  if (header.length < 3 || header[0] !== 'Zielzeile' || header[1] !== 'Zielparameter' || header[2] !== 'Funktion') {
    throw new Error('Ungueltiger TSV Header. Erwartet: Zielzeile<TAB>Zielparameter<TAB>Funktion');
  }

  const rows = [];
  for (let i = 1; i < lines.length; i += 1) {
    const raw = lines[i];
    const cols = raw.split('\t');
    if (cols.length < 2) {
      throw new Error(`TSV Parsefehler in Zeile ${i + 1}: mindestens 2 Spalten erwartet.`);
    }
    const zielzeile = String(cols[0] || '').trim();
    const zielparameter = String(cols[1] || '').trim();
    const funktion = String(cols[2] || '').trim();
    if (!zielzeile || !zielparameter) {
      throw new Error(`TSV Parsefehler in Zeile ${i + 1}: Zielzeile und Zielparameter sind erforderlich.`);
    }
    rows.push({ zielzeile, zielparameter, funktion });
  }
  return rows;
}

function buildCatalog(rows) {
  const grouped = new Map();

  rows.forEach((row) => {
    if (!grouped.has(row.zielparameter)) {
      grouped.set(row.zielparameter, {
        key: row.zielparameter,
        label: row.zielparameter,
        allowedLabs: [],
        functionGroup: null,
      });
    }
    const entry = grouped.get(row.zielparameter);
    if (!entry.allowedLabs.includes(row.zielzeile)) {
      entry.allowedLabs.push(row.zielzeile);
    }
    if (row.funktion === 'AN') {
      entry.functionGroup = 'AN';
    } else if (row.funktion === 'SM' && entry.functionGroup !== 'AN') {
      entry.functionGroup = 'SM';
    }
  });

  function sortLabs(labs) {
    const unique = Array.from(new Set(labs));
    const primary = [];
    if (unique.includes('EMD')) primary.push('EMD');
    if (unique.includes('HB')) primary.push('HB');
    const remaining = unique
      .filter((lab) => lab !== 'EMD' && lab !== 'HB')
      .sort((a, b) => a.localeCompare(b, 'de', { sensitivity: 'base' }));
    return [...primary, ...remaining];
  }

  const parameters = Array.from(grouped.values()).map((entry) => {
    const allowedLabs = sortLabs(entry.allowedLabs);
    const defaultLab = allowedLabs.includes('EMD')
      ? 'EMD'
      : allowedLabs[0];
    return {
      key: entry.key,
      label: entry.label,
      allowedLabs,
      defaultLab,
      allowedMedia: [...MEDIA_DEFAULT],
      requiresPv: [...REQUIRES_PV_DEFAULT],
      functionGroup: entry.functionGroup,
    };
  }).sort((a, b) => a.key.localeCompare(b.key, 'de', { sensitivity: 'base' }));

  return {
    version: 1,
    generatedAt: new Date().toISOString(),
    parameters,
  };
}

function main() {
  try {
    const source = fs.readFileSync(sourcePath, 'utf-8');
    const rows = parseTsv(source);
    const catalog = buildCatalog(rows);
    fs.writeFileSync(targetPath, `${JSON.stringify(catalog, null, 2)}\n`, 'utf-8');
    console.log(`single_parameter_catalog.json aktualisiert (${catalog.parameters.length} Parameter).`);
  } catch (error) {
    console.error(`Fehler beim Bauen des Single Parameter Katalogs: ${error.message}`);
    process.exit(1);
  }
}

main();
