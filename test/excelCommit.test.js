const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const os = require('os');
const ExcelJS = require('exceljs');
const { appendOrderBlockToYearSheet } = require('../src/excelCommit');
const { buildTodayPrefix } = require('../src/sheetState');

test('appendOrderBlockToYearSheet appends header and probes to year sheet end', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');
  const now = new Date('2026-02-13T10:00:00Z');
  const prefix = buildTodayPrefix(now);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');
  sheet.getCell('A1').value = `${prefix}01`;
  sheet.getCell('A2').value = 26203;
  await workbook.xlsx.writeFile(excelPath);

  const result = await appendOrderBlockToYearSheet({
    config: { yearSheetName: '' },
    rootDir: tmpDir,
    excelPath,
    now,
    termin: '2026-02-17',
    packages: [{ id: 'pkg1', text: 'PAKET-TEXT' }],
    order: {
      kunde: 'Kunde A',
      projekt: 'Projekt X',
      projektnummer: 'P-123',
      eilig: true,
      probenahmedatum: '2026-02-13',
      probenEingangDatum: '2026-02-13',
      pbTyp: 'PB',
      auftragsnotiz: 'Notiz',
      kopfBemerkung: 'Hinweis Kopf',
      ansprechpartner: 'Max',
      email: 'max@example.com',
      kuerzel: 'AB',
      erfasstKuerzel: 'AB',
      proben: [
        {
          probenbezeichnung: 'Probe 1',
          material: 'Sand',
          containers: {
            mode: 'perSample',
            items: ['K:1L', 'K:1L', 'K:250mL', 'G:1L'],
          },
          matrixTyp: 'Boden',
          gewicht: 1.2,
          tiefeOderVolumen: '12 cm',
          geruchAuffaelligkeit: 'neutral',
          bemerkung: 'trocken',
          packageId: 'pkg1',
        },
        { probenbezeichnung: 'Probe 2', matrixTyp: 'Boden', gewicht: 2.5, parameterTextPreview: 'FREITEXT' },
      ],
    },
  });

  assert.equal(result.appendRow, 4);
  assert.equal(result.orderNo, `${prefix}02`);
  assert.equal(result.auftragsnummer, `${prefix}02`);
  assert.equal(result.saved, true);
  assert.equal(result.writer, 'exceljs');
  assert.equal(result.startLabNo, 26204);
  assert.deepEqual(result.sampleNos, [26204, 26205]);
  assert.equal(result.ersteProbennr, 26204);
  assert.equal(result.letzteProbennr, 26205);
  assert.equal(result.endRowRange, 'A4:J6');

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');

  assert.equal(checkSheet.getCell('A4').value, `${prefix}02`);
  assert.equal(checkSheet.getCell('B4').value, 'y');
  assert.equal(checkSheet.getCell('C4').value, 'y');
  assert.equal(checkSheet.getCell('D4').value, '');
  assert.equal(checkSheet.getCell('E4').value, 'PB');
  assert.equal(checkSheet.getCell('I4').value, 'Kunde A\nMax\nProjekt Nr: P-123\nProjekt: Projekt X\nProbenahme: 13.02.2026');
  assert.equal(
    checkSheet.getCell('J4').value,
    'AB EILIG Termin: Di 17.02.2026\nHinweis Kopf',
  );

  assert.equal(checkSheet.getCell('A5').value, 26204);
  assert.equal(checkSheet.getCell('D5').value, 'PAKET-TEXT');
  assert.equal(checkSheet.getCell('F5').value, 'Probe 1');
  assert.equal(checkSheet.getCell('G5').value, '12 cm');
  assert.equal(checkSheet.getCell('H5').value, 'Sand, Kunststoff (2x 1L; 250mL) Glas (1L)');
  assert.equal(checkSheet.getCell('J5').value, 'Gewicht: 1.2 kg; Geruch: neutral; trocken');
  assert.equal(checkSheet.getCell('J5').alignment.wrapText, true);
  assert.equal(checkSheet.getCell('A6').value, 26205);
  assert.equal(checkSheet.getCell('D6').value, 'FREITEXT');
  assert.equal(checkSheet.getCell('F6').value, 'Probe 2');
  assert.equal(checkSheet.getCell('G6').value, '');
  assert.equal(checkSheet.getCell('J6').value, 'Gewicht: 2.5 kg');

  assert.equal(checkSheet.getCell('A7').value, null);
});

test('appendOrderBlockToYearSheet renders probeJ only from present fields', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');

  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now: new Date('2026-02-13T10:00:00Z'),
    termin: null,
    order: {
      kunde: 'Kunde A',
      probenEingangDatum: '2026-02-13',
      proben: [
        { probenbezeichnung: 'Probe leer', gewicht: '', geruch: '', bemerkung: '' },
        { probenbezeichnung: 'Probe gewicht', gewicht: 2, geruch: null, bemerkung: undefined },
        { probenbezeichnung: 'Probe mix', gewicht: 2, geruch: 'muffig', bemerkung: '' },
        { probenbezeichnung: 'Probe bemerkung', gewicht: null, geruch: undefined, bemerkung: 'wenig Material' },
      ],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(checkSheet.getCell('J3').value, '');
  assert.equal(checkSheet.getCell('J4').value, 'Gewicht: 2 kg');
  assert.equal(checkSheet.getCell('J5').value, 'Gewicht: 2 kg; Geruch: muffig');
  assert.equal(checkSheet.getCell('J6').value, 'wenig Material');
});

test('appendOrderBlockToYearSheet writes transport line in header I when probentransport is set', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');

  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now: new Date('2026-02-13T10:00:00Z'),
    termin: null,
    order: {
      kunde: 'Kunde A',
      probentransport: 'AG',
      probenEingangDatum: '2026-02-13',
      proben: [{ probenbezeichnung: 'Probe 1' }],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(
    checkSheet.getCell('I2').value,
    'Kunde A\nTransport: AG',
  );
});

test('header I uses only Probenahmedatum when Eingang is today', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');
  const now = new Date('2026-02-13T10:00:00Z');
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now,
    termin: null,
    order: {
      kunde: 'Kunde A',
      probenahmedatum: '2026-02-12',
      probenEingangDatum: '2026-02-13',
      proben: [{ probenbezeichnung: 'Probe 1' }],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(checkSheet.getCell('I2').value, 'Kunde A\nProbenahme: 12.02.2026');
});

test('header I skips both date lines when Probenahmedatum is empty and Eingang is today', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');
  const now = new Date('2026-02-13T10:00:00Z');
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now,
    termin: null,
    order: {
      kunde: 'Kunde A',
      probenEingangDatum: '2026-02-13',
      proben: [{ probenbezeichnung: 'Probe 1' }],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(checkSheet.getCell('I2').value, 'Kunde A');
});

test('header I adds Eingangsdatum when Eingang differs from today', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');
  const now = new Date('2026-02-13T10:00:00Z');
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now,
    termin: null,
    order: {
      kunde: 'Kunde A',
      probenahmedatum: '2026-02-10',
      probenEingangDatum: '2026-02-12',
      proben: [{ probenbezeichnung: 'Probe 1' }],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(
    checkSheet.getCell('I2').value,
    'Kunde A\nProbenahme: 10.02.2026\nEingangsdatum: 12.02.2026',
  );
});

test('appendOrderBlockToYearSheet parses suffix sample numbers and writes clean new numbers', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');
  const now = new Date('2026-02-13T10:00:00Z');
  const prefix = buildTodayPrefix(now);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');
  sheet.getCell('A1').value = `${prefix}01A`;
  sheet.getCell('A2').value = `${prefix}02-1`;
  sheet.getCell('A3').value = '26203A';
  sheet.getCell('A4').value = '26203-1';
  sheet.getCell('A5').value = '26203 B';
  await workbook.xlsx.writeFile(excelPath);

  const result = await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now,
    termin: '2026-02-19',
    order: {
      kunde: 'Kunde A',
      projekt: 'Projekt X',
      projektnummer: 'P-123',
      eilig: false,
      probenEingangDatum: '2026-02-13',
      proben: [{ probenbezeichnung: 'Probe X', matrixTyp: 'Boden', gewicht: 1.2 }],
    },
  });

  assert.equal(result.orderNo, `${prefix}03`);
  assert.equal(result.startLabNo, 26204);

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');

  assert.equal(checkSheet.getCell('A7').value, `${prefix}03`);
  assert.equal(checkSheet.getCell('A8').value, 26204);
});

test('appendOrderBlockToYearSheet writes header containers and only material in probes when sameContainersForAll is true', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');

  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now: new Date('2026-02-13T10:00:00Z'),
    termin: '2026-02-19',
    order: {
      kunde: 'Kunde A',
      projekt: 'Projekt X',
      projektnummer: 'P-123',
      eilig: false,
      probenEingangDatum: '2026-02-13',
      sameContainersForAll: true,
      headerContainers: {
        mode: 'perOrder',
        items: ['K:1L', 'K:1L', 'G:500mL'],
      },
      proben: [
        { probenbezeichnung: 'Probe 1', material: 'Boden' },
        { probenbezeichnung: 'Probe 2', material: '' },
      ],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(checkSheet.getCell('H2').value, 'Kunststoff (2x 1L) Glas (500mL)');
  assert.equal(checkSheet.getCell('H3').value, 'Boden');
  assert.equal(checkSheet.getCell('H4').value, '');
});

test('appendOrderBlockToYearSheet parses lab numbers with suffix when finding max', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');
  const now = new Date('2026-02-14T10:00:00Z');
  const prefix = buildTodayPrefix(now);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');
  sheet.getCell('A1').value = `${prefix}01`;
  sheet.getCell('A2').value = '10038A';
  await workbook.xlsx.writeFile(excelPath);

  const result = await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now,
    termin: '2026-02-19',
    order: {
      kunde: 'Kunde A',
      projekt: 'Projekt X',
      projektnummer: 'P-123',
      eilig: false,
      probenEingangDatum: '2026-02-14',
      proben: [{ probenbezeichnung: 'Probe X', matrixTyp: 'Boden' }],
    },
  });

  assert.equal(result.startLabNo, 10039);
});

test('appendOrderBlockToYearSheet does not map weight into column G', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');

  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now: new Date('2026-02-13T10:00:00Z'),
    termin: '2026-02-19',
    order: {
      kunde: 'Kunde A',
      projekt: 'Projekt X',
      projektnummer: 'P-123',
      eilig: false,
      probenEingangDatum: '2026-02-13',
      proben: [{ probenbezeichnung: 'Probe X', matrixTyp: 'Boden', gewicht: 1.2 }],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(checkSheet.getCell('G3').value, '');
  assert.equal(checkSheet.getCell('J3').value, 'Gewicht: 1.2 kg');
});

test('appendOrderBlockToYearSheet writes hint when sampleNotArrived is true', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');

  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now: new Date('2026-02-13T10:00:00Z'),
    termin: null,
    order: {
      kunde: 'Kunde A',
      projekt: 'Projekt X',
      projektnummer: 'P-123',
      eilig: false,
      probeNochNichtDa: true,
      probenEingangDatum: undefined,
      proben: [{ probenbezeichnung: 'Probe 1', matrixTyp: 'Boden', gewicht: 1.2 }],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(checkSheet.getCell('J2').value, '');
});

test('appendOrderBlockToYearSheet keeps header J as single termin line when kopfBemerkung is empty', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');

  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now: new Date('2026-02-13T10:00:00Z'),
    termin: '2026-02-19',
    order: {
      kunde: 'Kunde A',
      projekt: 'Projekt X',
      projektnummer: 'P-123',
      eilig: false,
      kopfBemerkung: '',
      probenEingangDatum: '2026-02-13',
      proben: [{ probenbezeichnung: 'Probe 1', matrixTyp: 'Boden' }],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(checkSheet.getCell('J2').value, 'Termin: Do 19.02.2026');
});

test('appendOrderBlockToYearSheet writes kopfBemerkung only to header J and keeps header D empty', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');

  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now: new Date('2026-02-13T10:00:00Z'),
    termin: '2026-02-19',
    order: {
      kunde: 'Kunde A',
      projekt: 'Projekt X',
      projektnummer: 'P-123',
      eilig: false,
      kopfBemerkung: 'PB an Hans',
      auftragsnotiz: 'PB an Hans',
      probenEingangDatum: '2026-02-13',
      proben: [{ probenbezeichnung: 'Probe 1', matrixTyp: 'Boden' }],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(checkSheet.getCell('D2').value, '');
  assert.equal(checkSheet.getCell('J2').value, 'Termin: Do 19.02.2026\nPB an Hans');
});

test('appendOrderBlockToYearSheet writes adresseBlock in header J when excelWriteAddressBlock is true', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');

  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026', excelWriteAddressBlock: true },
    rootDir: tmpDir,
    excelPath,
    now: new Date('2026-02-13T10:00:00Z'),
    termin: '2026-02-19',
    order: {
      kunde: 'Kunde A',
      projektName: 'Projekt X',
      projektnummer: 'P-123',
      eilig: false,
      probenEingangDatum: '2026-02-13',
      adresseBlock: 'Kunde A GmbH\nzH Max Muster\nMusterstrasse 1\n12345 MUSTERSTADT',
      kopfBemerkung: 'PB an Hans',
      proben: [{ probenbezeichnung: 'Probe 1', matrixTyp: 'Boden' }],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(
    checkSheet.getCell('J2').value,
    'Termin: Do 19.02.2026\nPB an Hans\nKunde A GmbH\nzH Max Muster\nMusterstrasse 1\n12345 MUSTERSTADT',
  );
});

test('appendOrderBlockToYearSheet omits adresseBlock in header J when excelWriteAddressBlock is false', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');

  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026', excelWriteAddressBlock: false },
    rootDir: tmpDir,
    excelPath,
    now: new Date('2026-02-13T10:00:00Z'),
    termin: '2026-02-19',
    order: {
      kunde: 'Kunde A',
      projektName: 'Projekt X',
      projektnummer: 'P-123',
      eilig: false,
      probenEingangDatum: '2026-02-13',
      adresseBlock: 'Kunde A GmbH\nzH Max Muster\nMusterstrasse 1\n12345 MUSTERSTADT',
      kopfBemerkung: 'PB an Hans',
      proben: [{ probenbezeichnung: 'Probe 1', matrixTyp: 'Boden' }],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(
    checkSheet.getCell('J2').value,
    'Termin: Do 19.02.2026\nPB an Hans',
  );
});

test('appendOrderBlockToYearSheet sets wrapText and auto row height for new probe rows', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');

  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('2026');
  await workbook.xlsx.writeFile(excelPath);

  await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now: new Date('2026-02-13T10:00:00Z'),
    termin: '2026-02-19',
    order: {
      kunde: 'Kunde A',
      projekt: 'Projekt X',
      projektnummer: 'P-123',
      eilig: false,
      probenEingangDatum: '2026-02-13',
      proben: [
        {
          probenbezeichnung: 'Probe 1',
          matrixTyp: 'Boden',
          gewicht: 1.2,
          parameterTextPreview: 'Zeile 1\nZeile 2\nZeile 3',
        },
      ],
    },
  });

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  const probeRow = checkSheet.getRow(3);
  const dCell = checkSheet.getCell('D3');

  assert.equal(dCell.value, 'Zeile 1\nZeile 2\nZeile 3');
  assert.equal(dCell.alignment.wrapText, true);
  assert.equal(probeRow.height, 50);
});

test('appendOrderBlockToYearSheet starts lab numbers at 10000 when no previous sample number exists', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'annahme-commit-'));
  const excelPath = path.join(tmpDir, 'lab.xlsx');
  const now = new Date('2026-02-13T10:00:00Z');
  const prefix = buildTodayPrefix(now);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('2026');
  sheet.getCell('A1').value = `${prefix}01`;
  await workbook.xlsx.writeFile(excelPath);

  const result = await appendOrderBlockToYearSheet({
    config: { yearSheetName: '2026' },
    rootDir: tmpDir,
    excelPath,
    now,
    termin: null,
    order: {
      kunde: 'Kunde A',
      projekt: 'Projekt X',
      projektnummer: 'P-123',
      eilig: false,
      probenEingangDatum: '2026-02-13',
      proben: [{ probenbezeichnung: 'Probe 1', matrixTyp: 'Boden', gewicht: 1.2 }],
    },
  });

  assert.equal(result.startLabNo, 10000);
  assert.deepEqual(result.sampleNos, [10000]);

  const check = new ExcelJS.Workbook();
  await check.xlsx.readFile(excelPath);
  const checkSheet = check.getWorksheet('2026');
  assert.equal(checkSheet.getCell(`A${result.appendRow}`).value, `${prefix}02`);
  assert.equal(checkSheet.getCell(`A${result.appendRow + 1}`).value, 10000);
});
