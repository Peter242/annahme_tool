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
      probenEingangDatum: '2026-02-13',
      pbTyp: 'PB',
      auftragsnotiz: 'Notiz',
      ansprechpartner: 'Max',
      email: 'max@example.com',
      kuerzel: 'AB',
      erfasstKuerzel: 'AB',
      proben: [
        {
          probenbezeichnung: 'Probe 1',
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
  assert.equal(checkSheet.getCell('D4').value, 'Notiz');
  assert.equal(checkSheet.getCell('E4').value, 'PB');
  assert.equal(checkSheet.getCell('I4').value, 'Kunde A\nMax\nProjekt Nr: P-123\nProjekt: Projekt X\nProbenahme: 13.02.2026');
  assert.equal(
    checkSheet.getCell('J4').value,
    'AB EILIG Termin: Di 17.02.2026\nMail: max@example.com',
  );

  assert.equal(checkSheet.getCell('A5').value, 26204);
  assert.equal(checkSheet.getCell('D5').value, 'PAKET-TEXT');
  assert.equal(checkSheet.getCell('F5').value, 'Probe 1');
  assert.equal(checkSheet.getCell('G5').value, '12 cm');
  assert.equal(checkSheet.getCell('J5').value, 'Gewicht: 1.2 kg\nGeruch: neutral\nBemerkung: trocken');
  assert.equal(checkSheet.getCell('J5').alignment.wrapText, true);
  assert.equal(checkSheet.getCell('A6').value, 26205);
  assert.equal(checkSheet.getCell('D6').value, 'FREITEXT');
  assert.equal(checkSheet.getCell('F6').value, 'Probe 2');
  assert.equal(checkSheet.getCell('G6').value, '');
  assert.equal(checkSheet.getCell('J6').value, 'Gewicht: 2.5 kg\nGeruch: -\nBemerkung: -');

  assert.equal(checkSheet.getCell('A7').value, null);
});

test('appendOrderBlockToYearSheet scans suffix numbers and writes clean new numbers', async () => {
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

test('appendOrderBlockToYearSheet parses leading digits for lab numbers with suffix', async () => {
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
  assert.equal(checkSheet.getCell('J3').value, 'Gewicht: 1.2 kg\nGeruch: -\nBemerkung: -');
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
