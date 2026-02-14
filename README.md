# Annahme Tool

Express-Projekt mit statischer Annahme-Maske in `public/`.

## Voraussetzungen

- Node.js 18+
- npm

## Installation

```bash
npm install
```

## Konfiguration

- `config.example.json` enthaelt alle verfuegbaren Konfig-Werte.
- Optional `config.json` im Projektroot anlegen.
- Wenn `config.json` fehlt, werden Dev-Defaults genutzt.
- Konfiguration wird beim Start mit Zod validiert.

Beispiel:

```json
{
  "port": 3000,
  "mode": "single",
  "writerBackend": "com",
  "excelPath": "./data/lab.xlsx",
  "yearSheetName": "",
  "writerHost": "http://localhost:3000",
  "writerToken": "change-me",
  "backupEnabled": true,
  "backupPolicy": "daily",
  "backupIntervalMinutes": 60,
  "backupRetentionDays": 14,
  "backupZip": false,
  "uiShowPackagePreview": true,
  "uiDefaultEilig": "ja"
}
```

Backup-Optionen:

- `backupEnabled`: Backups vor Commit aktiv/inaktiv
- `backupPolicy`: `daily` oder `interval`
- `backupIntervalMinutes`: Intervall in Minuten (nur bei `interval`)
- `backupRetentionDays`: Loeschung alter Backups
- `backupZip`: Dateiendung `.zip` statt `.xlsx`

## Start

Entwicklungsmodus mit Nodemon:

```bash
npm run dev
```

Produktivstart:

```bash
npm start
```

Server laeuft auf `http://localhost:3000`.

## Betriebsmodi

- `single`:
  - Draft und Commit ohne Token erlaubt
  - Commit Button sichtbar
- `writer`:
  - Commit nur fuer UI-Request mit passendem Writer-Token erlaubt
  - Commit Button sichtbar
  - Token-Login ueber `settings.html`, Token/Status in `sessionStorage`
- `client`:
  - Commit API deaktiviert
  - Commit Button ausgeblendet
  - nur Draft senden

## API

- `GET /api/config`
  - liefert `config` (ohne `writerToken`), `mode`, `excelPath`, `commitAllowed`, `canWriteConfig`
- `POST /api/config`
  - speichert erlaubte Config-Felder in `config.json` und laedt Runtime-Config ohne Neustart
  - Ebene 1 ohne Admin-Key: `excelPath`, Backup-Felder, UI-Felder
  - Ebene 2 nur mit Header `x-admin-key` passend zu `ANNAHME_ADMIN_KEY`: `mode`, `writerHost`, `writerToken`, `writerBackend`
  - `ANNAHME_ADMIN_KEY` muss als Umgebungsvariable gesetzt sein, sonst sind Ebene-2-Aenderungen gesperrt
- `GET /api/config/validate?excelPath=...`
  - prueft Excel-Pfad, Sheet `Vorlagen` mit mind. 1 Paket, Jahresblatt `yearSheetName` (oder aktuelles Jahr falls leer)
  - Antwort: `warnings`, `errors`
- `POST /api/writer/login`
  - Body: `{ "token": "..." }`
  - validiert gegen `writerToken` aus Config
- `POST /api/com-test`
  - Body: `{ "cellPath": "2026!Z1" }`
  - schreibt via COM in die laufende Excel-Instanz den Wert `COM_OK_<timestamp>`
  - Antwort: `{ "ok": true, "writtenValue": "COM_OK_<timestamp>", "saved": true }`
- `POST /api/order/draft`
  - validiert Auftrag und liefert Vorschau
- `POST /api/order/commit`
  - mode-abhaengig (siehe oben)
  - Writer-Auswahl:
    - `mode=single` + Windows-Pfad (`C:\...`) => COM Writer als Standard
    - `writerBackend: "comExceljs"` oder `writerBackend: "exceljs"` => ExcelJS erzwingen
    - Linux/WSL ohne COM => ExcelJS Fallback
    - `com`: Windows COM Automation ueber `scripts/writer.ps1`
    - `exceljs`: direkter XLSX-Schreibpfad in Node
  - erstellt vor Commit rotierende Backups (nicht pro Auftrag)
  - fuehrt Cleanup alter Backups aus
  - Antwort enthaelt u. a.: `{ ok, writer, saved, orderNo, sampleNos, ersteProbennr, letzteProbennr, endRowRange }`

Antwort-Vorschau enthaelt u. a.:

- `termin` (oder `null`)
- `orderNumberPreview` (oder `null`)
- `labNumberPreview`
- `backup` (Status des Backup-Schritts)

Backup-Dateinamen:

- `Annahme_backup_YYYYMMDD_HHMMSS.xlsx`
- oder `Annahme_backup_YYYYMMDD_HHMMSS.zip` (bei `backupZip: true`)

Commit-Logging:

- jeder Commit wird als JSON-Zeile in `logs/commit-log.jsonl` geschrieben
- kein Vollbackup pro Auftrag, nur gem. Backup-Policy

## Dev zuhause

- lokal nur `config.json` Werte anpassen (z. B. `mode`, `excelPath`, `writerToken`)
- Code bleibt gleich

## Office setup spaeter

- ebenfalls nur `config.json` Werte anpassen
- spaeter kann Excel-Auslesen hinterlegt werden, ohne UI/API-Struktur zu aendern

## Tests

```bash
npm test
```
