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
  "excelPath": "./data/lab.xlsx",
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
  - Ebene 2 nur mit Header `x-admin-key` passend zu `ANNAHME_ADMIN_KEY`: `mode`, `writerHost`, `writerToken`
  - `ANNAHME_ADMIN_KEY` muss als Umgebungsvariable gesetzt sein, sonst sind Ebene-2-Aenderungen gesperrt
- `GET /api/config/validate?excelPath=...`
  - prueft Excel-Pfad, Sheet `Vorlagen` mit mind. 1 Paket, Jahresblatt des aktuellen Jahres
  - Antwort: `warnings`, `errors`
- `POST /api/writer/login`
  - Body: `{ "token": "..." }`
  - validiert gegen `writerToken` aus Config
- `POST /api/order/draft`
  - validiert Auftrag und liefert Vorschau
- `POST /api/order/commit`
  - mode-abhaengig (siehe oben)
  - erstellt vor Commit rotierende Backups (nicht pro Auftrag)
  - fuehrt Cleanup alter Backups aus

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
