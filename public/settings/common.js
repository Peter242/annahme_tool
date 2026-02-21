const THEME_KEY = 'cua-theme';

function initThemeToggle(buttonId = 'theme-toggle') {
  const button = document.getElementById(buttonId);
  if (!button) return;

  function applyTheme(theme) {
    const resolved = theme === 'light' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', resolved);
    localStorage.setItem(THEME_KEY, resolved);
    button.textContent = resolved === 'dark' ? 'Light' : 'Dark';
  }

  applyTheme(localStorage.getItem(THEME_KEY) || 'dark');
  button.addEventListener('click', () => {
    const current = document.documentElement.getAttribute('data-theme') || 'dark';
    applyTheme(current === 'dark' ? 'light' : 'dark');
  });
}

function ensureToastContainer() {
  let container = document.getElementById('toast-container');
  if (container) return container;
  container = document.createElement('div');
  container.id = 'toast-container';
  container.className = 'toast-wrap';
  document.body.appendChild(container);
  return container;
}

function showToast(type, message) {
  const container = ensureToastContainer();
  const el = document.createElement('div');
  el.className = `toast ${type || 'info'}`;
  el.textContent = message;
  container.appendChild(el);
  setTimeout(() => el.remove(), 3200);
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);
  let data = {};
  try {
    data = await response.json();
  } catch (_error) {
    data = {};
  }
  return { response, data };
}

function flattenErrorMessage(data, fallback) {
  if (data && typeof data.message === 'string' && data.message.trim()) {
    return data.message.trim();
  }
  return fallback;
}

async function loadConfig() {
  const { response, data } = await fetchJson('/api/config');
  if (!response.ok || data.ok !== true || !data.config) {
    throw new Error(flattenErrorMessage(data, 'Einstellungen konnten nicht geladen werden'));
  }
  return data.config;
}

async function saveConfig(payload, adminKey = '') {
  const headers = { 'Content-Type': 'application/json' };
  if (adminKey) headers['x-admin-key'] = adminKey;
  const { response, data } = await fetchJson('/api/config', {
    method: 'POST',
    headers,
    body: JSON.stringify(payload),
  });
  if (!response.ok || data.ok !== true) {
    throw new Error(flattenErrorMessage(data, 'Speichern fehlgeschlagen'));
  }
  return data.config;
}

async function validateExcelPath(excelPath) {
  const query = new URLSearchParams({ excelPath }).toString();
  const { response, data } = await fetchJson(`/api/config/validate?${query}`);
  if (!response.ok || data.ok !== true) {
    const firstError = Array.isArray(data.errors) && data.errors.length > 0
      ? data.errors[0]
      : flattenErrorMessage(data, 'Pfadprüfung fehlgeschlagen');
    throw new Error(firstError);
  }
  return data;
}

async function validateBackupDir(backupDir) {
  const query = new URLSearchParams({ dir: backupDir }).toString();
  const { response, data } = await fetchJson(`/api/backups/validate?${query}`);
  if (!response.ok) {
    throw new Error(flattenErrorMessage(data, 'Backup-Pfadpruefung fehlgeschlagen'));
  }
  return data;
}

async function createManualBackup(force = false) {
  const { response, data } = await fetchJson('/api/backups/create', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ force: force === true }),
  });
  if (!response.ok || data.ok !== true || data.created !== true) {
    throw new Error(flattenErrorMessage(data, 'Manuelles Backup fehlgeschlagen'));
  }
  return data;
}

async function pickBackupDirNative() {
  const { response, data } = await fetchJson('/api/system/pick-backup-dir');
  if (!response.ok || data.ok !== true) {
    throw new Error(flattenErrorMessage(data, 'Ordnerauswahl fehlgeschlagen'));
  }
  if (data.canceled === true) {
    return { canceled: true, selectedPath: null };
  }
  return { canceled: false, selectedPath: data.selectedPath };
}

async function fetchSingleParamCatalog() {
  const { response, data } = await fetchJson('/api/single-parameter-catalog');
  if (!response.ok || data.ok !== true || !data.catalog) {
    throw new Error(flattenErrorMessage(data, 'Einzelparameter-Katalog konnte nicht geladen werden'));
  }
  return data.catalog;
}

async function saveSingleParamCatalog(catalog) {
  const { response, data } = await fetchJson('/api/single-parameter-catalog', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ catalog }),
  });
  if (!response.ok || data.ok !== true || !data.catalog) {
    throw new Error(flattenErrorMessage(data, 'Einzelparameter-Katalog konnte nicht gespeichert werden'));
  }
  return data.catalog;
}

async function resetCache() {
  const { response, data } = await fetchJson('/api/state/reset', { method: 'POST' });
  if (!response.ok || data.ok !== true) {
    throw new Error(flattenErrorMessage(data, 'Cache konnte nicht zurückgesetzt werden'));
  }
  return data;
}
