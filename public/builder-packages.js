initThemeToggle('theme-toggle');

const searchInput = document.getElementById('bp-search');
const reloadButton = document.getElementById('bp-reload');
const statusEl = document.getElementById('bp-status');
const listEl = document.getElementById('bp-list');

let packages = [];

function setStatus(message = '', isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle('hidden', !message);
  statusEl.classList.toggle('inline-ok', Boolean(message && !isError));
  statusEl.classList.toggle('inline-error', Boolean(message && isError));
}

function getFilteredPackages() {
  const query = String(searchInput.value || '').trim().toLocaleLowerCase('de-DE');
  const sorted = [...packages].sort((a, b) => {
    const aName = String(a?.name || '').trim();
    const bName = String(b?.name || '').trim();
    return aName.localeCompare(bName, 'de', { sensitivity: 'base' });
  });
  if (!query) return sorted;
  return sorted.filter((entry) => {
    const haystack = [
      String(entry?.name || ''),
      String(entry?.id || ''),
    ].join(' ').toLocaleLowerCase('de-DE');
    return haystack.includes(query);
  });
}

function renderList() {
  listEl.innerHTML = '';
  const rows = getFilteredPackages();
  if (rows.length < 1) {
    const empty = document.createElement('div');
    empty.className = 'meta';
    empty.textContent = 'Keine Builder Pakete vorhanden.';
    listEl.appendChild(empty);
    return;
  }

  const fragment = document.createDocumentFragment();
  rows.forEach((entry) => {
    const row = document.createElement('div');
    row.style.display = 'grid';
    row.style.gridTemplateColumns = '1fr auto auto';
    row.style.gap = '8px';
    row.style.alignItems = 'center';
    row.style.padding = '8px 10px';
    row.style.borderBottom = '1px solid rgba(255,255,255,0.06)';

    const nameWrap = document.createElement('div');
    const name = document.createElement('div');
    name.textContent = String(entry?.name || '').trim() || '(ohne Name)';
    name.style.fontWeight = '600';
    const meta = document.createElement('div');
    meta.className = 'meta';
    meta.style.fontSize = '12px';
    meta.textContent = String(entry?.id || '').trim();
    nameWrap.appendChild(name);
    nameWrap.appendChild(meta);

    const editButton = document.createElement('button');
    editButton.type = 'button';
    editButton.className = 'secondary';
    editButton.textContent = 'Bearbeiten';
    editButton.addEventListener('click', () => {
      const id = encodeURIComponent(String(entry?.id || '').trim());
      window.location.href = `/?editBuilderPackage=${id}`;
    });

    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'secondary';
    deleteButton.textContent = 'Löschen';
    deleteButton.addEventListener('click', async () => {
      const nameText = String(entry?.name || '').trim() || String(entry?.id || '').trim();
      if (!window.confirm(`Builder Paket "${nameText}" löschen?`)) return;
      try {
        const { response, data } = await fetchJson(`/api/builder-packages/${encodeURIComponent(String(entry?.id || '').trim())}`, {
          method: 'DELETE',
        });
        if (!response.ok || data.ok !== true) {
          throw new Error(flattenErrorMessage(data, 'Builder-Paket konnte nicht gelöscht werden'));
        }
        showToast('success', 'Builder-Paket gelöscht');
        await loadBuilderPackages();
      } catch (error) {
        showToast('error', error.message);
      }
    });

    row.appendChild(nameWrap);
    row.appendChild(editButton);
    row.appendChild(deleteButton);
    fragment.appendChild(row);
  });
  listEl.appendChild(fragment);
}

async function loadBuilderPackages() {
  try {
    const { response, data } = await fetchJson('/api/builder-packages', { cache: 'no-store' });
    if (!response.ok || data.ok !== true || !Array.isArray(data.packages)) {
      throw new Error(flattenErrorMessage(data, 'Builder-Pakete konnten nicht geladen werden'));
    }
    packages = data.packages;
    renderList();
    setStatus(`${packages.length} Builder Pakete geladen.`);
  } catch (error) {
    setStatus(error.message, true);
    showToast('error', error.message);
  }
}

searchInput.addEventListener('input', renderList);
reloadButton.addEventListener('click', loadBuilderPackages);

loadBuilderPackages();
