initThemeToggle('theme-toggle');

const searchInput = document.getElementById('sp-search');
const addButton = document.getElementById('sp-add');
const editGroupsButton = document.getElementById('edit-groups-btn');
const saveButton = document.getElementById('sp-save');
const reloadButton = document.getElementById('sp-reload');
const warningEl = document.getElementById('sp-warning');
const statusEl = document.getElementById('sp-status');
const listEl = document.getElementById('sp-list');
const groupsModal = document.getElementById('groups-modal');
const groupsListEl = document.getElementById('groups-list');
const groupsAddButton = document.getElementById('groups-add-btn');
const groupsCancelButton = document.getElementById('groups-cancel');
const groupsApplyButton = document.getElementById('groups-apply');

const STANDARD_MEDIA = ['FS', 'H2O', '2e', '10e'];
const STANDARD_LABS = ['EMD', 'HB'];
const SINGLE_PARAM_CATALOG_UPDATED_AT_KEY = 'singleParamCatalogUpdatedAtV1';
const DEFAULT_GROUPS = [
  { key: 'AN', label: 'AN' },
  { key: 'SM', label: 'SM' },
  { key: 'Organik', label: 'Organik' },
];

const state = {
  catalog: { version: 1, parameters: [], groups: [] },
  rows: [],
  visibleIds: [],
  filter: '',
  groupDraftRows: [],
};

let rowIdCounter = 0;

function nextRowId(seed = '') {
  rowIdCounter += 1;
  return `sp_${Date.now()}_${rowIdCounter}_${String(seed || '').replace(/\s+/g, '_')}`;
}

function setStatus(message, isError = false) {
  statusEl.textContent = message || '';
  statusEl.classList.toggle('hidden', !message);
  statusEl.classList.toggle('inline-error', Boolean(message && isError));
  statusEl.classList.toggle('inline-ok', Boolean(message && !isError));
}

function setWarning(message) {
  warningEl.textContent = message || '';
  warningEl.classList.toggle('hidden', !message);
}

function normalizeGroups(rawGroups) {
  const source = Array.isArray(rawGroups) ? rawGroups : [];
  const normalized = source
    .map((g) => ({
      key: String(g?.key || '').trim(),
      label: String(g?.label || '').trim(),
    }))
    .filter((g) => g.key)
    .map((g) => ({ ...g, label: g.label || g.key }));
  if (normalized.length > 0) {
    return normalized;
  }
  return DEFAULT_GROUPS.map((g) => ({ ...g }));
}

function getCatalogGroups() {
  state.catalog.groups = normalizeGroups(state.catalog.groups);
  return state.catalog.groups;
}

function rowWarnings(row) {
  const warnings = [];
  if (!String(row.key || '').trim()) warnings.push('Kuerzel fehlt');
  const hasLab = Boolean(row.labEMD || row.labHB || String(row.otherLabName || '').trim());
  if (!hasLab) warnings.push('Kein Labor gesetzt');
  const hasMedia = Boolean(row.mFS || row.mH2O || row.m2e || row.m10e || String(row.otherMediumName || '').trim());
  if (!hasMedia) warnings.push('Kein Medium gesetzt');
  return warnings;
}

function renderWarningsBanner() {
  const hasWarnings = state.rows.some((row) => rowWarnings(row).length > 0);
  setWarning(hasWarnings ? 'Eintraege mit Warnung pruefen.' : '');
}

function findRowById(id) {
  return state.rows.find((row) => row.__id === id) || null;
}

function normalizeRowFromCatalog(param = {}, index = 0) {
  const key = String(param.key || '').trim();
  const allowedLabs = Array.isArray(param.allowedLabs) ? param.allowedLabs.map((x) => String(x || '').trim()).filter(Boolean) : [];
  const allowedMedia = Array.isArray(param.allowedMedia) ? param.allowedMedia.map((x) => String(x || '').trim()).filter(Boolean) : [];
  const firstOtherLab = allowedLabs.find((lab) => !STANDARD_LABS.includes(lab)) || String(param.otherLabName || '').trim();
  const firstOtherMedium = allowedMedia.find((m) => !STANDARD_MEDIA.includes(m)) || String(param.otherMediumName || '').trim();
  const groupRaw = String(param.functionGroup || '').trim();
  return {
    __id: String(param.__id || '').trim() || nextRowId(`${key || 'row'}_${index}`),
    labelLong: String(param.labelLong || param.label || '').trim(),
    key,
    functionGroup: groupRaw || '',
    labEMD: allowedLabs.includes('EMD'),
    labHB: allowedLabs.includes('HB'),
    otherLabName: firstOtherLab || '',
    mFS: allowedMedia.includes('FS'),
    mH2O: allowedMedia.includes('H2O'),
    m2e: allowedMedia.includes('2e'),
    m10e: allowedMedia.includes('10e'),
    otherMediumName: firstOtherMedium || '',
    eluatePrefixE: param.eluatePrefixE === true,
    allowVorOrt: param.allowVorOrt === true,
    pvFlag: param.pvFlag === true,
  };
}

function normalizeGroupRow(group = {}, index = 0) {
  const key = String(group?.key || '').trim();
  const label = String(group?.label || '').trim();
  const name = label || key;
  return {
    __id: String(group?.__id || '').trim() || nextRowId(`${name || key || 'group'}_${index}`),
    name: String(name || '').trim(),
  };
}

function deriveGroupKey(name) {
  return String(name || '').trim();
}

function updateVisibleIds() {
  const query = String(state.filter || '').trim().toLowerCase();
  state.visibleIds = state.rows
    .filter((row) => {
      const hay = `${row.labelLong || ''} ${row.key || ''}`.toLowerCase();
      return !query || hay.includes(query);
    })
    .map((row) => row.__id);
}

function createRowCard(row) {
  const warnings = rowWarnings(row);
  const card = document.createElement('div');
  card.className = 'sp-row';
  card.dataset.id = row.__id;
  card.style.display = 'grid';
  card.style.gridTemplateColumns = '2.0fr 0.8fr 0.9fr 0.8fr 0.7fr 1.3fr 2.0fr 0.4fr 0.7fr';
  card.style.gap = '10px';
  card.style.alignItems = 'center';
  card.style.padding = '8px 10px';
  card.style.borderBottom = '1px solid rgba(255,255,255,0.06)';

  const nameWrap = document.createElement('div');
  nameWrap.innerHTML = `<input type="text" data-field="labelLong" placeholder="z. B. pH-Wert" style="height:32px;" value="${(row.labelLong || '').replace(/"/g, '&quot;')}" />`;

  const keyWrap = document.createElement('div');
  keyWrap.innerHTML = `<input type="text" data-field="key" placeholder="pH" style="height:32px; max-width:110px;" value="${(row.key || '').replace(/"/g, '&quot;')}" />`;

  const functionWrap = document.createElement('div');
  const functionSelect = document.createElement('select');
  functionSelect.style.height = '32px';
  functionSelect.style.maxWidth = '140px';
  functionSelect.dataset.field = 'functionGroup';
  const functionOptions = [''].concat(getCatalogGroups().map((group) => String(group.key || '').trim()).filter(Boolean));
  if (row.functionGroup && !functionOptions.includes(row.functionGroup)) {
    functionOptions.push(row.functionGroup);
  }
  functionOptions.forEach((value) => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value || '-';
    functionSelect.appendChild(option);
  });
  functionSelect.value = row.functionGroup || '';
  functionWrap.appendChild(functionSelect);

  const eluatePrefixWrap = document.createElement('div');
  eluatePrefixWrap.style.display = 'flex';
  eluatePrefixWrap.style.justifyContent = 'center';
  eluatePrefixWrap.innerHTML = `<input type="checkbox" data-field="eluatePrefixE" ${row.eluatePrefixE ? 'checked' : ''} />`;

  const allowVorOrtWrap = document.createElement('div');
  allowVorOrtWrap.style.display = 'flex';
  allowVorOrtWrap.style.justifyContent = 'center';
  allowVorOrtWrap.innerHTML = `<input type="checkbox" data-field="allowVorOrt" ${row.allowVorOrt ? 'checked' : ''} />`;

  const deleteWrap = document.createElement('div');
  deleteWrap.style.display = 'flex';
  deleteWrap.style.justifyContent = 'center';
  const deleteButton = document.createElement('button');
  deleteButton.type = 'button';
  deleteButton.className = 'secondary';
  deleteButton.dataset.action = 'delete';
  deleteButton.textContent = 'Loeschen';
  deleteButton.style.height = '30px';
  deleteButton.style.padding = '0 8px';
  deleteWrap.appendChild(deleteButton);

  const labsWrap = document.createElement('div');
  const labsRow = document.createElement('div');
  labsRow.style.display = 'flex';
  labsRow.style.alignItems = 'center';
  labsRow.style.gap = '6px';
  labsRow.style.flexWrap = 'nowrap';
  labsRow.style.fontSize = '12px';
  labsRow.innerHTML = `
    <label style="font-size:12px;"><input type="checkbox" data-field="labEMD" ${row.labEMD ? 'checked' : ''} /> EMD</label>
    <label style="font-size:12px;"><input type="checkbox" data-field="labHB" ${row.labHB ? 'checked' : ''} /> HB</label>
    <input type="text" data-field="otherLabName" placeholder="NLGA" style="height:30px; max-width:96px;" value="${(row.otherLabName || '').replace(/"/g, '&quot;')}" />
  `;
  labsWrap.appendChild(labsRow);

  const mediaWrap = document.createElement('div');
  const mediaRow = document.createElement('div');
  mediaRow.style.display = 'flex';
  mediaRow.style.alignItems = 'center';
  mediaRow.style.gap = '6px';
  mediaRow.style.flexWrap = 'wrap';
  mediaRow.style.fontSize = '12px';
  mediaRow.innerHTML = `
    <label style="font-size:12px;"><input type="checkbox" data-field="mFS" ${row.mFS ? 'checked' : ''} /> FS</label>
    <label style="font-size:12px;"><input type="checkbox" data-field="mH2O" ${row.mH2O ? 'checked' : ''} /> H2O</label>
    <label style="font-size:12px;"><input type="checkbox" data-field="m2e" ${row.m2e ? 'checked' : ''} /> 2e</label>
    <label style="font-size:12px;"><input type="checkbox" data-field="m10e" ${row.m10e ? 'checked' : ''} /> 10e</label>
    <input type="text" data-field="otherMediumName" placeholder="S4" style="height:30px; max-width:80px;" value="${(row.otherMediumName || '').replace(/"/g, '&quot;')}" />
  `;
  mediaWrap.appendChild(mediaRow);

  const pvWrap = document.createElement('div');
  pvWrap.style.display = 'flex';
  pvWrap.style.justifyContent = 'center';
  pvWrap.innerHTML = `<input type="checkbox" data-field="pvFlag" ${row.pvFlag ? 'checked' : ''} />`;
  const warn = document.createElement('small');
  warn.className = 'inline-error';
  warn.dataset.role = 'row-warning';
  warn.textContent = warnings.join(' | ');
  warn.classList.toggle('hidden', warnings.length === 0);
  pvWrap.appendChild(warn);

  card.appendChild(nameWrap);
  card.appendChild(keyWrap);
  card.appendChild(functionWrap);
  card.appendChild(eluatePrefixWrap);
  card.appendChild(allowVorOrtWrap);
  card.appendChild(labsWrap);
  card.appendChild(mediaWrap);
  card.appendChild(pvWrap);
  card.appendChild(deleteWrap);
  return card;
}

function createGroupDraftRow(group, index) {
  const row = document.createElement('div');
  row.dataset.groupIndex = String(index);
  row.style.display = 'grid';
  row.style.gridTemplateColumns = '1fr auto';
  row.style.gap = '8px';
  row.style.alignItems = 'center';
  row.style.padding = '6px 0';
  row.style.borderBottom = '1px solid rgba(255,255,255,0.06)';

  const nameWrap = document.createElement('div');
  nameWrap.style.display = 'flex';
  nameWrap.style.flexDirection = 'column';
  nameWrap.style.gap = '4px';
  const groupNameValue = String(group.name || '');
  const derivedKey = deriveGroupKey(groupNameValue);
  nameWrap.innerHTML = `<input type="text" data-group-field="groupName" placeholder="Gruppenname (z.B. AN)" style="height:30px;" value="${groupNameValue.replace(/"/g, '&quot;')}" />`;
  const derivedKeyHint = document.createElement('small');
  derivedKeyHint.className = 'meta';
  derivedKeyHint.textContent = derivedKey ? `Key: ${derivedKey}` : 'Key: -';
  nameWrap.appendChild(derivedKeyHint);
  const deleteWrap = document.createElement('div');
  const deleteButton = document.createElement('button');
  deleteButton.type = 'button';
  deleteButton.className = 'secondary';
  deleteButton.dataset.action = 'delete-group';
  deleteButton.textContent = 'Loeschen';
  deleteButton.style.height = '30px';
  deleteButton.style.padding = '0 8px';
  deleteWrap.appendChild(deleteButton);

  row.appendChild(nameWrap);
  row.appendChild(deleteWrap);
  return row;
}

function renderGroupsModalList() {
  groupsListEl.innerHTML = '';
  if (state.groupDraftRows.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'meta';
    empty.textContent = 'Keine Gruppen';
    groupsListEl.appendChild(empty);
    return;
  }
  const fragment = document.createDocumentFragment();
  state.groupDraftRows.forEach((group, index) => {
    fragment.appendChild(createGroupDraftRow(group, index));
  });
  groupsListEl.appendChild(fragment);
}

function openGroupsModal() {
  state.groupDraftRows = getCatalogGroups().map((group, index) => normalizeGroupRow(group, index));
  renderGroupsModalList();
  groupsModal.classList.remove('hidden');
}

function closeGroupsModal() {
  groupsModal.classList.add('hidden');
  state.groupDraftRows = [];
}

function addGroupDraftRow() {
  state.groupDraftRows.push(normalizeGroupRow({ name: '' }, state.groupDraftRows.length));
  renderGroupsModalList();
}

function updateGroupDraftFieldByElement(target) {
  const field = String(target?.dataset?.groupField || '').trim();
  if (!field) return;
  const row = target.closest('[data-group-index]');
  if (!row) return;
  const index = Number.parseInt(row.dataset.groupIndex, 10);
  if (!Number.isInteger(index) || index < 0) return;
  if (!state.groupDraftRows[index]) return;
  if (field === 'groupName') {
    state.groupDraftRows[index].name = target.value;
  }
}

function applyGroupsModal() {
  const cleanedGroups = [];
  const seenKeys = new Set();
  for (const draft of state.groupDraftRows) {
    const name = String(draft?.name || '').trim();
    const key = deriveGroupKey(name);
    if (!name) {
      showToast('error', 'Gruppenname darf nicht leer sein.');
      return;
    }
    if (seenKeys.has(key)) {
      showToast('error', `Gruppen-Key doppelt: ${key}`);
      return;
    }
    seenKeys.add(key);
    cleanedGroups.push({ key, label: name });
  }

  state.catalog.groups = cleanedGroups;
  const allowedGroupKeys = new Set(cleanedGroups.map((group) => group.key));
  state.rows = state.rows.map((row) => {
    const group = String(row.functionGroup || '').trim();
    if (!group || allowedGroupKeys.has(group)) return row;
    return { ...row, functionGroup: null };
  });
  closeGroupsModal();
  renderList();
}

function renderList() {
  listEl.innerHTML = '';
  const fragment = document.createDocumentFragment();
  state.visibleIds.forEach((id) => {
    const row = findRowById(id);
    if (!row) return;
    fragment.appendChild(createRowCard(row));
  });
  listEl.appendChild(fragment);
  renderWarningsBanner();
}

function updateRowWarningInDom(rowId) {
  const row = findRowById(rowId);
  const card = listEl.querySelector(`[data-id="${rowId}"]`);
  if (!row || !card) return;
  const warn = card.querySelector('[data-role="row-warning"]');
  if (!warn) return;
  const warnings = rowWarnings(row);
  warn.textContent = warnings.join(' | ');
  warn.classList.toggle('hidden', warnings.length === 0);
}

function updateRowFieldByElement(target) {
  const field = String(target?.dataset?.field || '').trim();
  if (!field) return;
  const card = target.closest('[data-id]');
  if (!card) return;
  const row = findRowById(card.dataset.id);
  if (!row) return;
  row[field] = target.type === 'checkbox' ? target.checked : target.value;
  updateRowWarningInDom(row.__id);
  renderWarningsBanner();
}

function addParameter() {
  const row = {
    __id: nextRowId('new'),
    labelLong: '',
    key: '',
    functionGroup: '',
    labEMD: true,
    labHB: false,
    otherLabName: '',
    mFS: false,
    mH2O: false,
    m2e: false,
    m10e: false,
    otherMediumName: '',
    eluatePrefixE: false,
    allowVorOrt: false,
    pvFlag: false,
  };
  state.rows.push(row);
  updateVisibleIds();
  renderList();
  const firstField = listEl.querySelector(`[data-id="${row.__id}"] [data-field="labelLong"]`);
  if (firstField) firstField.focus();
}

function normalizeCatalogForSave() {
  const groups = getCatalogGroups()
    .map((group) => ({
      key: String(group?.key || '').trim(),
      label: String(group?.label || '').trim(),
    }))
    .filter((group) => group.key)
    .map((group) => ({ ...group, label: group.label || group.key }));
  const allowedGroupKeys = new Set(groups.map((group) => group.key));

  const parameters = state.rows.map((row) => {
    const key = String(row.key || '').trim();
    const labelLong = String(row.labelLong || '').trim();
    const label = labelLong || key;
    const otherLabName = String(row.otherLabName || '').trim();
    const otherMediumName = String(row.otherMediumName || '').trim();
    const allowedLabs = [];
    if (row.labEMD) allowedLabs.push('EMD');
    if (row.labHB) allowedLabs.push('HB');
    if (otherLabName) allowedLabs.push(otherLabName);
    const allowedMedia = [];
    if (row.mFS) allowedMedia.push('FS');
    if (row.mH2O) allowedMedia.push('H2O');
    if (row.m2e) allowedMedia.push('2e');
    if (row.m10e) allowedMedia.push('10e');
    if (otherMediumName) allowedMedia.push(otherMediumName);
    const requiresPv = [];
    if (allowedMedia.includes('2e')) requiresPv.push('2e');
    if (allowedMedia.includes('10e')) requiresPv.push('10e');
    const groupRaw = String(row.functionGroup || '').trim();
    const functionGroup = groupRaw && allowedGroupKeys.has(groupRaw) ? groupRaw : null;
    return {
      key,
      label,
      labelLong,
      allowedLabs,
      defaultLab: allowedLabs.includes('EMD') ? 'EMD' : (allowedLabs.includes('HB') ? 'HB' : (otherLabName || allowedLabs[0] || 'EMD')),
      allowedMedia,
      otherLabName: otherLabName || undefined,
      otherMediumName: otherMediumName || undefined,
      eluatePrefixE: row.eluatePrefixE === true,
      allowVorOrt: row.allowVorOrt === true,
      pvFlag: row.pvFlag === true,
      requiresPv,
      functionGroup,
    };
  });

  return {
    version: Number(state.catalog.version || 1),
    groups,
    parameters,
  };
}

async function loadCatalog() {
  setStatus('', false);
  try {
    const catalog = await fetchSingleParamCatalog();
    state.catalog = catalog && typeof catalog === 'object' ? catalog : { version: 1, parameters: [] };
    state.catalog.groups = normalizeGroups(state.catalog.groups);
    state.rows = (Array.isArray(state.catalog.parameters) ? state.catalog.parameters : [])
      .map((p, index) => normalizeRowFromCatalog(p, index));
    updateVisibleIds();
    renderList();
    setStatus(`Katalog geladen (${state.rows.length} Parameter).`, false);
  } catch (error) {
    setStatus(error.message, true);
    showToast('error', error.message);
  }
}

async function saveCatalog() {
  setStatus('', false);
  const hasEmptyKey = state.rows.some((row) => !String(row.key || '').trim());
  if (hasEmptyKey) {
    const message = 'Kuerzel darf nicht leer sein.';
    setStatus(message, true);
    showToast('error', message);
    return;
  }
  try {
    const payload = normalizeCatalogForSave();
    const saved = await saveSingleParamCatalog(payload);
    state.catalog = saved && typeof saved === 'object' ? saved : payload;
    state.catalog.groups = normalizeGroups(state.catalog.groups);
    const updatedAt = String(state.catalog.updatedAt || '').trim();
    if (updatedAt) {
      localStorage.setItem(SINGLE_PARAM_CATALOG_UPDATED_AT_KEY, updatedAt);
    }
    state.rows = (Array.isArray(state.catalog.parameters) ? state.catalog.parameters : [])
      .map((p, index) => normalizeRowFromCatalog(p, index));
    updateVisibleIds();
    renderList();
    setStatus(`Gespeichert (${state.rows.length} Parameter).`, false);
    showToast('success', 'Einzelparameter gespeichert');
  } catch (error) {
    setStatus(error.message, true);
    showToast('error', error.message);
  }
}

listEl.addEventListener('input', (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) && !(target instanceof HTMLTextAreaElement) && !(target instanceof HTMLSelectElement)) return;
  updateRowFieldByElement(target);
});

listEl.addEventListener('change', (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) && !(target instanceof HTMLTextAreaElement) && !(target instanceof HTMLSelectElement)) return;
  updateRowFieldByElement(target);
});

listEl.addEventListener('click', (event) => {
  const button = event.target.closest('button[data-action="delete"]');
  if (!button) return;
  const card = button.closest('[data-id]');
  if (!card) return;
  state.rows = state.rows.filter((row) => row.__id !== card.dataset.id);
  updateVisibleIds();
  renderList();
});

groupsListEl.addEventListener('input', (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  updateGroupDraftFieldByElement(target);
  if (String(target.dataset.groupField || '') === 'groupName') {
    const row = target.closest('[data-group-index]');
    const hint = row ? row.querySelector('small.meta') : null;
    if (hint) {
      const key = deriveGroupKey(target.value);
      hint.textContent = key ? `Key: ${key}` : 'Key: -';
    }
  }
});

groupsListEl.addEventListener('click', (event) => {
  const button = event.target.closest('button[data-action="delete-group"]');
  if (!button) return;
  const row = button.closest('[data-group-index]');
  if (!row) return;
  const index = Number.parseInt(row.dataset.groupIndex, 10);
  if (!Number.isInteger(index) || index < 0) return;
  state.groupDraftRows.splice(index, 1);
  renderGroupsModalList();
});

searchInput.addEventListener('input', () => {
  state.filter = searchInput.value;
  updateVisibleIds();
  renderList();
});

addButton.addEventListener('click', addParameter);
editGroupsButton.addEventListener('click', openGroupsModal);
groupsAddButton.addEventListener('click', addGroupDraftRow);
groupsCancelButton.addEventListener('click', closeGroupsModal);
groupsApplyButton.addEventListener('click', applyGroupsModal);
groupsModal.addEventListener('click', (event) => {
  if (event.target === groupsModal) closeGroupsModal();
});

saveButton.addEventListener('click', saveCatalog);
reloadButton.addEventListener('click', loadCatalog);

loadCatalog();
