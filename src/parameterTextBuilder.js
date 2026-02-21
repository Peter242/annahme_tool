const LAB_ORDER = ['EMD', 'HB'];
const MEDIUM_ORDER = ['FS', 'H2O', '2e', '10e'];

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function normalizeToken(value) {
  if (typeof value !== 'string') return '';
  return value.trim();
}

function renderPvLine(pv) {
  if (!isPlainObject(pv)) return '';
  const bySite = { E: { tokens: [], ratios: [] }, B: { tokens: [], ratios: [] } };
  if (isPlainObject(pv.itemsBySite)) {
    ['E', 'B'].forEach((site) => {
      const rawItems = Array.isArray(pv.itemsBySite[site]) ? pv.itemsBySite[site] : [];
      rawItems.forEach((raw) => {
        const token = normalizeToken(raw);
        if (token) bySite[site].tokens.push(token);
      });
    });
  }
  if (Array.isArray(pv.eluate)) {
    pv.eluate.forEach((entry) => {
      if (!isPlainObject(entry)) return;
      const site = entry.site === 'E' || entry.site === 'B' ? entry.site : null;
      const ratio = entry.ratio === '2e' || entry.ratio === '10e' ? entry.ratio : null;
      if (!site || !ratio) return;
      bySite[site].ratios.push(ratio);
    });
  }
  const ratioOrder = { '2e': 0, '10e': 1 };
  const siteParts = [];
  ['E', 'B'].forEach((site) => {
    const tokens = Array.from(new Set(bySite[site].tokens)).sort((a, b) => a.localeCompare(b, 'de'));
    const ratios = Array.from(new Set(bySite[site].ratios)).sort((a, b) => (ratioOrder[a] ?? 99) - (ratioOrder[b] ?? 99));
    const members = [...tokens, ...ratios];
    if (members.length > 0) {
      siteParts.push(`${site}(${members.join(', ')})`);
    }
  });
  if (siteParts.length > 0) {
    return `PV: ${siteParts.join(', ')}`;
  }
  if (pv.ts === true) {
    return 'PV: TS';
  }
  return '';
}

function renderVorOrtLine(vorOrt) {
  if (!Array.isArray(vorOrt)) return '';
  const items = vorOrt.map(normalizeToken).filter(Boolean);
  if (items.length === 0) return '';
  return `vor Ort: ${items.join(', ')}`;
}

function renderItemsWithAnBlock(items) {
  if (!Array.isArray(items)) return '';

  const anItems = [];
  const restItems = [];

  items.forEach((raw) => {
    const token = normalizeToken(raw);
    if (!token) return;
    if (token.startsWith('AN:')) {
      const value = token.slice(3).trim();
      if (value) {
        anItems.push(value);
      }
      return;
    }
    restItems.push(token);
  });

  const rendered = [];
  if (anItems.length > 0) {
    rendered.push(`AN(${anItems.join(', ')})`);
  }
  if (restItems.length > 0) {
    rendered.push(...restItems);
  }
  return rendered.join(', ');
}

function collectLabMedia(selection, labName) {
  const byMedium = new Map();
  MEDIUM_ORDER.forEach((m) => byMedium.set(m, []));

  const labs = Array.isArray(selection.labs) ? selection.labs : [];
  labs.forEach((labEntry) => {
    if (!isPlainObject(labEntry) || labEntry.lab !== labName) return;
    const media = Array.isArray(labEntry.media) ? labEntry.media : [];
    media.forEach((mediumEntry) => {
      if (!isPlainObject(mediumEntry)) return;
      const medium = mediumEntry.medium;
      if (!byMedium.has(medium)) return;
      byMedium.get(medium).push(mediumEntry.items);
    });
  });

  return byMedium;
}

function formatMediumLabelForLab(lab, mediumKey) {
  if (lab === 'HB') {
    if (mediumKey === '2e') return '2:1-Eluat';
    if (mediumKey === '10e') return '10:1-Eluat';
  }
  return mediumKey;
}

function renderLabLine(selection, labName) {
  const byMedium = collectLabMedia(selection, labName);
  const mediumParts = [];

  MEDIUM_ORDER.forEach((medium) => {
    const itemGroups = byMedium.get(medium) || [];
    const flatItems = [];
    itemGroups.forEach((group) => {
      if (!Array.isArray(group)) return;
      group.forEach((item) => flatItems.push(item));
    });
    const renderedItems = renderItemsWithAnBlock(flatItems);
    if (!renderedItems) return;
    const mediumLabel = formatMediumLabelForLab(labName, medium);
    mediumParts.push(`${mediumLabel}: ${renderedItems}`);
  });

  if (mediumParts.length === 0) return '';
  return `${labName}: ${mediumParts.join('; ')}`;
}

function renderExternLines(selection) {
  const extern = Array.isArray(selection.extern) ? selection.extern : [];
  return extern
    .map((entry) => {
      if (!isPlainObject(entry)) return '';
      const lab = normalizeToken(entry.lab);
      if (!lab) return '';
      if (!Array.isArray(entry.items)) return '';
      const items = entry.items.map(normalizeToken).filter(Boolean);
      if (items.length === 0) return '';
      return `${lab}: ${items.join(', ')}`;
    })
    .filter(Boolean);
}

function buildParameterTextFromSelection(selection) {
  if (!isPlainObject(selection)) return '';

  const lines = [];
  const vorOrtLine = renderVorOrtLine(selection.vorOrt);
  if (vorOrtLine) {
    lines.push(vorOrtLine);
  }
  const pvLine = renderPvLine(selection.pv);
  if (pvLine) {
    lines.push(pvLine);
  }

  LAB_ORDER.forEach((labName) => {
    const labLine = renderLabLine(selection, labName);
    if (labLine) {
      lines.push(labLine);
    }
  });

  lines.push(...renderExternLines(selection));
  return lines.join('\n').trim();
}

module.exports = {
  buildParameterTextFromSelection,
};
