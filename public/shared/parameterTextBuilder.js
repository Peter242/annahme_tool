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
  const parts = [];
  if (pv.ts === true) {
    parts.push('TS');
  }
  if (Array.isArray(pv.eluate)) {
    pv.eluate.forEach((entry) => {
      if (!isPlainObject(entry)) return;
      const site = entry.site === 'E' || entry.site === 'B' ? entry.site : null;
      const ratio = entry.ratio === '2e' || entry.ratio === '10e' ? entry.ratio : null;
      if (!site || !ratio) return;
      parts.push(`${site}(${ratio})`);
    });
  }
  if (parts.length === 0) return '';
  return `PV: ${parts.join(', ')}`;
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
    mediumParts.push(`${medium}: ${renderedItems}`);
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

window.buildParameterTextFromSelection = function buildParameterTextFromSelection(selection) {
  if (!isPlainObject(selection)) return '';

  const lines = [];
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
};
