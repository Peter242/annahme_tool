const QUICK_CONTAINER_DEFAULTS = Object.freeze({
  plastic: Object.freeze([
    '5L',
    '3L',
    '1L',
    '500mL',
    '250mL',
    '250mL + CaCO3',
    '250mL + NaOH',
    '100mL',
    '100mL + NaOH',
    '30mL',
    '30mL + HCl',
    '30mL + HNO3',
  ]),
  glass: Object.freeze([
    '1L',
    '1L + H2SO4',
    '500mL',
    '500mL + H2SO4',
    '250mL',
    '250mL Schliff',
    '250mL Duran',
    'HS',
    'HS + CuSO4',
    'HS + MeOH',
  ]),
});

function toTrimmedString(value) {
  return String(value || '').trim();
}

function normalizeQuickLabel(rawLabel) {
  return toTrimmedString(rawLabel).replace(/\s+/g, ' ').replace(/\s*\+\s*/g, ' + ');
}

function makeTokenIdFromLabel(label) {
  const normalized = normalizeQuickLabel(label);
  if (!normalized) {
    return '';
  }
  return normalized.replace(/\s*\+\s*/g, '+').replace(/\s+/g, '-');
}

function makeDisplayFromTokenId(tokenId) {
  return toTrimmedString(tokenId).replace(/-/g, ' ').replace(/\+/g, ' + ');
}

function parseQuickList(list, prefix) {
  const normalizedPrefix = prefix === 'G' ? 'G' : 'K';
  const seen = new Set();
  const options = [];
  for (const raw of Array.isArray(list) ? list : []) {
    const label = normalizeQuickLabel(raw);
    const id = makeTokenIdFromLabel(label);
    if (!id || seen.has(id)) {
      continue;
    }
    seen.add(id);
    options.push({
      prefix: normalizedPrefix,
      id,
      token: `${normalizedPrefix}:${id}`,
      label: makeDisplayFromTokenId(id),
    });
  }
  return options;
}

function normalizeQuickContainerConfig(config = {}) {
  const plasticRaw = Array.isArray(config.quickContainerPlastic)
    ? config.quickContainerPlastic
    : QUICK_CONTAINER_DEFAULTS.plastic;
  const glassRaw = Array.isArray(config.quickContainerGlass)
    ? config.quickContainerGlass
    : QUICK_CONTAINER_DEFAULTS.glass;
  const plastic = parseQuickList(plasticRaw, 'K').map((option) => option.label);
  const glass = parseQuickList(glassRaw, 'G').map((option) => option.label);
  return {
    plastic: plastic.length > 0 ? plastic : [...QUICK_CONTAINER_DEFAULTS.plastic],
    glass: glass.length > 0 ? glass : [...QUICK_CONTAINER_DEFAULTS.glass],
  };
}

function normalizeToken(raw, fallbackPrefix = null) {
  const text = toTrimmedString(raw);
  if (!text) {
    return '';
  }
  const match = text.match(/^([KG]):(.+)$/i);
  if (match) {
    const prefix = match[1].toUpperCase();
    const id = toTrimmedString(match[2]);
    return id ? `${prefix}:${id}` : '';
  }
  if (fallbackPrefix) {
    const prefix = fallbackPrefix === 'G' ? 'G' : 'K';
    const id = makeTokenIdFromLabel(text);
    return id ? `${prefix}:${id}` : '';
  }
  return '';
}

function normalizeContainerItems(items) {
  const normalized = [];
  for (const raw of Array.isArray(items) ? items : []) {
    const token = normalizeToken(raw);
    if (token) {
      normalized.push(token);
    }
  }
  return normalized;
}

function normalizeContainers(containers, options = {}) {
  const modeDefault = options.modeDefault || 'perSample';
  const source = containers && typeof containers === 'object' ? containers : {};
  const mode = source.mode === 'perOrder' ? 'perOrder' : (source.mode === 'perSample' ? 'perSample' : modeDefault);
  return {
    mode,
    items: normalizeContainerItems(source.items),
    history: [],
  };
}

function buildOptionMaps(config = {}) {
  const normalized = normalizeQuickContainerConfig(config);
  const plasticOptions = parseQuickList(normalized.plastic, 'K');
  const glassOptions = parseQuickList(normalized.glass, 'G');
  const tokenLabel = new Map();
  const order = new Map();
  plasticOptions.forEach((option, index) => {
    tokenLabel.set(option.token, option.label);
    order.set(option.token, index);
  });
  glassOptions.forEach((option, index) => {
    tokenLabel.set(option.token, option.label);
    order.set(option.token, index);
  });
  return { tokenLabel, order };
}

function renderContainers(items, options = {}) {
  const tokens = normalizeContainerItems(items);
  if (tokens.length === 0) {
    return '';
  }

  const { tokenLabel, order } = buildOptionMaps(options.config || {});
  const groups = {
    K: new Map(),
    G: new Map(),
  };
  const unknownOrder = {
    K: new Map(),
    G: new Map(),
  };
  let unknownIndex = 10000;

  for (const token of tokens) {
    const match = token.match(/^([KG]):(.+)$/);
    if (!match) {
      continue;
    }
    const prefix = match[1];
    const id = match[2];
    const current = groups[prefix].get(token) || 0;
    groups[prefix].set(token, current + 1);
    if (!order.has(token) && !unknownOrder[prefix].has(token)) {
      unknownOrder[prefix].set(token, unknownIndex);
      unknownIndex += 1;
    }
    if (!tokenLabel.has(token)) {
      tokenLabel.set(token, makeDisplayFromTokenId(id));
    }
  }

  const renderGroup = (prefix, title) => {
    const entries = Array.from(groups[prefix].entries());
    if (entries.length === 0) {
      return '';
    }
    entries.sort((a, b) => {
      const ao = order.has(a[0]) ? order.get(a[0]) : (unknownOrder[prefix].get(a[0]) || 99999);
      const bo = order.has(b[0]) ? order.get(b[0]) : (unknownOrder[prefix].get(b[0]) || 99999);
      return ao - bo;
    });
    const parts = entries.map(([token, qty]) => (qty > 1 ? `${qty}x ${tokenLabel.get(token)}` : tokenLabel.get(token)));
    return `${title} (${parts.join('; ')})`;
  };

  const blocks = [renderGroup('K', 'Kunststoff'), renderGroup('G', 'Glas')].filter(Boolean);
  return blocks.join(' ');
}

function renderContainersSummary(containers, options = {}) {
  if (Array.isArray(containers)) {
    return renderContainers(containers, options);
  }
  return renderContainers(containers?.items, options);
}

function renderColumnHFromProbe(probe, options = {}) {
  const material = toTrimmedString(probe?.material);
  if (options.onlyMaterial === true) {
    return material;
  }
  const summary = renderContainersSummary(probe?.containers, options);
  if (material && summary) {
    return `${material}, ${summary}`;
  }
  if (material) {
    return material;
  }
  if (summary) {
    return summary;
  }
  return '';
}

module.exports = {
  QUICK_CONTAINER_DEFAULTS,
  normalizeQuickContainerConfig,
  parseQuickList,
  normalizeToken,
  normalizeContainerItems,
  normalizeContainers,
  renderContainers,
  renderContainersSummary,
  renderColumnHFromProbe,
};
