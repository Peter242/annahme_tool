function splitTopLevel(input, delimiterChar) {
  const text = String(input || '');
  const out = [];
  let current = '';
  let depth = 0;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (ch === '(') depth += 1;
    if (ch === ')' && depth > 0) depth -= 1;
    if (ch === delimiterChar && depth === 0) {
      out.push(current.trim());
      current = '';
      continue;
    }
    current += ch;
  }
  out.push(current.trim());
  return out.filter(Boolean);
}

function parseLabeledLine(line) {
  const raw = String(line || '');
  const idx = raw.indexOf(':');
  if (idx < 0) return null;
  const label = raw.slice(0, idx).trim();
  const value = raw.slice(idx + 1).trim();
  if (!label) return null;
  return { label, value };
}

function normalizeLabel(label) {
  return String(label || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function isPvLabel(labelNorm) {
  return labelNorm === 'pv';
}

function isVorOrtLabel(labelNorm) {
  return labelNorm === 'vor ort' || labelNorm === 'vor-ort';
}

function canonicalMedium(rawMedium, labelNorm) {
  const source = String(rawMedium || '').trim();
  const raw = String(source || '').trim().toLowerCase();
  const key = raw
    .replace(/\./g, '')
    .replace(/-/g, '')
    .replace(/\s+/g, '');
  if (key === 'fs') return 'FS';
  if (key === 'h2o') return 'H2O';
  if (labelNorm === 'hb') {
    if ((key.includes('2:1') && key.includes('eluat')) || key === '2e') return '2e';
    if ((key.includes('10:1') && key.includes('eluat')) || key === '10e') return '10e';
    if (key === '1eluat' || key === '1:1eluat') return '10e';
    if (key.startsWith('1') && key.includes('eluat')) return '10e';
    if (key.includes('eluat') && key.includes('10')) return '10e';
    if (key.includes('eluat') && key.includes('2')) return '2e';
  } else {
    if (key === '2e') return '2e';
    if (key === '10e') return '10e';
  }
  console.warn(`[legacy-merge] unrecognized-medium:${labelNorm}:${source}`);
  return source;
}

function mediumDisplayName(canonical, labelNorm, original) {
  if (labelNorm === 'hb') {
    if (canonical === '2e') return '2:1-Eluat';
    if (canonical === '10e') return '10:1-Eluat';
  }
  if (canonical === 'FS') return 'FS';
  if (canonical === 'H2O') return 'H2O';
  if (canonical === '2e') return '2e';
  if (canonical === '10e') return '10e';
  return String(original || canonical || '').trim();
}

function parseMediumBlocks(value, labelNorm) {
  const rawBlocks = splitTopLevel(value, ';');
  const blocks = [];
  for (const rawBlock of rawBlocks) {
    const idx = rawBlock.indexOf(':');
    if (idx < 0) {
      return { ok: false, reason: `medium_block_missing_colon:${rawBlock}` };
    }
    const medium = rawBlock.slice(0, idx).trim();
    const content = rawBlock.slice(idx + 1).trim();
    if (!medium) {
      return { ok: false, reason: 'medium_label_empty' };
    }
    const canonical = canonicalMedium(medium, labelNorm);
    const tokens = splitTopLevel(content, ',');
    blocks.push({
      originalMedium: medium,
      canonical,
      tokens,
    });
  }
  return { ok: true, blocks };
}

function dedupeKeepOrder(tokens) {
  const out = [];
  const seen = new Set();
  for (const token of tokens) {
    const value = String(token || '').trim();
    if (!value || seen.has(value)) continue;
    seen.add(value);
    out.push(value);
  }
  return out;
}

function parseGroupToken(token, groupName) {
  const value = String(token || '').trim();
  const plainPrefix = `${groupName}(`;
  const eluPrefix = `e${groupName}(`;
  const isEluat = value.startsWith(eluPrefix);
  const isPlain = value.startsWith(plainPrefix);
  if (!isEluat && !isPlain) return null;
  const start = value.indexOf('(');
  const end = value.lastIndexOf(')');
  if (start < 0 || end <= start) return null;
  const members = value
    .slice(start + 1, end)
    .split(',')
    .map((x) => String(x || '').trim())
    .filter(Boolean);
  return {
    isEluat,
    members,
  };
}

function mergeSpecialEluatGroupTokens(baseTokensInput, addonTokensInput, groupName) {
  const baseTokens = Array.isArray(baseTokensInput) ? [...baseTokensInput] : [];
  const addonTokens = Array.isArray(addonTokensInput) ? [...addonTokensInput] : [];
  const eluPrefix = `e${groupName}(`;
  const plainPrefix = `${groupName}(`;
  const eluIndex = baseTokens.findIndex((token) => String(token || '').trim().startsWith(eluPrefix));
  const plainIndex = baseTokens.findIndex((token) => String(token || '').trim().startsWith(plainPrefix));
  const targetIndex = eluIndex >= 0 ? eluIndex : plainIndex;
  if (targetIndex < 0) {
    return { baseTokens, addonTokens };
  }

  const parsedBase = parseGroupToken(baseTokens[targetIndex], groupName);
  if (!parsedBase) {
    return { baseTokens, addonTokens };
  }

  const consumedAddonIndexes = [];
  const addonMembers = [];
  addonTokens.forEach((token, index) => {
    const parsed = parseGroupToken(token, groupName);
    if (!parsed) return;
    consumedAddonIndexes.push(index);
    addonMembers.push(...parsed.members);
  });
  if (addonMembers.length < 1) {
    return { baseTokens, addonTokens };
  }

  const mergedMembers = dedupeKeepOrder([...parsedBase.members, ...addonMembers]);
  const nextPrefix = parsedBase.isEluat ? `e${groupName}` : groupName;
  baseTokens[targetIndex] = `${nextPrefix}(${mergedMembers.join(', ')})`;
  const consumedSet = new Set(consumedAddonIndexes);
  const nextAddonTokens = addonTokens.filter((_, index) => !consumedSet.has(index));
  return { baseTokens, addonTokens: nextAddonTokens };
}

function mergePvLine(baseValue, addonValue) {
  const merged = dedupeKeepOrder([
    ...splitTopLevel(baseValue, ','),
    ...splitTopLevel(addonValue, ','),
  ]);
  return merged.join(', ');
}

function mergeVorOrtLine(baseValue, addonValue) {
  const merged = dedupeKeepOrder([
    ...String(baseValue || '').split(','),
    ...String(addonValue || '').split(','),
  ]);
  return merged.join(', ');
}

function mergeMediumLine(baseValue, addonValue, labelNorm) {
  const parsedBase = parseMediumBlocks(baseValue, labelNorm);
  if (!parsedBase.ok) return parsedBase;
  const parsedAddon = parseMediumBlocks(addonValue, labelNorm);
  if (!parsedAddon.ok) return parsedAddon;

  const byCanonical = new Map();
  const order = [];
  parsedBase.blocks.forEach((block) => {
    byCanonical.set(block.canonical, {
      originalMedium: block.originalMedium,
      tokens: [...block.tokens],
    });
    order.push(block.canonical);
  });
  parsedAddon.blocks.forEach((block) => {
    const existing = byCanonical.get(block.canonical);
    if (existing) {
      let baseTokens = [...existing.tokens];
      let addonTokens = [...block.tokens];
      if (labelNorm === 'hb' && (block.canonical === '2e' || block.canonical === '10e')) {
        const smMerged = mergeSpecialEluatGroupTokens(baseTokens, addonTokens, 'SM');
        baseTokens = smMerged.baseTokens;
        addonTokens = smMerged.addonTokens;
        const anMerged = mergeSpecialEluatGroupTokens(baseTokens, addonTokens, 'AN');
        baseTokens = anMerged.baseTokens;
        addonTokens = anMerged.addonTokens;
      }
      existing.tokens = dedupeKeepOrder([...baseTokens, ...addonTokens]);
      return;
    }
    byCanonical.set(block.canonical, {
      originalMedium: block.originalMedium,
      tokens: [...block.tokens],
    });
    order.push(block.canonical);
  });

  const blockTexts = order
    .map((canonical) => {
      const block = byCanonical.get(canonical);
      if (!block) return '';
      const display = mediumDisplayName(canonical, labelNorm, block.originalMedium);
      const mergedTokens = dedupeKeepOrder(block.tokens);
      return `${display}: ${mergedTokens.join(', ')}`.trim();
    })
    .filter(Boolean);
  return { ok: true, value: blockTexts.join('; ') };
}

function mergeLineValues(labelNorm, baseValue, addonValue) {
  if (isPvLabel(labelNorm)) return { ok: true, value: mergePvLine(baseValue, addonValue) };
  if (isVorOrtLabel(labelNorm)) return { ok: true, value: mergeVorOrtLine(baseValue, addonValue) };
  return mergeMediumLine(baseValue, addonValue, labelNorm);
}

function tryMergeLegacyPackage(baseText, addonText, _options = {}) {
  const base = String(baseText || '').trim();
  const addon = String(addonText || '').trim();
  if (!addon) {
    return { ok: true, mergedText: base, reason: null };
  }
  const baseLines = base ? base.split(/\r?\n/) : [];
  const addonLines = addon.split(/\r?\n/).map((x) => String(x || '').trim()).filter(Boolean);

  const parsedBase = baseLines.map((line, idx) => {
    const parsed = parseLabeledLine(line);
    if (!parsed) return { raw: line, index: idx, label: null, value: null, labelNorm: null };
    return {
      raw: line,
      index: idx,
      label: parsed.label,
      value: parsed.value,
      labelNorm: normalizeLabel(parsed.label),
    };
  });

  for (const addonLine of addonLines) {
    const parsedAddon = parseLabeledLine(addonLine);
    if (!parsedAddon) {
      console.warn(`[legacy-merge] fallback:addon_line_missing_colon:${addonLine}`);
      return { ok: false, mergedText: base, reason: `addon_line_missing_colon:${addonLine}` };
    }
    const addonLabelNorm = normalizeLabel(parsedAddon.label);
    const baseMatch = parsedBase.find((entry) => entry.labelNorm === addonLabelNorm);
    if (!baseMatch) {
      const appended = `${parsedAddon.label}: ${parsedAddon.value}`.trim();
      baseLines.push(appended);
      parsedBase.push({
        raw: appended,
        index: baseLines.length - 1,
        label: parsedAddon.label,
        value: parsedAddon.value,
        labelNorm: addonLabelNorm,
      });
      continue;
    }
    const merged = mergeLineValues(addonLabelNorm, baseMatch.value, parsedAddon.value);
    if (!merged.ok) {
      const reason = merged.reason || `merge_failed:${parsedAddon.label}`;
      console.warn(`[legacy-merge] fallback:${reason}`);
      return {
        ok: false,
        mergedText: base,
        reason,
      };
    }
    const newLine = `${baseMatch.label}: ${merged.value}`.trim();
    baseLines[baseMatch.index] = newLine;
    baseMatch.raw = newLine;
    baseMatch.value = merged.value;
  }

  return {
    ok: true,
    mergedText: baseLines.join('\n').trim(),
    reason: null,
  };
}

module.exports = {
  tryMergeLegacyPackage,
};
