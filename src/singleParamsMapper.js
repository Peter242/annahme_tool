const LAB_ORDER = ['EMD', 'HB'];
const MEDIUM_ORDER = ['FS', 'H2O', '2e', '10e'];
const DEFAULT_GROUPS = [
  { key: 'AN', label: 'AN', supportsEluateE: true },
  { key: 'SM', label: 'SM', supportsEluateE: true },
  { key: 'Organik', label: 'Organik', supportsEluateE: true },
];

function normalizeArrayStrings(value) {
  return Array.isArray(value)
    ? value.map((x) => String(x || '').trim()).filter(Boolean)
    : [];
}

function normalizeGroups(catalog) {
  const raw = Array.isArray(catalog?.groups) ? catalog.groups : [];
  const cleaned = raw
    .map((g) => ({
      key: String(g?.key || '').trim(),
      label: String(g?.label || '').trim(),
      supportsEluateE: g?.supportsEluateE === true,
    }))
    .filter((g) => g.key);
  return cleaned.length > 0
    ? cleaned.map((g) => ({ ...g, label: g.label || g.key }))
    : DEFAULT_GROUPS.map((g) => ({ ...g }));
}

function createMediumBucket(groupDefs) {
  const groupMembers = {};
  const groupEFlags = {};
  groupDefs.forEach((g) => {
    groupMembers[g.key] = [];
    groupEFlags[g.key] = false;
  });
  return {
    plainTokens: [],
    groupMembers,
    groupEFlags,
  };
}

function dedupeSorted(values) {
  return Array.from(new Set((Array.isArray(values) ? values : []).map((v) => String(v || '').trim()).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, 'de'));
}

function buildMediumItems(bucket, medium, groupDefs) {
  const isEluate = medium === '2e' || medium === '10e';
  const out = [];
  groupDefs.forEach((groupDef) => {
    const members = dedupeSorted(bucket.groupMembers[groupDef.key]);
    if (members.length === 0) return;
    let functionName = groupDef.label || groupDef.key;
    if (isEluate && groupDef.supportsEluateE === true && bucket.groupEFlags[groupDef.key]) {
      functionName = `e${functionName}`;
    }
    out.push(`${functionName}(${members.join(', ')})`);
  });
  out.push(...dedupeSorted(bucket.plainTokens));
  return out;
}

function mapTogglesToSelection({ catalog, toggles }) {
  const catalogParams = Array.isArray(catalog?.parameters) ? catalog.parameters : [];
  const byKey = new Map();
  catalogParams.forEach((p) => {
    const key = String(p?.key || '').trim();
    if (key) byKey.set(key, p);
  });

  const groupDefs = normalizeGroups(catalog);
  const groupByKey = new Map(groupDefs.map((g) => [g.key, g]));

  const sourceToggles = toggles && typeof toggles === 'object' ? toggles : {};
  const labMediumBuckets = {
    EMD: { FS: createMediumBucket(groupDefs), H2O: createMediumBucket(groupDefs), '2e': createMediumBucket(groupDefs), '10e': createMediumBucket(groupDefs) },
    HB: { FS: createMediumBucket(groupDefs), H2O: createMediumBucket(groupDefs), '2e': createMediumBucket(groupDefs), '10e': createMediumBucket(groupDefs) },
  };
  const pvSeen = new Set();
  const pvEluate = [];
  const pvItemsBySite = { E: [], B: [] };
  const vorOrtTokens = [];

  Object.entries(sourceToggles).forEach(([key, toggle]) => {
    if (!toggle || toggle.selected !== true) return;
    const catalogParam = byKey.get(key);
    if (!catalogParam) return;

    const label = String(catalogParam.label || key).trim() || key;
    const tokenBase = String(catalogParam.key || label || key).trim() || key;
    if (toggle.vorOrt === true) {
      vorOrtTokens.push(tokenBase);
      return;
    }
    const allowedLabs = normalizeArrayStrings(catalogParam.allowedLabs);
    const allowedMedia = normalizeArrayStrings(catalogParam.allowedMedia);
    if (allowedLabs.length === 0 || allowedMedia.length === 0) return;

    const requestedLab = String(toggle.lab || '').trim();
    let lab = allowedLabs.includes(requestedLab)
      ? requestedLab
      : String(catalogParam.defaultLab || '').trim();
    if (!allowedLabs.includes(lab)) {
      lab = allowedLabs[0];
    }
    if (!LAB_ORDER.includes(lab)) return;
    const site = lab === 'EMD' ? 'E' : (lab === 'HB' ? 'B' : null);
    if (catalogParam.pvFlag === true && site) {
      pvItemsBySite[site].push(tokenBase);
    }

    const selectedMediaMap = toggle.media && typeof toggle.media === 'object' ? toggle.media : {};
    const requiresPv = new Set(normalizeArrayStrings(catalogParam.requiresPv));
    const groupKeyRaw = String(catalogParam.functionGroup || '').trim();
    let groupDef = groupByKey.get(groupKeyRaw) || null;
    if (!groupDef && groupKeyRaw) {
      groupDef = { key: groupKeyRaw, label: groupKeyRaw, supportsEluateE: true };
      groupDefs.push(groupDef);
      groupByKey.set(groupKeyRaw, groupDef);
      LAB_ORDER.forEach((labName) => {
        MEDIUM_ORDER.forEach((mediumName) => {
          const bucket = labMediumBuckets[labName][mediumName];
          bucket.groupMembers[groupKeyRaw] = bucket.groupMembers[groupKeyRaw] || [];
          bucket.groupEFlags[groupKeyRaw] = bucket.groupEFlags[groupKeyRaw] === true;
        });
      });
    }

    MEDIUM_ORDER.forEach((medium) => {
      if (!allowedMedia.includes(medium)) return;
      if (selectedMediaMap[medium] !== true) return;

      const isEluate = medium === '2e' || medium === '10e';
      const wantsE = isEluate && catalogParam.eluatePrefixE === true;
      const bucket = labMediumBuckets[lab][medium];

      if (groupDef) {
        bucket.groupMembers[groupDef.key].push(tokenBase);
        if (wantsE) bucket.groupEFlags[groupDef.key] = true;
      } else {
        bucket.plainTokens.push(wantsE ? `e${tokenBase}` : tokenBase);
      }

      if (isEluate && requiresPv.has(medium)) {
        const site = lab === 'EMD' ? 'E' : 'B';
        const pvKey = `${site}:${medium}`;
        if (!pvSeen.has(pvKey)) {
          pvSeen.add(pvKey);
          pvEluate.push({ site, ratio: medium });
        }
      }
    });
  });

  const labs = [];
  LAB_ORDER.forEach((lab) => {
    const media = [];
    MEDIUM_ORDER.forEach((medium) => {
      const items = buildMediumItems(labMediumBuckets[lab][medium], medium, groupDefs);
      if (items.length > 0) {
        media.push({ medium, items });
      }
    });
    if (media.length > 0) {
      labs.push({ lab, media });
    }
  });

  return {
    vorOrt: dedupeSorted(vorOrtTokens),
    pv: {
      ts: false,
      eluate: pvEluate,
      itemsBySite: {
        E: dedupeSorted(pvItemsBySite.E),
        B: dedupeSorted(pvItemsBySite.B),
      },
    },
    labs,
    extern: [],
  };
}

module.exports = {
  mapTogglesToSelection,
};
