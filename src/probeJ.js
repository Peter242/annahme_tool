const PLACEHOLDER_VALUES = new Set(['-', '—', 'keine', 'unauffaellig', 'unauffällig', 'n/a']);

function normalizeText(value) {
  return String(value).trim().toLocaleLowerCase('de-DE');
}

function isBlank(value) {
  if (value === null || value === undefined) {
    return true;
  }
  const normalized = normalizeText(value);
  return normalized === '' || PLACEHOLDER_VALUES.has(normalized);
}

function asTrimmedText(value) {
  return String(value).trim();
}

function hasMeaningfulWeight(value) {
  if (isBlank(value)) {
    return false;
  }
  const trimmed = asTrimmedText(value);
  const numericValue = Number.parseFloat(trimmed.replace(',', '.'));
  if (Number.isFinite(numericValue)) {
    return numericValue > 0;
  }
  return normalizeText(trimmed) !== '0';
}

function buildProbeJ(probe = {}) {
  const parts = [];

  if (hasMeaningfulWeight(probe.gewicht)) {
    parts.push(`Gewicht: ${asTrimmedText(probe.gewicht)} kg`);
  }

  const geruchSource = !isBlank(probe.geruch) ? probe.geruch : probe.geruchAuffaelligkeit;
  if (!isBlank(geruchSource)) {
    parts.push(`Geruch: ${asTrimmedText(geruchSource)}`);
  }

  if (!isBlank(probe.bemerkung)) {
    parts.push(asTrimmedText(probe.bemerkung));
  }

  return parts.length > 0 ? parts.join('; ') : '';
}

module.exports = {
  buildProbeJ,
};
