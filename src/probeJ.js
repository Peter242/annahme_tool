function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

function asTrimmedText(value) {
  return String(value).trim();
}

function buildProbeJ(probe = {}) {
  const parts = [];

  if (!isBlank(probe.gewicht)) {
    parts.push(`Gewicht: ${asTrimmedText(probe.gewicht)} kg`);
  }

  const geruchSource = !isBlank(probe.geruch) ? probe.geruch : probe.geruchAuffaelligkeit;
  if (!isBlank(geruchSource)) {
    parts.push(`Geruch: ${asTrimmedText(geruchSource)}`);
  }

  if (!isBlank(probe.bemerkung)) {
    parts.push(asTrimmedText(probe.bemerkung));
  }

  return parts.join('; ');
}

module.exports = {
  buildProbeJ,
};
