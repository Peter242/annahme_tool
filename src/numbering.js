function normalizeDate(date) {
  if (date instanceof Date) {
    return date;
  }

  const parsed = new Date(date);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }
  return parsed;
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function formatOrderDate(date) {
  const normalized = normalizeDate(date);
  if (!normalized) {
    return null;
  }

  const day = pad2(normalized.getDate());
  const month = pad2(normalized.getMonth() + 1);
  const year = pad2(normalized.getFullYear() % 100);
  return `${day}${month}${year}`;
}

function makeOrderNumber(date, xy) {
  const formatted = formatOrderDate(date);
  if (!formatted) {
    return null;
  }

  const xyPad2 = pad2(xy);
  return `${formatted}8${xyPad2}`;
}

function nextLabNumbers(lastLab, count) {
  if (!Number.isInteger(lastLab) || !Number.isInteger(count) || count < 0) {
    return [];
  }

  return Array.from({ length: count }, (_, idx) => lastLab + idx + 1);
}

module.exports = {
  formatOrderDate,
  makeOrderNumber,
  nextLabNumbers,
};
