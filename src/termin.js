function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}

function formatDateYmd(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function toDateFromYmd(receivedDate) {
  if (typeof receivedDate !== 'string') {
    return null;
  }

  const value = receivedDate.trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return null;
  }

  const parsed = new Date(`${value}T00:00:00`);
  if (!Number.isFinite(parsed.getTime())) {
    return null;
  }

  return parsed;
}

function addBusinessDays(startDate, workdays) {
  const cursor = new Date(startDate.getTime());
  let added = 0;

  while (added < workdays) {
    cursor.setDate(cursor.getDate() + 1);
    if (isWeekend(cursor)) {
      continue;
    }
    added += 1;
  }

  return cursor;
}

function calculateTermin(receivedDate, isRush) {
  const startDate = toDateFromYmd(receivedDate);
  if (!startDate) {
    return null;
  }

  const workdays = isRush ? 2 : 4;
  const result = addBusinessDays(startDate, workdays);
  return formatDateYmd(result);
}

module.exports = {
  calculateTermin,
};
