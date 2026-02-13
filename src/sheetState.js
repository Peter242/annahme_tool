function pad2(value) {
  return String(value).padStart(2, '0');
}

function buildTodayPrefix(now) {
  const day = pad2(now.getDate());
  const month = pad2(now.getMonth() + 1);
  const year = pad2(now.getFullYear() % 100);
  return `${day}${month}${year}8`;
}

function cellValueToString(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'string') {
    return value;
  }

  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }

  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => (part && part.text ? String(part.text) : '')).join('');
    }
    if (value.text !== undefined && value.text !== null) {
      return String(value.text);
    }
    if (value.result !== undefined && value.result !== null) {
      return String(value.result);
    }
  }

  return String(value);
}

function rowHasContentInAToJ(sheet, rowNumber) {
  for (let column = 1; column <= 10; column += 1) {
    const value = sheet.getRow(rowNumber).getCell(column).value;
    if (cellValueToString(value).trim() !== '') {
      return true;
    }
  }
  return false;
}

function getSheetState(sheet, now = new Date()) {
  const todayPrefix = buildTodayPrefix(now);
  const orderRegex = new RegExp(`^${todayPrefix}(\\d{2})(?!\\d)`);

  let lastUsedRow = 0;
  let maxLabNumber = 0;
  let maxOrderSeqToday = 0;

  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (rowHasContentInAToJ(sheet, rowNumber)) {
      lastUsedRow = rowNumber;
    }

    const colA = cellValueToString(sheet.getRow(rowNumber).getCell(1).value).trim();
    if (!colA) {
      continue;
    }

    const isOrderHeader = /^\d{9}(?!\d)/.test(colA);

    const orderMatch = colA.match(orderRegex);
    if (orderMatch) {
      const seq = Number.parseInt(orderMatch[1], 10);
      if (Number.isFinite(seq) && seq > maxOrderSeqToday) {
        maxOrderSeqToday = seq;
      }
    }

    if (isOrderHeader) {
      continue;
    }

    const labMatch = colA.match(/^(\d{5,6})(?!\d)/);
    if (!labMatch) {
      continue;
    }

    const labNumber = Number.parseInt(labMatch[1], 10);
    if (Number.isFinite(labNumber) && labNumber > maxLabNumber) {
      maxLabNumber = labNumber;
    }
  }

  return {
    lastUsedRow,
    maxLabNumber,
    maxOrderSeqToday,
  };
}

module.exports = {
  getSheetState,
  buildTodayPrefix,
};
