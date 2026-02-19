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

function extractOrderCore(valueString) {
  const normalized = cellValueToString(valueString).trim();
  if (!normalized) {
    return null;
  }

  const match = normalized.match(/^(\d{6}8\d{2})(?!\d)/);
  if (!match) {
    return null;
  }

  const core = match[1];
  return core;
}

function extractLeadingDigits(valueString) {
  const normalized = cellValueToString(valueString).trim();
  if (!normalized) {
    return '';
  }
  const match = normalized.match(/^(\d+)/);
  return match ? match[1] : '';
}

function extractOrderSeqForTodayPrefix(valueString, todayPrefix) {
  const core = extractOrderCore(valueString);
  if (!core || !core.startsWith(todayPrefix)) {
    return null;
  }

  const seq = Number.parseInt(core.slice(-2), 10);
  return Number.isFinite(seq) ? seq : null;
}

function resolveYearPrefix(sheet, now = new Date()) {
  const sheetName = String(sheet?.name || '').trim();
  const fromSheet = sheetName.match(/^(\d{4})$/);
  const year = fromSheet ? Number.parseInt(fromSheet[1], 10) : now.getFullYear();
  return pad2(year % 100);
}

function extractLeadingLabNumber(valueString) {
  const normalized = cellValueToString(valueString).trim();
  if (!normalized) {
    return null;
  }

  if (extractOrderCore(normalized)) {
    return null;
  }
  const match = normalized.match(/^(\d{5,6})([A-Za-z]|-\d+)?$/);
  if (!match) {
    return null;
  }
  const digits = match[1];

  const parsed = Number.parseInt(digits, 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function isLikelyLabNo(valueString, yearPrefix) {
  const extracted = extractLeadingLabNumber(valueString);
  if (!Number.isFinite(extracted)) {
    return false;
  }
  return true;
}

function scanLabNumberCandidates(sheet, options = {}) {
  const debugScan = Boolean(options.debugScan);
  const debugLogger = typeof options.debugLogger === 'function'
    ? options.debugLogger
    : console.log;
  const candidates = [];
  let maxLabNumber = 0;
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const rawColA = sheet.getRow(rowNumber).getCell(1).value;
    const rawText = cellValueToString(rawColA).trim();
    if (!rawText) {
      continue;
    }
    const parsed = extractLeadingLabNumber(rawColA);
    if (!Number.isFinite(parsed)) {
      if (debugScan) {
        const startsWithDigits = /^\d/.test(rawText);
        const hasEightDigitLead = /^\d{8}/.test(rawText);
        if (startsWithDigits && !extractOrderCore(rawText)) {
          const reason = hasEightDigitLead ? 'ignored_foreign_labno' : 'ignored_non_matching_labno';
          debugLogger(`[sheet-scan] row=${rowNumber} value="${rawText}" reason=${reason}`);
        }
      }
      continue;
    }
    candidates.push({
      row: rowNumber,
      raw: rawText,
      parsed,
    });
    if (parsed > maxLabNumber) {
      maxLabNumber = parsed;
    }
  }
  return {
    maxLabNumber,
    candidates,
  };
}

function getSheetState(sheet, now = new Date(), options = {}) {
  const debugScan = Boolean(options.debugScan);
  const debugLogger = typeof options.debugLogger === 'function'
    ? options.debugLogger
    : console.log;
  const todayPrefix = buildTodayPrefix(now);

  let lastUsedRow = 0;
  let maxLabNumber = 0;
  let maxOrderSeqToday = 0;
  const labScan = scanLabNumberCandidates(sheet, { debugScan, debugLogger });
  maxLabNumber = labScan.maxLabNumber;
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (rowHasContentInAToJ(sheet, rowNumber)) {
      lastUsedRow = rowNumber;
    }

    const rawColA = sheet.getRow(rowNumber).getCell(1).value;
    const colA = cellValueToString(rawColA).trim();
    if (!colA) {
      continue;
    }

    const seqForToday = extractOrderSeqForTodayPrefix(rawColA, todayPrefix);
    if (Number.isFinite(seqForToday)) {
      if (seqForToday > maxOrderSeqToday) {
        maxOrderSeqToday = seqForToday;
      }
      continue;
    }

    if (debugScan) {
      const hasOrderDigits = /^\d{6}\d{2,}/.test(colA);
      if (hasOrderDigits && !extractOrderCore(colA)) {
        debugLogger(`[sheet-scan] row=${rowNumber} value="${colA}" reason=ignored_foreign_order`);
      }
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
  extractOrderCore,
  extractOrderSeqForTodayPrefix,
  isLikelyLabNo,
  extractLeadingLabNumber,
  scanLabNumberCandidates,
};
