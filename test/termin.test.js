const test = require('node:test');
const assert = require('node:assert/strict');
const { calculateTermin } = require('../src/termin');

test('calculateTermin uses 2 business days for rush', () => {
  // Friday +2 business days = Tuesday
  const result = calculateTermin('2026-02-13', true);
  assert.equal(result, '2026-02-17');
});

test('calculateTermin uses 4 business days for normal', () => {
  // Friday +4 business days = Thursday
  const result = calculateTermin('2026-02-13', false);
  assert.equal(result, '2026-02-19');
});

test('calculateTermin returns null for invalid date', () => {
  const result = calculateTermin('13.02.2026', true);
  assert.equal(result, null);
});
