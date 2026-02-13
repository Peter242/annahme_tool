const test = require('node:test');
const assert = require('node:assert/strict');
const { makeOrderNumber, nextLabNumbers } = require('../src/numbering');

test('makeOrderNumber builds TTMMYY8xy', () => {
  const result = makeOrderNumber(new Date('2026-02-12T00:00:00Z'), 1);
  assert.equal(result, '120226801');
});

test('makeOrderNumber pads xy to 2 digits', () => {
  const result = makeOrderNumber(new Date('2026-02-12T00:00:00Z'), 9);
  assert.equal(result, '120226809');
});

test('nextLabNumbers returns consecutive numbers', () => {
  const result = nextLabNumbers(26203, 4);
  assert.deepEqual(result, [26204, 26205, 26206, 26207]);
});

test('nextLabNumbers returns empty list for count 0', () => {
  const result = nextLabNumbers(26203, 0);
  assert.deepEqual(result, []);
});
