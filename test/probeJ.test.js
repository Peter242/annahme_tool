const test = require('node:test');
const assert = require('node:assert/strict');
const { buildProbeJ } = require('../src/probeJ');

test('buildProbeJ returns empty text when gewicht, geruch and bemerkung are all empty', () => {
  assert.equal(buildProbeJ({ gewicht: '', geruch: '', bemerkung: '' }), '');
  assert.equal(buildProbeJ({ gewicht: null, geruch: undefined, bemerkung: '   ' }), '');
});

test('buildProbeJ renders only gewicht when set', () => {
  assert.equal(buildProbeJ({ gewicht: 2, geruch: '', bemerkung: '' }), 'Gewicht: 2 kg');
});

test('buildProbeJ renders gewicht and geruch without placeholders', () => {
  assert.equal(
    buildProbeJ({ gewicht: 2, geruch: 'muffig', bemerkung: '' }),
    'Gewicht: 2 kg; Geruch: muffig',
  );
});

test('buildProbeJ renders only bemerkung text when only bemerkung is set', () => {
  assert.equal(
    buildProbeJ({ gewicht: null, geruch: undefined, bemerkung: 'wenig Material' }),
    'wenig Material',
  );
});

test('buildProbeJ falls back to geruchAuffaelligkeit when geruch is empty', () => {
  assert.equal(
    buildProbeJ({ gewicht: '', geruch: '', geruchAuffaelligkeit: 'neutral', bemerkung: '' }),
    'Geruch: neutral',
  );
});
