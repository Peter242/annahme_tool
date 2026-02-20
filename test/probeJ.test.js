const test = require('node:test');
const assert = require('node:assert/strict');
const { buildProbeJ } = require('../src/probeJ');

test('buildProbeJ returns empty text for fully empty input', () => {
  assert.equal(buildProbeJ({}), '');
});

test('buildProbeJ treats placeholder geruch as empty', () => {
  assert.equal(buildProbeJ({ geruch: 'unauffaellig' }), '');
  assert.equal(buildProbeJ({ geruch: 'unauffällig' }), '');
});

test('buildProbeJ treats placeholder bemerkung as empty', () => {
  assert.equal(buildProbeJ({ bemerkung: '-' }), '');
});

test('buildProbeJ treats weight zero as empty', () => {
  assert.equal(buildProbeJ({ gewicht: 0 }), '');
  assert.equal(buildProbeJ({ gewicht: '0' }), '');
});

test('buildProbeJ renders weight when meaningful', () => {
  assert.equal(buildProbeJ({ gewicht: 2 }), 'Gewicht: 2 kg');
});

test('buildProbeJ renders gewicht and geruch together when both are meaningful', () => {
  const rendered = buildProbeJ({ gewicht: 2, geruch: 'muffig' });
  assert.match(rendered, /Gewicht: 2 kg/);
  assert.match(rendered, /Geruch: muffig/);
});

test('buildProbeJ renders only bemerkung when meaningful', () => {
  assert.equal(buildProbeJ({ bemerkung: 'wenig Material' }), 'wenig Material');
});
