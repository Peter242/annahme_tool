const test = require('node:test');
const assert = require('node:assert/strict');
const {
  normalizeContainerItems,
  normalizeContainers,
  renderContainers,
  renderContainersSummary,
  renderColumnHFromProbe,
} = require('../src/containers');

test('normalizeContainerItems keeps valid tokens in sequence', () => {
  const items = normalizeContainerItems(['K:30mL+HCl', 'G:1L', 'invalid', '', 'K:30mL+HCl']);
  assert.deepEqual(items, ['K:30mL+HCl', 'G:1L', 'K:30mL+HCl']);
});

test('normalizeContainers keeps mode and token items', () => {
  const normalized = normalizeContainers({
    mode: 'perSample',
    items: ['K:30mL+HCl', 'K:30mL+HCl', 'G:1L'],
  });

  assert.equal(normalized.mode, 'perSample');
  assert.deepEqual(normalized.items, ['K:30mL+HCl', 'K:30mL+HCl', 'G:1L']);
});

test('renderContainers renders grouped token summary in default order', () => {
  const text = renderContainers(['K:30mL+HCl', 'K:30mL+HCl', 'K:30mL+HNO3', 'G:1L']);
  assert.equal(text, 'Kunststoff (2x 30mL + HCl; 30mL + HNO3) Glas (1L)');
});

test('renderContainersSummary reads items from containers object', () => {
  const text = renderContainersSummary({ mode: 'perSample', items: ['K:250mL', 'G:1L+H2SO4'] });
  assert.equal(text, 'Kunststoff (250mL) Glas (1L + H2SO4)');
});

test('renderColumnHFromProbe supports material-only and full modes', () => {
  const probe = {
    material: 'Boden',
    containers: {
      mode: 'perSample',
      items: ['K:30mL+HCl', 'K:30mL+HCl'],
    },
  };

  assert.equal(renderColumnHFromProbe(probe), 'Boden, Kunststoff (2x 30mL + HCl)');
  assert.equal(renderColumnHFromProbe(probe, { onlyMaterial: true }), 'Boden');
});

test('renderContainers omits 1x prefix and keeps 2x prefix for glass examples', () => {
  assert.equal(renderContainers(['G:HS+MeOH']), 'Glas (HS + MeOH)');
  assert.equal(renderContainers(['G:250mL', 'G:250mL', 'G:1L']), 'Glas (1L; 2x 250mL)');
});
