const test = require('node:test');
const assert = require('node:assert/strict');
const { buildParameterTextFromSelection } = require('../src/parameterTextBuilder');

test('empty selection returns empty string', () => {
  assert.equal(buildParameterTextFromSelection(), '');
  assert.equal(buildParameterTextFromSelection(null), '');
  assert.equal(buildParameterTextFromSelection({}), '');
});

test('PV only TS renders correctly', () => {
  const result = buildParameterTextFromSelection({
    pv: { ts: true },
  });
  assert.equal(result, 'PV: TS');
});

test('PV TS plus eluates renders correctly', () => {
  const result = buildParameterTextFromSelection({
    pv: {
      ts: true,
      eluate: [
        { site: 'E', ratio: '2e' },
        { site: 'B', ratio: '10e' },
      ],
    },
  });
  assert.equal(result, 'PV: TS, E(2e), B(10e)');
});

test('EMD FS groups AN items first and keeps remaining order', () => {
  const result = buildParameterTextFromSelection({
    labs: [
      {
        lab: 'EMD',
        media: [
          {
            medium: 'FS',
            items: ['GV', 'AN:Cl', 'AN:Br'],
          },
        ],
      },
    ],
  });
  assert.equal(result, 'EMD: FS: AN(Cl, Br), GV');
});

test('EMD media are rendered in fixed order and separated by semicolon', () => {
  const result = buildParameterTextFromSelection({
    labs: [
      {
        lab: 'EMD',
        media: [
          { medium: '10e', items: ['SO4'] },
          { medium: 'FS', items: ['AN:Cl', 'GV'] },
          { medium: '2e', items: ['pH'] },
        ],
      },
    ],
  });
  assert.equal(result, 'EMD: FS: AN(Cl), GV; 2e: pH; 10e: SO4');
});

test('extern NLGA legionella renders correctly', () => {
  const result = buildParameterTextFromSelection({
    extern: [
      { lab: 'NLGA', items: ['Legionellen'] },
    ],
  });
  assert.equal(result, 'NLGA: Legionellen');
});

test('combined multiline output keeps line order PV, EMD, HB, extern', () => {
  const result = buildParameterTextFromSelection({
    pv: {
      ts: true,
      eluate: [{ site: 'E', ratio: '2e' }],
    },
    labs: [
      {
        lab: 'HB',
        media: [{ medium: 'H2O', items: ['AN:Cl', 'el Lf'] }],
      },
      {
        lab: 'EMD',
        media: [{ medium: 'FS', items: ['GV'] }],
      },
    ],
    extern: [
      { lab: 'NLGA', items: ['Legionellen'] },
      { lab: 'GBA', items: ['PAK'] },
    ],
  });

  assert.equal(
    result,
    [
      'PV: TS, E(2e)',
      'EMD: FS: GV',
      'HB: H2O: AN(Cl), el Lf',
      'NLGA: Legionellen',
      'GBA: PAK',
    ].join('\n'),
  );
});
