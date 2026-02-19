const test = require('node:test');
const assert = require('node:assert/strict');
const { spawnSync } = require('child_process');

test('powershell utf8 stdin/stdout roundtrip keeps umlauts', () => {
  if (process.platform !== 'win32') {
    return;
  }

  const payload = {
    value: 'Umlaute: äöüÄÖÜß',
  };
  const psCommand = [
    "$utf8=[System.Text.UTF8Encoding]::new($false)",
    '[Console]::InputEncoding=$utf8',
    '[Console]::OutputEncoding=$utf8',
    '$json=[Console]::In.ReadToEnd()',
    '$obj=$json | ConvertFrom-Json',
    '@{ ok=$true; echo=[string]$obj.value } | ConvertTo-Json -Compress',
  ].join('; ');

  const result = spawnSync('powershell.exe', [
    '-NoProfile',
    '-Command',
    psCommand,
  ], {
    input: JSON.stringify(payload),
    encoding: 'utf8',
  });

  assert.equal(result.status, 0);
  const output = String(result.stdout || '').trim();
  assert.ok(output.length > 0);
  const lastLine = output.split(/\r?\n/).filter(Boolean).pop();
  const parsed = JSON.parse(lastLine);
  assert.equal(parsed.ok, true);
  assert.equal(parsed.echo, payload.value);
});
