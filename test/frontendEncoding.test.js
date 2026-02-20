const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

function readPublic(file) {
  return fs.readFileSync(path.join(__dirname, '..', 'public', file), 'utf8');
}

test('settings page keeps Zurück and Zurücksetzen text in DOM source', () => {
  const html = readPublic('settings.html');
  assert.match(html, />Zurück</);
  assert.match(html, />Cache zurücksetzen</);
});

test('frontend no longer ships mojibake decoder hacks', () => {
  const indexHtml = readPublic('index.html');
  const packagesHtml = readPublic('packages.html');

  assert.equal(indexHtml.includes('decodeMojibakeText'), false);
  assert.equal(indexHtml.includes('normalizeUiEncoding'), false);
  assert.equal(packagesHtml.includes('decodeMojibakeText'), false);
  assert.equal(packagesHtml.includes('normalizeUiEncoding'), false);
});

test('frontend html source has no typical mojibake markers', () => {
  const indexHtml = readPublic('index.html');
  const settingsHtml = readPublic('settings.html');
  const packagesHtml = readPublic('packages.html');
  const joined = [indexHtml, settingsHtml, packagesHtml].join('\n');

  assert.equal(/Ã|ï¿½/.test(joined), false);
});
