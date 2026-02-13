const fs = require('fs');
const path = require('path');

const PACKAGES_PATH = path.join(__dirname, '..', '..', 'data', 'packages.json');

function readPackages() {
  if (!fs.existsSync(PACKAGES_PATH)) {
    return [];
  }

  const raw = fs.readFileSync(PACKAGES_PATH, 'utf-8');
  if (!raw.trim()) {
    return [];
  }

  const parsed = JSON.parse(raw);
  return Array.isArray(parsed) ? parsed : [];
}

function writePackages(packages) {
  const dir = path.dirname(PACKAGES_PATH);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(PACKAGES_PATH, `${JSON.stringify(packages, null, 2)}\n`, 'utf-8');
}

module.exports = {
  readPackages,
  writePackages,
};
