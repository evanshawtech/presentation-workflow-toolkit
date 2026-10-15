'use strict';
// Process-local guard for the installed image-size v1 dependency. Its detector
// calls validate() BEFORE disableTypes() is checked; replacing calculate alone
// leaves the JXL box-walking validator reachable. Reject by bounded header reads
// before either vulnerable validator or calculator runs. No upstream file edits.
const {createRequire} = require('node:module');
const {existsSync} = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const manifests = [
  'skills-src/presentation-skill/package.json',
  'skills-src/dashi-ppt-skill/skills/dashi-ppt/project/package.json',
  'skills-src/dashi-ppt-skill/skills/dashi-ppt/project/packages/html-deck-to-pptx/package.json',
];
const guarded = new Set();
function ascii(input, start, value) {
  if (!input || input.length < start + value.length) return false;
  for (let i = 0; i < value.length; i++) {
    if (input[start + i] !== value.charCodeAt(i)) return false;
  }
  return true;
}
const signatures = {
  icns: input => ascii(input, 0, 'icns'),
  jxl: input => ascii(input, 4, 'JXL '),
  'jxl-stream': input => input && input.length >= 2 && input[0] === 0xff && input[1] === 0x0a,
  heif: input => ascii(input, 4, 'ftyp'),
};
function install(localRequire) {
  let entry;
  try { entry = localRequire.resolve('image-size'); }
  catch (err) { if (err.code === 'MODULE_NOT_FOUND') return; throw err; }
  if (guarded.has(entry)) return;
  const mod = localRequire('image-size');
  const handlersFile = path.join(path.dirname(entry), 'types', 'index.js');
  if (typeof mod.disableTypes !== 'function' || !existsSync(handlersFile)) {
    throw new Error('image-size API changed; review the guard before rendering');
  }
  const handlers = localRequire(handlersFile).typeHandlers;
  for (const [type, matches] of Object.entries(signatures)) {
    if (!handlers || !handlers[type] || typeof handlers[type].validate !== 'function') {
      throw new Error('image-size handlers changed; review the guard before rendering');
    }
    const reject = () => { throw new TypeError('Toolkit blocked image format: ' + type); };
    handlers[type].validate = input => { if (matches(input)) reject(); return false; };
    handlers[type].calculate = reject;
  }
  mod.disableTypes(Object.keys(signatures));
  guarded.add(entry);
}
for (const relative of manifests) {
  const manifest = path.join(root, relative);
  if (!existsSync(manifest)) continue;
  const localRequire = createRequire(manifest);
  install(localRequire);
  // Also resolve from pptxgenjs so a future nested image-size copy is protected.
  try { install(createRequire(localRequire.resolve('pptxgenjs'))); }
  catch (err) {
    if (err.code !== 'MODULE_NOT_FOUND') throw err;
  }
}
module.exports = {guardedModuleCount: guarded.size};
