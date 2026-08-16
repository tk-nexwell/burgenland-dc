import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const app = readFileSync(new URL('../gdc_app.js', import.meta.url), 'utf8');
const match = app.match(/function xirr\(cfs,dt\)\{[\s\S]*?\r?\n\}\r?\n\/\/ Net delivered CF/);
if (!match) throw new Error('Could not locate xirr() in gdc_app.js');

const sandbox = {};
const source = match[0].replace(/\n\/\/ Net delivered CF[\s\S]*$/, '') + '\nglobalThis.xirr = xirr;';
vm.runInNewContext(source, sandbox);

const near = (actual, expected, tolerance, label) => {
  if (!Number.isFinite(actual) || Math.abs(actual - expected) > tolerance) {
    throw new Error(`${label}: expected ${expected}, received ${actual}`);
  }
};

near(sandbox.xirr([-100, 110], [0, 1]), 0.10, 1e-8, 'positive return');
near(sandbox.xirr([-100, 110, -1e-14], [0, 1, 40]), 0.10, 1e-8, 'immaterial late tail');
near(sandbox.xirr([-100, 90], [0, 1]), -0.10, 1e-8, 'genuine negative return');

if (!Number.isNaN(sandbox.xirr([-100, 230, -132], [0, 1, 2]))) {
  throw new Error('Multiple-sign-change cash flow should not report an ambiguous IRR');
}

console.log('XIRR regression checks passed.');
