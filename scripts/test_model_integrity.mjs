import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const root = new URL('../', import.meta.url);
const app = readFileSync(new URL('gdc_app.js', root), 'utf8');
const data = readFileSync(new URL('gdc_data.js', root), 'utf8');
const excel = readFileSync(new URL('model_export.js', root), 'utf8');

const fail = (message) => { throw new Error(message); };
const near = (actual, expected, tolerance, label) => {
  if (!Number.isFinite(actual) || Math.abs(actual - expected) > tolerance) {
    fail(`${label}: expected ${expected}, received ${actual}`);
  }
};

const mStart = app.indexOf('const M={');
const defaultsStart = app.indexOf('const DEFAULTS_JSON', mStart);
const mEnd = defaultsStart < 0 ? -1 : app.lastIndexOf('};', defaultsStart);
if (mStart < 0 || mEnd < 0) fail('Could not locate the default model state.');
const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(`${data}\n${app.slice(mStart, mEnd + 2)}\nglobalThis.audit={M,TRANCHE};`, sandbox);
const { M, TRANCHE } = sandbox.audit;
const fleetAC = M.wind.mw + M.solar.mw / M.conn.dcac;
const t1 = Math.min(TRANCHE.t1MW, fleetAC);
const blend = (t1 * TRANCHE.p1 + Math.max(0, fleetAC - t1) * TRANCHE.p2) / fleetAC;
near(M.wind.ppa, +blend.toFixed(1), 1e-9, 'default wind PPA');
near(M.solar.ppa, +blend.toFixed(1), 1e-9, 'default solar PPA');
if (TRANCHE.enabled !== true) fail('Default tranche pricing must be enabled.');

if (!/let soc=0, served=0/.test(app)) fail('Supply dispatch must not start with free tradeable energy.');

const cleanStart = app.indexOf('function cleanState(');
const cleanEnd = app.indexOf('\nfunction restore(', cleanStart);
if (cleanStart < 0 || cleanEnd < 0) fail('Could not locate scenario validation.');
vm.runInContext(`${app.slice(cleanStart, cleanEnd)}\nglobalThis.cleanState=cleanState;`, sandbox);
const malicious = sandbox.cleanState({
  dc: { srcMode: '\"><img src=x onerror=alert(1)>', firmMW: 500 },
  battery: { gridYear: Infinity },
  unknown: { injected: true },
}, M);
if (malicious.dc.srcMode !== undefined) fail('Scenario validation accepted markup in a string field.');
if (malicious.dc.firmMW !== 500) fail('Scenario validation rejected a valid numeric field.');
if (malicious.battery.gridYear !== undefined) fail('Scenario validation accepted a non-finite number.');
if ('unknown' in malicious) fail('Scenario validation accepted an unknown top-level field.');

for (const required of ['ROUNDDOWN(', 'B_AVGBUY', 'B_AVGSELL', 'B_FSELF', 'B_MKT', "['FEE_C'"]) {
  if (!excel.includes(required)) fail(`Excel reconciliation input/formula missing: ${required}`);
}
if (/Small diffs vs dashboard are timing conventions/.test(excel)) fail('Workbook still excuses reconciliation differences as timing conventions.');
if (/const\s+TERMS_BLOB\s*=/.test(data)) fail('Encrypted commercial terms must not ship in the public data bundle.');

console.log('Model integrity regression checks passed.');
