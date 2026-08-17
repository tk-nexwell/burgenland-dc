import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const root = new URL('../', import.meta.url);
const app = readFileSync(new URL('gdc_app.js', root), 'utf8');
const data = readFileSync(new URL('gdc_data.js', root), 'utf8');
const html = readFileSync(new URL('index.html', root), 'utf8');

const fail = (message) => { throw new Error(message); };
const check = (condition, message) => { if (!condition) fail(message); };
const near = (actual, expected, tolerance, label) => {
  if (!Number.isFinite(actual) || Math.abs(actual - expected) > tolerance) {
    fail(`${label}: expected ${expected}, received ${actual}`);
  }
};
const section = (source, from, to, label) => {
  const start = source.indexOf(from);
  const end = source.indexOf(to, start + from.length);
  if (start < 0 || end < 0) fail(`Could not locate ${label}.`);
  return source.slice(start, end);
};
const clone = (value) => JSON.parse(JSON.stringify(value));

// The public build must not rely on a browser-side password or ship an encrypted terms payload.
check(!/\bTERMS_BLOB\b/.test(data), 'Encrypted internal-terms payload remains in the public data bundle.');
check(!/\bgateTry\s*\(/.test(app), 'Interactive client-side gate logic remains in the public application.');
check(!/<input\b[^>]*\btype\s*=\s*["']password["']/i.test(html), 'Password input remains in the public page.');
check(!/\b(?:pbkdf2|deriveKey|crypto\.subtle\.digest)\b/i.test(app), 'Client-side password verification remains in the public application.');

const gateContext = {};
vm.createContext(gateContext);
vm.runInContext(
  `${section(app, 'const GATE=', '/* ============================================================', 'public gate stub')}\n` +
  'globalThis.audit={GATE,gateLocked};',
  gateContext,
);
check(gateContext.audit.gateLocked('overview') === false, 'The public application can still lock a page.');
check(gateContext.audit.GATE.open === true, 'The public gate stub is not explicitly open.');
check(Array.isArray(gateContext.audit.GATE.tabs) && gateContext.audit.GATE.tabs.length === 0, 'The public gate stub still identifies restricted tabs.');
check(Object.keys(gateContext.audit.GATE).every((key) => ['tabs', 'open'].includes(key)), 'The public gate stub contains unexpected verifier material.');

// Load the real public tranche, defaults and scenario sanitizer without starting the browser app.
const trancheContext = {};
vm.createContext(trancheContext);
vm.runInContext(
  `${section(data, 'const TRANCHE=', '\nconst TECH', 'public tranche assumptions')}\n` +
  'globalThis.auditTrancheData=TRANCHE;',
  trancheContext,
);
const publicTranche = clone(trancheContext.auditTrancheData);
const stateContext = {
  TRANCHE: clone(publicTranche),
  buildNav: () => {},
  render: () => {},
  _SUPC: null,
  _CLIPF: null,
};
vm.createContext(stateContext);
vm.runInContext(
  `${section(app, 'const M={', 'const Y0=', 'default model state')}\n` +
  'globalThis.auditState={M,DEFAULTS_JSON};',
  stateContext,
);
const defaults = clone(stateContext.auditState.M);

vm.runInContext(
  `${section(app, 'function solarAC(){', 'function fleetDC(){', 'AC fleet calculation')}\n` +
  `${section(app, 'function trancheBlend(){', '// Residual (non-RES) power:', 'tranche synchronization')}\n` +
  'globalThis.auditTranche={trancheBlend,syncTranchePpa,resetAll,fleetAC};',
  stateContext,
);
const trancheFns = stateContext.auditTranche;
const fleetAC = trancheFns.fleetAC();
const trancheOneMW = Math.min(publicTranche.t1MW, fleetAC);
const expectedBlend = +(
  (trancheOneMW * publicTranche.p1 + Math.max(0, fleetAC - trancheOneMW) * publicTranche.p2) / fleetAC
).toFixed(1);
near(defaults.wind.ppa, expectedBlend, 1e-12, 'default wind PPA tranche blend');
near(defaults.solar.ppa, expectedBlend, 1e-12, 'default solar PPA tranche blend');

stateContext.auditState.M.wind.ppa = 1;
stateContext.auditState.M.solar.ppa = 2;
near(trancheFns.syncTranchePpa(), expectedBlend, 1e-12, 'synchronized tranche blend');
near(stateContext.auditState.M.wind.ppa, expectedBlend, 1e-12, 'synchronized wind PPA');
near(stateContext.auditState.M.solar.ppa, expectedBlend, 1e-12, 'synchronized solar PPA');

stateContext.TRANCHE.p1 = 1;
stateContext.TRANCHE.p2 = 2;
stateContext.auditState.M.wind.ppa = 3;
stateContext.auditState.M.solar.ppa = 4;
trancheFns.resetAll();
near(stateContext.TRANCHE.p1, publicTranche.p1, 1e-12, 'reset tranche-one price');
near(stateContext.TRANCHE.p2, publicTranche.p2, 1e-12, 'reset tranche-two price');
const resetBlend = +trancheFns.trancheBlend().toFixed(1);
near(stateContext.auditState.M.wind.ppa, resetBlend, 1e-12, 'reset wind PPA tranche blend');
near(stateContext.auditState.M.solar.ppa, resetBlend, 1e-12, 'reset solar PPA tranche blend');

vm.runInContext(
  `${section(app, 'const MODEL_SCHEMA=', '\nfunction restore(', 'scenario validation')}\n` +
  'globalThis.auditCleaner={MODEL_SCHEMA,modelNumberOK,cleanModel};',
  stateContext,
);
const { MODEL_SCHEMA, cleanModel } = stateContext.auditCleaner;

const poisoned = JSON.parse(`{
  "dc":{"srcMode":"\\\"><img src=x onerror=alert(1)>","firmMW":6001,"marginMode":"flat"},
  "battery":{"durationH":3,"captureFactor":9,"gridYear":null},
  "macro":{"infl":-1},
  "wind":null,
  "unknown":{"injected":true},
  "__proto__":{"polluted":true}
}`);
const cleaned = clone(cleanModel(poisoned, MODEL_SCHEMA));
check(cleaned.dc.srcMode === defaults.dc.srcMode, 'Scenario sanitizer accepted markup in an enum field.');
check(cleaned.dc.firmMW === defaults.dc.firmMW, 'Scenario sanitizer accepted an out-of-range firm load.');
check(cleaned.dc.marginMode === 'flat', 'Scenario sanitizer rejected a valid enum value.');
check(cleaned.battery.durationH === defaults.battery.durationH, 'Scenario sanitizer accepted a non-whitelisted battery duration.');
check(cleaned.battery.captureFactor === defaults.battery.captureFactor, 'Scenario sanitizer accepted an out-of-range capture factor.');
check(cleaned.battery.gridYear === defaults.battery.gridYear, 'Scenario sanitizer accepted an invalid year.');
check(cleaned.macro.infl === defaults.macro.infl, 'Scenario sanitizer accepted an out-of-range percentage.');
check(JSON.stringify(cleaned.wind) === JSON.stringify(defaults.wind), 'A corrupt partial object erased wind defaults.');
check(!Object.prototype.hasOwnProperty.call(cleaned, 'unknown'), 'Scenario sanitizer accepted an unknown top-level field.');
check(!Object.prototype.hasOwnProperty.call(cleaned, '__proto__'), 'Scenario sanitizer accepted a prototype key.');

const partial = clone(cleanModel({ dc: { firmMW: 750 } }, MODEL_SCHEMA));
check(partial.dc.firmMW === 750, 'Scenario sanitizer rejected a valid partial update.');
check(partial.dc.srcMode === defaults.dc.srcMode, 'Scenario sanitizer did not preserve a missing nested default.');
check(partial.battery.durationH === defaults.battery.durationH, 'Scenario sanitizer did not preserve an omitted model branch.');
check(cleanModel(null, MODEL_SCHEMA) === null, 'Scenario sanitizer accepted a non-object root.');

for (const duration of [2, 4, 6, 8]) {
  const candidate = clone(cleanModel({ battery: { durationH: duration } }, MODEL_SCHEMA));
  check(candidate.battery.durationH === duration, `Scenario sanitizer rejected supported ${duration} h duration.`);
}

// Exercise the actual battery calculator with a fractional tradeable duration and an oversized unit.
const batteryModel = clone(defaults);
batteryModel.battery.durationH = 2;
batteryModel.battery.socFloor = 0.10;
batteryModel.battery.cyclesDay = 1;
batteryModel.battery.powerMW = 500;
const batteryContext = {
  M: batteryModel,
  PRICES: { per_year: { '2025': { ph: [] } } },
  SPREAD: { byDurYear: { '2': [{ y: '2025', buy: 10, sell: 30 }] } },
  spreadDur: () => 2,
  supplyStats: () => ({ battIn: 0 }),
};
vm.createContext(batteryContext);
vm.runInContext(
  `${section(app, 'function computeBattery(b){', '\nlet measYear=', 'battery calculator')}\n` +
  'globalThis.auditBattery=computeBattery;',
  batteryContext,
);
const battery = clone(batteryContext.auditBattery(batteryModel.battery));
near(battery.nCh, 1.8, 1e-12, 'fractional battery dispatch duration');
check(!Number.isInteger(battery.nCh), 'Battery dispatch duration was rounded to a whole hour.');
near(battery.ancMW, 225, 1e-12, 'ancillary-market capacity cap');
near(battery.ancRev, batteryModel.battery.ancPerMW * 225 / 1000, 1e-12, 'ancillary revenue at capped capacity');

// Run the real hourly supply loop on an all-deficit year. Any free opening state would serve load.
const emptyProfile = Array(8760).fill(0);
const supplyModel = clone(defaults);
supplyModel.wind.mw = 0;
supplyModel.solar.mw = 0;
supplyModel.dc.firmMW = 1;
supplyModel.battery.on = true;
const supplyContext = {
  M: supplyModel,
  PUBLIC_PROFILES: { w: emptyProfile, s: emptyProfile, window: 'regression fixture' },
  DAYS: [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31],
  measYear: 'fixture',
  _SUPC: null,
  typicalYear: () => ({ w: emptyProfile, s: emptyProfile }),
  measSeries: () => ({ meta: { window: 'regression fixture' } }),
  avg: (values) => values.reduce((sum, value) => sum + value, 0) / values.length,
  effCF: (asset) => asset.grossCF * (1 - asset.loss) * (1 - asset.lineLoss),
};
vm.createContext(supplyContext);
vm.runInContext(
  `${section(app, 'function supplyStats(){', '/* ============ 5 · FORMATTING', 'hourly supply model')}\n` +
  'globalThis.auditSupply=supplyStats;',
  supplyContext,
);
const supply = clone(supplyContext.auditSupply());
near(supply.fromBatt, 0, 1e-12, 'battery delivery from an uncharged opening state');
near(supply.endSoc, 0, 1e-12, 'ending state of charge for an all-deficit year');
near(supply.gridE, 8760, 1e-9, 'hourly grid shortfall for an all-deficit year');

// Exercise the SPV row builder with a deliberately non-zero hourly shortfall even though annual
// renewable generation exceeds annual load. This catches a regression to annual load-minus-output.
const spvModel = clone(defaults);
spvModel.dc.firmMW = 100;
spvModel.battery.on = false;
const hourlyGridMWh = 123456;
const years = Array.from({ length: 2070 - 2026 + 1 }, (_, index) => 2026 + index);
const asset = { totalCapex: 0, lcoe: 50, prod: 1_000_000, rows: years.map((y) => ({ y, prod: 1_000_000 })) };
const spvContext = {
  M: spvModel,
  Y0: 2026,
  YN: 2070,
  COD: 2028,
  FF: 2029,
  CAP7: { wind: 50, solar: 50 },
  computeAsset: () => asset,
  computeBattery: () => ({ capex: 0, arbRev: 0, ancRev: 0, capRev: 0, opex: 0, gridFee: 0 }),
  supplyStats: () => ({ gridE: hourlyGridMWh }),
  lineMW: () => 0,
  resPrice: () => 100,
  resCapFeeM: () => 0,
  feeF: () => 1,
  annPay: () => 0,
  xirr: () => 0,
};
vm.createContext(spvContext);
vm.runInContext(
  `${section(app, 'function computeSPV(dcP){', '\nfunction solveDcFor(', 'SPV calculator')}\n` +
  'globalThis.auditSPV=computeSPV;',
  spvContext,
);
const spv = clone(spvContext.auditSPV(spvModel.dc.dcPrice));
const firstFullYear = spv.rows.find((row) => row.y === 2029);
check(firstFullYear, 'SPV calculator did not produce the first full operating year.');
near(firstFullYear.gridMWh, hourlyGridMWh, 1e-9, 'SPV hourly grid-energy plumbing');
check(hourlyGridMWh > Math.max(0, spvModel.dc.firmMW * 8760 - 2_000_000), 'Regression fixture does not distinguish hourly and annual shortfall.');
check(/const\s+gridMWh\s*=\s*Math\.max\(0,y1b\.gridMWh\|\|0\)/.test(app), 'Data-center bill no longer reads grid energy from the SPV row.');
check(/const\s+dcLoad=M\.dc\.firmMW\*8760,\s*gridMWh=Math\.max\(0,yr\.gridMWh\|\|0\)/.test(app), 'SPV waterfall no longer reads grid energy from the selected SPV row.');

// The workbook remains in the repository for future reconciliation work, but it must have no UI entry point.
check(!/\bdlXLSX\b/.test(html), 'Excel export is still exposed in the public HTML.');
const visibleHtml = html.replace(/<!--[\s\S]*?-->/g, '');
const buttons = visibleHtml.match(/<button\b[\s\S]*?<\/button>/gi) || [];
check(!buttons.some((button) => /\b(?:Excel|XLSX)\b/i.test(button)), 'Excel export button remains in the public UI.');
check(!/<script\b[^>]*\bsrc\s*=\s*["'][^"']*model_export\.js/i.test(html), 'Workbook builder is still loaded by the public page.');

console.log('Model integrity regression checks passed (gate, sanitizer, battery, hourly supply/SPV, UI export).');
