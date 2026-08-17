import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const dataPath = path.join(repo, 'gdc_data.js');
const appPath = path.join(repo, 'gdc_app.js');
const privateConstants = new Set(['BEDATA', 'BENCH', 'NEWD', 'MEAS', 'CLIP']);
const days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

function planningProfile() {
  const wind = [];
  const solar = [];
  const months = [];
  let month = 0;
  let monthStart = 0;
  const clamp = (value, low, high) => Math.max(low, Math.min(high, value));

  for (let day = 0; day < 365; day += 1) {
    while (day >= monthStart + days[month]) {
      monthStart += days[month];
      month += 1;
    }
    const season = Math.cos((2 * Math.PI * (day - 15)) / 365);
    const daylight = 12 + 4.15 * Math.sin((2 * Math.PI * (day - 80)) / 365);
    const sunrise = 12 - daylight / 2;
    const synoptic =
      0.095 * Math.sin((2 * Math.PI * day) / 8.7) +
      0.065 * Math.sin((2 * Math.PI * (day + 17)) / 21.3) +
      0.035 * Math.cos((2 * Math.PI * day) / 3.9);
    const cloud = clamp(
      0.78 +
        0.13 * Math.sin((2 * Math.PI * (day + 9)) / 11.7) +
        0.08 * Math.cos((2 * Math.PI * day) / 27.1),
      0.34,
      0.98,
    );

    for (let hour = 0; hour < 24; hour += 1) {
      const sunPhase = (hour + 0.5 - sunrise) / daylight;
      solar.push(
        sunPhase > 0 && sunPhase < 1
          ? Math.sin(Math.PI * sunPhase) ** 1.36 * cloud
          : 0,
      );
      wind.push(
        clamp(
          0.255 +
            0.055 * season +
            synoptic +
            0.025 * Math.cos((2 * Math.PI * (hour - 3)) / 24) +
            0.018 * Math.sin((2 * Math.PI * (day * 24 + hour)) / 61),
          0.015,
          0.92,
        ),
      );
      months.push(month);
    }
  }
  return { wind, solar, months };
}

function groupedMean(values, groups, count) {
  const sums = Array(count).fill(0);
  const samples = Array(count).fill(0);
  values.forEach((value, index) => {
    const group = groups[index];
    sums[group] += value;
    samples[group] += 1;
  });
  return sums.map((sum, index) => +(sum / samples[index]).toFixed(4));
}

function planningCapture(row, profile) {
  const hourlyMean = row.ph.reduce((sum, value) => sum + value, 0) / row.ph.length;
  const monthlyPrices = row.pm.map((value) =>
    Number.isFinite(value) ? value : row.baseload,
  );
  const monthlyMean = monthlyPrices.reduce(
    (sum, value, month) => sum + value * days[month],
    0,
  ) / 365;
  let weighted = 0;
  let production = 0;

  profile.forEach((value, index) => {
    const hour = index % 24;
    const month = publicProfile.months[index];
    const syntheticPrice =
      row.baseload +
      (row.ph[hour] - hourlyMean) +
      (monthlyPrices[month] - monthlyMean);
    weighted += syntheticPrice * value;
    production += value;
  });
  return +(weighted / production).toFixed(1);
}

let dataSource = fs.readFileSync(dataPath, 'utf8');
dataSource = dataSource.replace(/^const BENCH=\{[\s\S]*?^\};\r?\n?/m, '');
let dataLines = dataSource.split(/\r?\n/);
const orphanBenchStart = dataLines.findIndex((line) =>
  /^\s*wind:\{label:'WP Nickelsdorf'/.test(line),
);
if (orphanBenchStart >= 0) {
  const orphanBenchEnd = dataLines.findIndex(
    (line, index) => index >= orphanBenchStart && line.trim().endsWith('}}};'),
  );
  if (orphanBenchEnd < 0) throw new Error('Incomplete benchmark block found');
  dataLines.splice(orphanBenchStart, orphanBenchEnd - orphanBenchStart + 1);
}
const pricesIndex = dataLines.findIndex((line) => line.startsWith('const PRICES='));
if (pricesIndex < 0) throw new Error('PRICES dataset not found');

const prices = JSON.parse(dataLines[pricesIndex].slice('const PRICES='.length, -1));
const publicProfile = planningProfile();
const hourGroups = publicProfile.wind.map((_, index) => index % 24);

prices.windHr = groupedMean(publicProfile.wind, hourGroups, 24);
prices.solarHr = groupedMean(publicProfile.solar, hourGroups, 24);
prices.windMo = groupedMean(publicProfile.wind, publicProfile.months, 12);
prices.solarMo = groupedMean(publicProfile.solar, publicProfile.months, 12);
prices.src.shape = 'Illustrative deterministic planning profile; not project metering';

Object.values(prices.per_year).forEach((row) => {
  row.wind = planningCapture(row, publicProfile.wind);
  row.solar = planningCapture(row, publicProfile.solar);
  row.basis = 'planning shape';
  delete row.meteredShare;
});

dataLines[pricesIndex] = `const PRICES=${JSON.stringify(prices)};`;
dataLines = dataLines.filter((line) => {
  const match = line.match(/^const\s+([A-Z][A-Z0-9_]*)=/);
  return !match || !privateConstants.has(match[1]);
});
while (dataLines.at(-1) === '') dataLines.pop();
fs.writeFileSync(dataPath, `${dataLines.join('\r\n')}\r\n`, 'utf8');

let appSource = fs.readFileSync(appPath, 'utf8');
const rawStart = appSource.indexOf(
  '/* ---------- raw 15-minute data explorer (Production > Raw data) ----------',
);
const rawEnd = appSource.indexOf('/* ============ 16 · MOTION UPGRADES', rawStart);
if (rawStart >= 0 && rawEnd >= 0) {
  appSource = `${appSource.slice(0, rawStart)}${appSource.slice(rawEnd)}`;
}
const meterNarrativeStart = appSource.indexOf(
  '/* ---- Is the 1.20 DC to AC ratio an assumption or a measurement?',
);
const termsNarrativeStart = appSource.indexOf(
  '/* ---- Commercial terms, sealed',
  meterNarrativeStart,
);
if (meterNarrativeStart >= 0 && termsNarrativeStart >= 0) {
  appSource = `${appSource.slice(0, meterNarrativeStart)}/* ---- The 500 MW wall ---------------------------------------------------------------------
   Headroom is tested against the explicit public planning profile. It is a sizing illustration,
   not a production forecast or a claim based on project meters.
-------------------------------------------------------------------------------------------- */

${appSource.slice(termsNarrativeStart)}`;
}
fs.writeFileSync(appPath, appSource, 'utf8');

console.log('Public release datasets sanitized.');
