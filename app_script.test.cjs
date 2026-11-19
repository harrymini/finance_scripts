const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const code = fs.readFileSync(__dirname + '/app_script.js', 'utf8');
const csv = 'observation_date,DTWEXBGS\r\n2026-11-13,122.6503\r\n2026-11-11,.\r\n2026-11-06,121.8803\r\n2026-11-12,122.5433\r\n2026-11-16,\r\n';
function app({ data = csv, status = 200 } = {}) {
  const store = new Map(), calls = [];
  const ctx = vm.createContext({
    Logger: { log() {} },
    CacheService: { getScriptCache() { return { get: k => store.get(k) ?? null, put: (k,v) => store.set(k,v) }; } },
    UrlFetchApp: { fetch(url) { calls.push(url); return { getResponseCode: () => status, getContentText: () => data }; } }
  });
  vm.runInContext(code, ctx);
  return { ctx, calls };
}
test('holiday uses a dated weekly observation, not two raw rows', () => {
  const {ctx} = app();
  const latest = ctx.getFredData('DTWEXBGS');
  const prior = ctx.getFredDataHistorical('DTWEXBGS', 7, latest.date);
  assert.equal(latest.date, '2026-11-13');
  assert.equal(prior.date, '2026-11-06');
  assert.equal(ctx.observationDelta(latest, prior).toFixed(2), '0.77');
  assert.equal(ctx.getPreviousRelease('DTWEXBGS').date, '2026-11-12');
});
test('missing HTTP responses abstain and preserve valid zero', () => {
  const {ctx} = app({status: 503});
  const current = ctx.getFredData('DTWEXBGS');
  assert.equal(current.value, null);
  assert.equal(ctx.observationDelta(current, {value: 100}), null);
  assert.equal(ctx.getContextualDXYScore(null).score, 0);
  assert.equal(ctx.determineChinaSignal(null, 2).score, 0);
  assert.equal(ctx.observationDelta({value:0}, {value:0}), 0);
  assert.equal(ctx.formatObservation(null), '데이터 없음');
});
test('zero CSV observations survive cache; invalid dates and values are rejected', () => {
  const {ctx,calls} = app({data:'date,value\n2026-11-06,0\n2026-11-07,12garbage\n2026-11-31,2\n2026-11-08,Infinity'});
  assert.equal(ctx.getFredData('X').value, 0);
  assert.equal(ctx.getFredData('X').value, 0);
  assert.equal(calls.length, 1);
  assert.equal(ctx.getFredDataHistorical('X', 7).value, null);
});
test('dollar regime only uses finite comparisons; SOFR config is a rate', () => {
  const {ctx} = app();
  ctx.getRateCycleStatus = () => ({cycle:'HOLDING'});
  ctx.getQTStatus = () => ({status:'QT'});
  assert.notEqual(ctx.detectMarketRegime().regime, 'DXY_VOLATILE');
  ctx.getFredData = () => ({value:122.65, date:'2026-11-13'});
  ctx.getFredDataHistorical = () => ({value:null,date:null});
  assert.notEqual(ctx.detectMarketRegime().regime, 'DXY_VOLATILE');
  const series = [];
  ctx.getFredData = id => {series.push(id); return {value: id === 'SOFR' ? 3.69 : 3.65};};
  assert.equal(ctx.getRepoSpreadStatus().spread.toFixed(2), '4.00');
  assert.deepEqual(series, ['SOFR', 'IORB']);
  ctx.getFredData = id => ({value:id === 'SOFR' ? null : 3.65});
  assert.equal(ctx.getRepoSpreadStatus().status, 'UNKNOWN');
});
