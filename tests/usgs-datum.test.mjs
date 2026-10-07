import { readFile } from "node:fs/promises";
import vm from "node:vm";
import assert from "node:assert/strict";

const src = await readFile("src/PSH_Automation.gs", "utf8");
const ctx = vm.createContext({ Date, console });
vm.runInContext(src, ctx, { filename: "PSH_Automation.gs" });
vm.runInContext(`
  globalThis.__calls = [];
  globalThis.__logs = [];
  log_ = (...args) => __logs.push(args);
  fetchUsgsContinuousMax_ = (site, param) => {
    __calls.push({site, param});
    if (__mode === 'direct' && param === '63160') return {value:1.87,time:new Date('2026-07-23T17:00:00Z')};
    if (param === '00065') return {value:2.97,time:new Date('2026-07-22T09:00:00Z')};
    return null;
  };
  fetchUsgsSiteDatum_ = () => {
    throw Error('Current gage altitude must not be queried for automatic stage conversion.');
  };
  globalThis.__mode = 'stage';
`, ctx);
const start = new Date("2026-07-22T00:00:00Z");
const end = new Date("2026-07-23T23:59:00Z");

let result = ctx.fetchUsgsWater_("PSIL1", "https://waterdata.usgs.gov/monitoring-location/07374525/", "NAVD88", start, end);
assert.equal(result, null, "Stage-only historical site must remain blank without a validated event-effective datum");
assert.equal(ctx.__calls.some(x => x.param === "00065"), true, "Allowlisted station should query stage solely for diagnostic purposes");
assert(ctx.__logs.some(x => x[3].includes("event-effective gage-datum elevation")));
console.log("PASS: historical stage-only PSIL1 stays blank instead of publishing unverified NAVD88.");

ctx.__mode="direct";
ctx.__calls.length=0;
result=ctx.fetchUsgsWater_("BPPL1", "https://waterdata.usgs.gov/monitoring-location/07374525/", "NAVD88", start, end);
assert.equal(result.value, 1.87, "Genuine direct NAVD88 elevations must remain eligible");
assert.equal(result.comment, "", "Direct NAVD88 observation must not be flagged as estimated");
assert.equal(result.provenance, "direct-elevation-63160");
assert.equal(ctx.__calls.some(x => x.param === "00065"), false);
console.log("PASS: direct NAVD88 elevation is retained without an estimated conversion.");

assert.equal(ctx.usgsSiteId_("7380255",""),"07380255");
assert.equal(ctx.usgsSiteId_(7380212,""),"07380212");
assert.equal(ctx.usgsSiteId_("07380255",""),"07380255");
assert.equal(ctx.usgsSiteId_("BPPL1","https://waterdata.usgs.gov/monitoring-location/7380255/"),"07380255");
assert.equal(ctx.usgsSiteId_("BPPL1","https://waterdata.usgs.gov/monitoring-location/07380212/"),"07380212");
console.log("PASS: USGS IDs stripped of leading zeros by Sheets are restored.");
