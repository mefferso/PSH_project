import assert from "node:assert/strict";
import {readFile} from "node:fs/promises";
import vm from "node:vm";
const src=await readFile("src/PSH_Automation.gs","utf8");
const ctx=vm.createContext({Date,console});
vm.runInContext(src,ctx);
function fakeSheet(data) {
  const sh={
    data,
    getRange(row,col,n=1){ return {
      getValues:()=>Array.from({length:n},(_,i)=>[sh.data[row+i-1]||""]),
      setValues:vals=>vals.forEach((v,i)=>sh.data[row+i-1]=v[0])
    }; }
  };return sh;
}
const sh=fakeSheet(["human note | [AUTO-QC] E","E","[AUTO-QC] previous","human | [AUTO-QC] old"]);
ctx.clearAutoQc_(sh,1,4,14,14);
assert.deepEqual(sh.data,["human note","E","","human"]);
ctx.mergeQcColumn_(sh,1,14,["[AUTO-QC] E","","[AUTO-QC] new",""]);
assert.deepEqual(sh.data,["human note | [AUTO-QC] E","E","[AUTO-QC] new","human"]);
ctx.mergeQcColumn_(sh,1,14,["","","",""]);
assert.deepEqual(sh.data,["human note","E","","human"]);
console.log("PASS: tagged QC is transient; human E/comments preserved.");
assert.equal(src.includes("Synoptic COOP alias fallback unavailable"),false);
assert.equal(src.includes("Recovered COOP rainfall via Synoptic alias"),false);
assert.equal(src.includes("Review equal sustained wind and gust >=25 kt"),true);
console.log("PASS: noisy COOP prefix retry removed; airport review guard present.");

assert.equal(ctx.parseUtc_("10/08/2026 18:30").toISOString(),"2026-10-08T18:30:00.000Z");
assert.equal(ctx.parseUtc_("02/30/2026 12:00"),null);
assert.equal(ctx.parseUtc_("13/01/2026 12:00"),null);
assert.equal(ctx.parseUtc_("10/08/2026 25:00"),null);
assert.equal(ctx.parseUtc_("2026-10-08 18:30").toISOString(),"2026-10-08T18:30:00.000Z");
assert.match(src,/Math\.round\(p\.wind\.value\)/);
assert.match(src,/Math\.round\(p\.gust\.value\)/);
console.log("PASS: unambiguous US UTC date/time parsing and whole-knot sheet output.");

assert.match(src,/function pshAuditCwmsMappings\(/);
assert.match(src,/NO EXACT TSID MATCH; requires agency mapping/);
assert.match(src,/No PSH observations were changed/);
assert.equal(src.includes("Audit USACE/CPRA Mappings (no data changes)"),true);
console.log("PASS: CWMS discovery is read-only, exact-match and visible in menu.");

assert.match(src,/function pshAuditHistoricalWater\(/);
assert.match(src,/Audit All Missing Water Archives/);
assert.match(src,/REVIEW ONLY: stage datum not validated/);
assert.match(src,/Stages are NOT converted to NAVD88/);
console.log("PASS: all-station HML archive audit is read-only and datum gated.");

const hmlStart=new Date("2026-10-08T00:00:00Z"),hmlEnd=new Date("2026-10-09T00:00:00Z");
const hmlCsv=[
  ["station","valid[utc]","ceiling height[ft]","stage[ft]","tide height[ft]"],
  ["BSGL1","2026-10-08 20:00","100","1.77","9.99"],
  ["BSGL1","2026-10-08 21:00","","1.68",""],
  ["BSGL1","2026-10-07 23:00","","99.00",""],
  ["LBWL1","2026-10-08 22:00","","2.54",""]
];
const hml=ctx.parseIemHmlStageCsv_(hmlCsv,["BSGL1","LBWL1"],hmlStart,hmlEnd);
assert.equal(hml.error,"");
assert.equal(hml.maxById.BSGL1.value,1.77);
assert.equal(hml.maxById.BSGL1.time.toISOString(),"2026-10-08T20:00:00.000Z");
assert.equal(hml.maxById.LBWL1.value,2.54);
assert.equal(ctx.parseIemHmlStageCsv_([["station","valid[utc]","lake elev abv datum[ft]"],["BSGL1","2026-10-08 20:00","5"]],["BSGL1"],hmlStart,hmlEnd).error.startsWith("No explicit stage"),true);
console.log("PASS: IEM actual CSV stage[ft] extraction; do not substitute tide/other elevations or outside-window values.");
