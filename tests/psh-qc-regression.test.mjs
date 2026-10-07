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
