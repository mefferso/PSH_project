import { readFile } from "node:fs/promises";
import vm from "node:vm";
import assert from "node:assert/strict";

const src=await readFile("src/PSH_Automation.gs","utf8");
const ctx=vm.createContext({Date,console});
vm.runInContext(src,ctx,{filename:"PSH_Automation.gs"});
vm.runInContext(`
  Utilities = { parseCsv: text => text.trim().split(/\\r?\\n/).map(row => row.split(',')) };
  fetchTextResilient_ = () => [
    "station,valid,sknt,drct,gust,peak_wind_gust,peak_wind_drct,peak_wind_time,mslp",
    "HSA,2026-07-22 17:47,20,120,30,M,M,M,M",
    "HSA,2026-07-23 16:47,30,120,15,M,M,M,M",
    "HSA,2026-07-23 17:47,18,130,22,M,M,M,M"
  ].join("\\n");
`,ctx);
const start=new Date("2026-07-22T00:00:00Z");
const end=new Date("2026-07-23T23:59:00Z");
assert.equal(ctx.airportWindPairConsistent_(30,15,0.5),false);
assert.equal(ctx.airportWindPairConsistent_(20,30,0.5),true);
assert.equal(ctx.airportWindPairConsistent_(21,null,0.5),true);
const observed=ctx.fetchIemAirportWind_("KHSA",start,end);
assert.equal(observed.wind.value,20,"Inconsistent 30-kt sustained/15-kt gust must not become the event maximum");
assert.equal(observed.gust.value,30,"Valid peak gust should survive individual-report QC");
assert.equal(observed.wind.time.toISOString(),"2026-07-22T17:47:00.000Z");
console.log("PASS: internally inconsistent Bertha KHSA observation excluded; 20-kt sustained/30-kt gust retained.");
