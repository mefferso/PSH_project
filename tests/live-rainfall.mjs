import { readFile } from "node:fs/promises";
import vm from "node:vm";
import assert from "node:assert/strict";

async function loadRainCore() {
  const source = await readFile("src/core/rainfall_core.js", "utf8");
  const context = vm.createContext({});
  vm.runInContext(source + "\nglobalThis.__PSHRainCore = PSHRainCore;", context);
  return context.__PSHRainCore;
}

const core = await loadRainCore();

const iemDaily = new URL("https://mesonet.agron.iastate.edu/cgi-bin/request/daily.py");
for (const [key, value] of Object.entries({
  network: "LA_COCORAHS",
  year1: "2024", month1: "9", day1: "11",
  year2: "2024", month2: "9", day2: "12",
  var: "precip_in", format: "csv", na: "blank"
})) iemDaily.searchParams.set(key, value);

const dailyResponse = await fetch(iemDaily, { headers: { "User-Agent": "PSH-project-live-integration" } });
assert.equal(dailyResponse.ok, true, "IEM daily endpoint HTTP " + dailyResponse.status);
const dailyText = await dailyResponse.text();
const totals = core.parseIemDailyCsv(dailyText);
assert.equal(typeof totals["LA-JF-20"], "number", "IEM daily response missing LA-JF-20");
assert(Math.abs(totals["LA-JF-20"] - 9.48) <= 0.05,
  "IEM LA-JF-20 historical total changed: " + totals["LA-JF-20"]);
console.log("PASS live IEM daily LA-JF-20: " + totals["LA-JF-20"].toFixed(2) + " in");

const catalogResponse = await fetch("https://mesonet.agron.iastate.edu/geojson/network.php?network=LA_COCORAHS",
  { headers: { "User-Agent": "PSH-project-live-integration" } });
assert.equal(catalogResponse.ok, true, "IEM network endpoint HTTP " + catalogResponse.status);
const catalog = await catalogResponse.json();
assert((catalog.features || []).some(f => String(f.id || "").toUpperCase() === "LA-JF-20"),
  "IEM LA_COCORAHS catalog missing LA-JF-20");
console.log("PASS live IEM catalog contains LA-JF-20");

const token = process.env.SYNOPTIC_TOKEN;
if (!token) {
  console.log("SKIP live Synoptic regression: SYNOPTIC_TOKEN secret is not configured.");
} else {
  const url = new URL("https://api.synopticdata.com/v2/stations/precip");
  for (const [key, value] of Object.entries({
    token,
    bbox: "-91.8,28.7,-88.0,31.6",
    start: "202409101200",
    end: "202409121200",
    pmode: "totals",
    units: "english,precip|in",
    obtimezone: "UTC",
    all_reports: "0",
    complete: "1"
  })) url.searchParams.set(key, value);
  const response = await fetch(url, { headers: { "User-Agent": "PSH-project-live-integration" } });
  assert.equal(response.ok, true, "Synoptic precip endpoint HTTP " + response.status);
  const json = await response.json();
  const station = (json.STATION || []).find(st =>
    core.aliasesOverlap(String(st.STID || ""), core.normalizeRainNetwork(st.MNET_SHORTNAME || st.SOURCE || st.MNET_ID || ""),
      "LIX", "COOP")
  );
  assert(station, "Synoptic bbox response missing an exact LIX/COOPLIX alias");
  const total = core.bestPrecipTotal(station.OBSERVATIONS || {});
  assert.equal(typeof total, "number", "Synoptic LIX alias has no usable precip total");
  assert(Math.abs(total - 7.93) <= 0.25, "Synoptic LIX total changed: " + total);
  console.log("PASS live Synoptic LIX: " + total.toFixed(2) + " in");
}
