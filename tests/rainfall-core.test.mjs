import { readFile } from "node:fs/promises";
import vm from "node:vm";
import assert from "node:assert/strict";

async function loadRainCore() {
  const source = await readFile("src/core/rainfall_core.js", "utf8");
  const context = vm.createContext({});
  vm.runInContext(source + "\nglobalThis.__PSHRainCore = PSHRainCore;", context, {
    filename: "src/core/rainfall_core.js"
  });
  return context.__PSHRainCore;
}

const core = await loadRainCore();
const fixture = JSON.parse(await readFile("fixtures/api/francine/rainfall_sources.json", "utf8"));
const expected = JSON.parse(await readFile("tests/francine_expected.json", "utf8"));
const iemCsv = await readFile("fixtures/api/francine/iem_cocorahs_daily.csv", "utf8");

assert.equal(core.canonicalCocorahsId("LA-SC-06"), "LA-SC-6");
assert(core.aliasesOverlap("LIX", "COOP", "COOPLIX", "COOP"));
assert(core.aliasesOverlap("MSY", "ASOS", "KMSY", "ASOS"));
assert.equal(core.synopticRequestId("LIX", "COOP"), "LIX");
assert.equal(core.synopticRequestId("MSY", "ASOS"), "KMSY");

const parsedIem = core.parseIemDailyCsv(iemCsv);
assert(Math.abs(parsedIem["LA-SC-6"] - 9.10) <= 1e-9);
assert.equal(parsedIem["LA-ST-999"], 0.15, "IEM trace sentinel should count as zero");

const index = core.buildSynopticRowIndex(
  fixture.rows.map((row, rowNum) => ({ ...row, rowNum }))
);
const lixRow = fixture.rows.findIndex(row => row.id === "LIX");
const msyRow = fixture.rows.findIndex(row => row.id === "MSY");
assert(index.byAlias.LIX.includes(lixRow));
assert(index.byAlias.COOPLIX.includes(lixRow));
assert(index.byAlias.KMSY.includes(msyRow));

const resolved = core.resolveRainfallSources({
  rows: fixture.rows,
  synopticDirect: fixture.synoptic_direct,
  synopticBulk: fixture.synoptic_bulk,
  cocorahsOfficial: fixture.cocorahas_official || fixture.cocorahs_official,
  iemCocorahs: parsedIem,
  acis: fixture.acis
});

for (const row of fixture.rows) {
  const id = row.id;
  const target = expected.rainfall_in[id];
  assert.equal(typeof target, "number", "Missing Francine rainfall reference for " + id);
  const actual = resolved.byId[id];
  assert(actual, "No deterministic rainfall result for " + id);
  assert(Math.abs(actual.value - target) <= 0.001, id + ": " + actual.value + " != issued " + target);
  console.log("PASS rainfall " + id + ": " + actual.value.toFixed(2) + " in via " + actual.source);
}

assert.equal(resolved.byId["LA-SC-06"].source, "CoCoRaHS",
  "Official CoCoRaHS must outrank the IEM mirror");
assert.equal(resolved.byId.LIX.source, "Synoptic",
  "Exact-window Synoptic must outrank the intentionally wrong ACIS LIX fixture");
assert.equal(resolved.byId.LIX.value, 7.93,
  "Francine LIX must not regress to the 4.33-inch ACIS calendar-bin value");

assert.equal(expected.rainfall_in["LA-JF-20"], 9.48);
console.log("KNOWN MANUAL Francine rainfall LA-JF-20: issued 9.48 in; current live sources do not defensibly automate it.");
console.log("PASS: deterministic Francine rainfall core regression.");
