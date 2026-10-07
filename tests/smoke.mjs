import { readFile, writeFile, unlink } from "node:fs/promises";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";

const core = await readFile("src/core/rainfall_core.js", "utf8");
const src = await readFile("src/PSH_Automation.gs", "utf8");
const dist = await readFile("dist/PSH_Automation.gs", "utf8");
const expectedDist = core.trimEnd() + "\n\n" + src.trimStart();

function fail(msg) {
  console.error("FAIL:", msg);
  process.exitCode = 1;
}

if (expectedDist !== dist) {
  fail("dist/PSH_Automation.gs is stale. Run npm run build.");
}

const required = [
  "function onOpen()",
  "function pshRunAll()",
  "function pshRunWindPressure()",
  "function pshRunRainfall()",
  "function pshRunWaterLevels()",
  "function pshRunFrancineRegression()",
  "function fetchWeatherStemWind_",
  "function fetchCocorahsTotal_",
  "function fetchUsgsWater_"
];
for (const needle of required) {
  if (!src.includes(needle)) fail("Required Apps Script entry point/helper missing: " + needle);
}

const coreRequired = [
  "const PSHRainCore",
  "function canonicalCocorahsId",
  "function rainIdAliasKeys",
  "function buildSynopticRowIndex",
  "function matchSynopticRows",
  "function parseIemDailyCsv",
  "function resolveRainfallSources"
];
for (const needle of coreRequired) {
  if (!core.includes(needle)) fail("Required rainfall core helper missing: " + needle);
}

if (!src.includes("v0.15 - executable rainfall core + Francine repo regression")) {
  fail("Expected v0.15 baseline header was not found.");
}

const adapterGuards = [
  "return PSHRainCore.synopticRequestId(rawId, network);",
  "return PSHRainCore.canonicalCocorahsId(id);",
  "return PSHRainCore.rainIdAliasKeys(id, network);",
  "return PSHRainCore.normalizeRainNetwork(x);",
  "PSHRainCore.buildSynopticRowIndex(rowMeta)",
  "PSHRainCore.matchSynopticRows(st, byApiId)",
  "PSHRainCore.parseIemDailyCsv(text)",
  "return PSHRainCore.bestPrecipTotal(obs);"
];
for (const needle of adapterGuards) {
  if (!src.includes(needle)) fail("Apps Script adapter is not using tested rainfall core: " + needle);
}

if (src.includes("WSEBRAlexBox',col:17,expectedBlank:true")) {
  fail("Stale Alex Box deliberate-blank WeatherSTEM regression assertion is still present.");
}

const francine = JSON.parse(await readFile("tests/francine_expected.json", "utf8"));
for (const id of [
  "WSEBRAlexBox","WSEBRTigerStadium","WSNOLakefront",
  "WSNOMidCIty","WSNOBayouSauvage","WSSCEOC","WSSCLuling"
]) {
  const ref = francine.wind_pressure?.[id];
  if (!ref || typeof ref.sustained_kt !== "number" || typeof ref.gust_kt !== "number") {
    fail("Missing issued WeatherSTEM Francine reference for " + id);
  }
}

const combined = core + "\n" + src;
for (const forbidden of [/AIza[0-9A-Za-z_-]{20,}/, /ghp_[0-9A-Za-z]{20,}/, /github_pat_[0-9A-Za-z_]{20,}/]) {
  if (forbidden.test(combined)) fail("Possible credential string found in source.");
}

const tmp = ".psh-syntax-check.js";
await writeFile(tmp, dist, "utf8");
try {
  execFileSync(process.execPath, ["--check", tmp], { stdio: "inherit" });
} catch {
  fail("Built Apps Script JavaScript syntax check failed.");
} finally {
  await unlink(tmp).catch(() => {});
}

const hash = createHash("sha256").update(dist).digest("hex");
console.log("PASS: PSH Apps Script smoke checks. dist sha256=" + hash);
