import { readFile, writeFile, unlink } from "node:fs/promises";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";

const src = await readFile("src/PSH_Automation.gs", "utf8");
const dist = await readFile("dist/PSH_Automation.gs", "utf8");

function fail(msg) {
  console.error("FAIL:", msg);
  process.exitCode = 1;
}

if (src !== dist) fail("src/PSH_Automation.gs and dist/PSH_Automation.gs differ. Run npm run build.");

const required = [
  "function onOpen()",
  "function pshRunAll()",
  "function pshRunWindPressure()",
  "function pshRunRainfall()",
  "function pshRunWaterLevels()",
  "function pshRunFrancineRegression()",
  "function fetchWeatherStemWind_",
  "function weatherStemMaxAnemometer_",
  "function fetchCocorahsTotal_",
  "function fetchUsgsWater_"
];
for (const needle of required) {
  if (!src.includes(needle)) fail(`Required entry point/helper missing: ${needle}`);
}

if (!src.includes("v0.11 - Rainfall source/ID hardening + WeatherSTEM v0.10 validation")) {
  fail("Expected v0.11 rainfall-hardening baseline header was not found.");
}

const rainfallGuards = [
  "function canonicalCocorahsId_",
  "net === 'COOP' && id && !/^COOP/.test(id)",
  "const maxHistoricalChecks=200;",
  "ACIS as a conservative fallback only",
  "findRainRowByAlias_(sh, c.id, 'COCORAHS')"
];
for (const needle of rainfallGuards) {
  if (!src.includes(needle)) fail(`Rainfall hardening guard missing: ${needle}`);
}

if (src.includes("WSEBRAlexBox',col:17,expectedBlank:true")) {
  fail("Stale Alex Box deliberate-blank WeatherSTEM regression assertion is still present.");
}

const francine = JSON.parse(await readFile("tests/francine_expected.json", "utf8"));
const weatherStemPriority = [
  "WSEBRAlexBox",
  "WSEBRTigerStadium",
  "WSNOLakefront",
  "WSNOMidCIty",
  "WSNOBayouSauvage",
  "WSSCEOC",
  "WSSCLuling"
];
for (const id of weatherStemPriority) {
  const ref = francine.wind_pressure?.[id];
  if (!ref || typeof ref.sustained_kt !== "number" || typeof ref.gust_kt !== "number") {
    fail(`Missing issued WeatherSTEM Francine reference for ${id}.`);
  }
  if (!src.includes(`id:'${id}',col:11,expected:${ref.sustained_kt},tol:1`)) {
    fail(`Source regression is missing WeatherSTEM sustained target for ${id}.`);
  }
  if (!src.includes(`id:'${id}',col:17,expected:${ref.gust_kt},tol:1`)) {
    fail(`Source regression is missing WeatherSTEM gust target for ${id}.`);
  }
}

for (const forbidden of [/AIza[0-9A-Za-z_-]{20,}/, /ghp_[0-9A-Za-z]{20,}/, /github_pat_[0-9A-Za-z_]{20,}/]) {
  if (forbidden.test(src)) fail("Possible credential string found in source.");
}

const tmp = ".psh-syntax-check.js";
await writeFile(tmp, src, "utf8");
try {
  execFileSync(process.execPath, ["--check", tmp], { stdio: "inherit" });
} catch {
  fail("JavaScript syntax check failed.");
} finally {
  await unlink(tmp).catch(() => {});
}

const hash = createHash("sha256").update(src).digest("hex");
console.log(`PASS: PSH Apps Script smoke checks. src sha256=${hash}`);
