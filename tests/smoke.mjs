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

if (!src.includes("v0.10 - WeatherSTEM maximum minute Anemometer sustained wind")) {
  fail("Expected v0.10 baseline header was not found.");
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
