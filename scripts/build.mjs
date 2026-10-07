import { readFile, mkdir, writeFile } from "node:fs/promises";

const core = await readFile("src/core/rainfall_core.js", "utf8");
const app = await readFile("src/PSH_Automation.gs", "utf8");
const output = core.trimEnd() + "\n\n" + app.trimStart();

await mkdir("dist", { recursive: true });
await writeFile("dist/PSH_Automation.gs", output, "utf8");
console.log("Built dist/PSH_Automation.gs from rainfall core + Apps Script adapter");
