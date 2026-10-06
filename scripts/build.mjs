import { copyFile, mkdir } from "node:fs/promises";

await mkdir("dist", { recursive: true });
await copyFile("src/PSH_Automation.gs", "dist/PSH_Automation.gs");
console.log("Built dist/PSH_Automation.gs");
