import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const input = path.resolve(__dirname, "..", "outputs", "curtain-soft-furnishing-pilot-tracker-pilot-002-simulation.xlsx");
const outputDir = __dirname;
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(input));

for (const range of [
  "Leads!A1:V7",
  "Follow-ups!A1:I8",
  "Quotes!A1:K6",
  "Dashboard!A1:H22",
  "Data Model Issues - Backlog!A1:F11",
]) {
  const result = await wb.inspect({ kind: "table,computedStyle", range, include: "values,formulas", tableMaxRows: 24, tableMaxCols: 24, maxChars: 8000 });
  console.log(`--- ${range} ---\n${result.ndjson}`);
}

for (const sheetName of ["Quotes", "Dashboard", "Data Model Issues - Backlog"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/before-pilot-003-${sheetName}.png`, new Uint8Array(await preview.arrayBuffer()));
}
