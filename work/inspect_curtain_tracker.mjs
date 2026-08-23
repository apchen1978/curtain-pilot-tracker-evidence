import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const input = path.resolve(__dirname, "..", "outputs", "curtain-soft-furnishing-pilot-tracker-pilot-001-simulation.xlsx");
const workDir = __dirname;
const file = await FileBlob.load(input);
const wb = await SpreadsheetFile.importXlsx(file);
const summary = await wb.inspect({ kind: "workbook,sheet,table,formula", maxChars: 12000, tableMaxRows: 16, tableMaxCols: 24 });
console.log(summary.ndjson);
for (const sheetName of ["Leads", "Follow-ups", "Quotes", "Dashboard"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${workDir}/before-${sheetName}.png`, new Uint8Array(await preview.arrayBuffer()));
}
