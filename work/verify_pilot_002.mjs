import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const input = path.resolve(__dirname, "..", "outputs", "curtain-soft-furnishing-pilot-tracker-pilot-002-simulation.xlsx");
const file = await FileBlob.load(input);
const wb = await SpreadsheetFile.importXlsx(file);
for (const [sheetName, range] of [
  ["Leads", "A3:V7"],
  ["Follow-ups", "A3:I8"],
  ["Pilot-002 Timeline", "A3:I15"],
  ["Dashboard", "A19:B20"],
]) {
  const check = await wb.inspect({ kind: "table", range: `${sheetName}!${range}`, include: "values,formulas", tableMaxRows: 20, tableMaxCols: 24 });
  console.log(check.ndjson);
}
const quoteCheck = await wb.inspect({ kind: "match", searchTerm: "SIM-PILOT-002", options: { useRegex: false, maxResults: 50 }, summary: "Pilot #002 cross-sheet reference check" });
console.log(quoteCheck.ndjson);
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
console.log(errors.ndjson);
