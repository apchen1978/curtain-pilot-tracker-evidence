import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const input = path.resolve(__dirname, "..", "outputs", "curtain-soft-furnishing-pilot-tracker-pilot-003-quote-version-simulation.xlsx");
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
const checks = [
  ["Leads", "A3:V8"],
  ["Follow-ups", "A3:I9"],
  ["Quotes", "A3:K7"],
  ["Quote Version Log", "A3:L5"],
  ["Pilot-003 Timeline", "A3:I11"],
  ["Dashboard", "A3:H27"],
  ["Data Model Issues - Backlog", "A3:F12"],
];
for (const [sheet, range] of checks) {
  const result = await wb.inspect({ kind: "match", sheetId: sheet, searchTerm: "SIM-PILOT-003|SIM-Q-003", options: { useRegex: true, maxResults: 30 }, summary: `${sheet} links` });
  console.log(`${sheet}: ${result.ndjson}`);
}
const v2 = await wb.inspect({ kind: "table", range: "Quote Version Log!A3:L5", include: "values,formulas", tableMaxRows: 4, tableMaxCols: 12 });
const dashboard = await wb.inspect({ kind: "table", range: "Dashboard!A3:B10", include: "values,formulas", tableMaxRows: 10, tableMaxCols: 2 });
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "final formula scan" });
console.log(v2.ndjson);
console.log(dashboard.ndjson);
console.log(errors.ndjson);
