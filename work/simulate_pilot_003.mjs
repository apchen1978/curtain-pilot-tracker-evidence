import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const input = path.resolve(__dirname, "..", "outputs", "curtain-soft-furnishing-pilot-tracker-pilot-002-simulation.xlsx");
const output = path.resolve(__dirname, "..", "outputs", "curtain-soft-furnishing-pilot-tracker-pilot-003-quote-version-simulation.xlsx");
const workDir = __dirname;
const navy = "#18344A";
const teal = "#1F6F6D";
const sand = "#F7F2E8";
const rose = "#FDE2E2";
const lightBorder = "#D9E2E7";
const text = "#1F2933";
const pilotId = "SIM-PILOT-003";

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
const leads = wb.worksheets.getItem("Leads");
const followUps = wb.worksheets.getItem("Follow-ups");
const quotes = wb.worksheets.getItem("Quotes");
const dashboard = wb.worksheets.getItem("Dashboard");
const backlog = wb.worksheets.getItem("Data Model Issues - Backlog");

function title(sheet, range, value) {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[value]];
  sheet.getRange(range).format = { fill: navy, font: { bold: true, color: "#FFFFFF", size: 16 }, verticalAlignment: "center" };
  sheet.getRange(range).format.rowHeight = 30;
}
function header(sheet, range) {
  sheet.getRange(range).format = { fill: teal, font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { preset: "all", style: "thin", color: lightBorder } };
}
function body(sheet, range) {
  sheet.getRange(range).format = { font: { color: text, size: 10 }, verticalAlignment: "center", wrapText: true, borders: { preset: "inside", style: "thin", color: lightBorder } };
}

// Current state remains intentionally open: a price objection is not a win or a loss.
leads.getRange("A8:V8").values = [[
  pilotId,
  new Date("2026-08-12T10:00:00"),
  "LINE（SIMULATION）",
  "line-virtual-inquiry",
  "自宅需求（SIMULATION）",
  "初步報價後議價（SIMULATION）",
  "陳小姐（DEMO / SIMULATION）",
  "SIM-0912-000-003",
  "SIM_LINE_PILOT003",
  "桃園市中壢區（SIMULATION）",
  "客廳＋主臥（SIMULATION）",
  "新成屋窗簾；虛擬丈量後提出初步方案。客戶認為 V1 超出期待，明確希望控制在 NT$105,000 內；仍在比較，未承諾成交。",
  "可（SIMULATION）",
  "quoted",
  "SIMULATION 業務",
  "2026-08-16 進行 V2 報價後需求確認；不得自行再折扣（SIMULATION）",
  new Date("2026-08-16T10:00:00"),
  "DEMO / SIMULATION。初始 78/100（P2），價格異議後調整為 70/100（P2）。V1 NT$128,800；V2 NT$104,800 為調整方案而非未核准折扣；案件維持 quoted。",
  "高",
  null,
  "已排程（SIMULATION）",
  "DEMO / SIMULATION | Pilot Customer #003 | 報價後嫌貴／議價中／V2 待確認",
]];
leads.getRange("T8").formulas = [["=IF(A8=\"\",\"\",Q8-INT(B8))"]];

// Existing Quotes remains the commercial snapshot. Each version is represented as a row, while the new log preserves version semantics.
quotes.getRange("A6:K7").values = [
  [new Date("2026-08-14T09:00:00"), pilotId, "陳小姐（DEMO / SIMULATION）", "V1｜客廳雙層簾＋主臥遮光布簾（SIMULATION）", 128800, "待客戶確認（SIMULATION）", null, null, null, 0.39, "N/A — 價格異議，案件仍進行中（SIMULATION）"],
  [new Date("2026-08-14T15:00:00"), pilotId, "陳小姐（DEMO / SIMULATION）", "V2｜調整為客廳手拉雙層簾＋主臥遮光布簾（SIMULATION）", 104800, "待客戶確認（SIMULATION）", null, null, null, 0.35, "N/A — V2 已送出，待需求確認（SIMULATION）"],
];
quotes.getRange("A6:A7").format.numberFormat = "yyyy-mm-dd hh:mm";
quotes.getRange("E6:G7").format.numberFormat = "NT$#,##0";
quotes.getRange("J6:J7").format.numberFormat = "0.0%";

// Minimal additive log: validates version, objection, and approval flow without changing the base Quotes schema.
const quoteLog = wb.worksheets.add("Quote Version Log");
title(quoteLog, "A1:L1", "Quote Version Log｜Pilot #003（DEMO / SIMULATION）");
quoteLog.getRange("A3:L3").values = [["quote_version_id", "lead_id", "version", "issued_at", "previous_version", "scenario", "quoted_amount", "change_vs_previous", "customer_objection", "negotiation_status", "approval_required", "next_action"]];
header(quoteLog, "A3:L3");
quoteLog.getRange("A4:L5").values = [
  ["SIM-Q-003-V1", pilotId, "V1", new Date("2026-08-14T09:00:00"), "N/A", "原始方案：客廳雙層簾＋主臥遮光布簾", 128800, null, "SIMULATION：『比預期高，希望控制在 NT$105,000 內。』", "objection_received", "否；尚未提出折扣", "以需求調整取代直接折扣；產出 V2。"],
  ["SIM-Q-003-V2", pilotId, "V2", new Date("2026-08-14T15:00:00"), "SIM-Q-003-V1", "調整方案：客廳改手拉雙層簾；保留主臥遮光功能", 104800, null, "價格期待已回應；尚未接受或拒絕。", "awaiting_customer", "若再降價或改付款條件，需人工核准", "2026-08-16 需求確認；若無回覆，建立跟進事件。"],
];
quoteLog.getRange("H4").formulas = [["=IF(G4=\"\",\"\",0)"]];
quoteLog.getRange("H5").formulas = [["=IF(OR(G5=\"\",G4=\"\"),\"\",G5-G4)"]];
body(quoteLog, "A4:L5");
quoteLog.getRange("D4:D5").format.numberFormat = "yyyy-mm-dd hh:mm";
quoteLog.getRange("G4:H5").format.numberFormat = "NT$#,##0;[Red]-NT$#,##0";
quoteLog.getRange("A:A").format.columnWidth = 18;
quoteLog.getRange("B:B").format.columnWidth = 18;
quoteLog.getRange("C:C").format.columnWidth = 10;
quoteLog.getRange("D:D").format.columnWidth = 18;
quoteLog.getRange("E:E").format.columnWidth = 18;
quoteLog.getRange("F:F").format.columnWidth = 37;
quoteLog.getRange("G:H").format.columnWidth = 16;
quoteLog.getRange("I:I").format.columnWidth = 42;
quoteLog.getRange("J:J").format.columnWidth = 20;
quoteLog.getRange("K:K").format.columnWidth = 28;
quoteLog.getRange("L:L").format.columnWidth = 37;
quoteLog.showGridLines = false;
quoteLog.freezePanes.freezeRows(3);
quoteLog.tables.add("A3:L5", true, "QuoteVersionLogTable").showFilterButton = true;

// Follow-up is a task only; no actual outbound message was sent.
followUps.getRange("A9:I9").values = [[
  new Date("2026-08-16T10:00:00"), pilotId, "陳小姐（DEMO / SIMULATION）", "LINE（SIMULATION）",
  "V2 報價後需求確認任務：確認調整方案是否符合預算與功能；未發送真實訊息。",
  "SIMULATION：尚未收到結果，不將價格異議誤判為 Lost。",
  "人工確認是否接受 V2；若要求再折扣，取得核准後才可更新版本。",
  new Date("2026-08-18T10:00:00"), "SIMULATION 業務",
]];

const timeline = wb.worksheets.add("Pilot-003 Timeline");
title(timeline, "A1:I1", "Pilot Customer #003｜報價版本與價格異議壓力測試（DEMO / SIMULATION）");
timeline.getRange("A3:I3").values = [["step", "timestamp", "status_after", "lead_score", "priority", "owner", "action", "result", "next_action"]];
header(timeline, "A3:I3");
timeline.getRange("A4:I11").values = [
  ["1. 建立 Lead", new Date("2026-08-12T10:00:00"), "qualified", 78, "P2 / High", "SIMULATION 業務", "建立虛擬新成屋兩空間案件。", "已完成虛擬需求確認；資料足以建立初步報價。", "產出 V1。"],
  ["2. 建立 V1", new Date("2026-08-14T09:00:00"), "quoted", 78, "P2 / High", "SIMULATION 業務", "建立 V1 NT$128,800；報價有效性與範圍已留在 Version Log。", "V1 為可追溯的商業快照。", "等待客戶意見。"],
  ["3. 價格異議", new Date("2026-08-14T10:15:00"), "quoted", 70, "P2 / High", "SIMULATION 客戶", "模擬客戶表示高於期待，提出 NT$105,000 內的價格期待。", "異議已記錄；非拒絕、非 Lost。", "評估範圍調整。"],
  ["4. 人工決策", new Date("2026-08-14T11:00:00"), "quoted", 70, "P2 / High", "SIMULATION 業務", "決定先調整產品範圍，不自行提供未核准折扣。", "折扣權限成為明確人工關卡。", "建立 V2。"],
  ["5. 建立 V2", new Date("2026-08-14T15:00:00"), "quoted", 70, "P2 / High", "SIMULATION 業務", "V2 NT$104,800；客廳改手拉雙層簾，保留主臥遮光。", "較 V1 減少 NT$24,000；版本與差額可追溯。", "安排需求確認。"],
  ["6. 建立 Follow-up", new Date("2026-08-14T15:10:00"), "quoted", 70, "P2 / High", "SIMULATION 業務", "建立 8/16 內部需求確認任務。", "未發送真實 LINE 或其他外部訊息。", "人工確認 V2 接受度。"],
  ["7. 狀態檢查", new Date("2026-08-14T15:15:00"), "quoted", 70, "P2 / High", "SIMULATION 系統", "同步 Leads、Quotes、Quote Version Log、Follow-ups 與 Dashboard。", "案件仍開放；沒有成交／失單假設。", "依下一次跟進決定後續。"],
  ["8. Dashboard 更新", new Date("2026-08-14T15:20:00"), "quoted", 70, "P2 / High", "SIMULATION 系統", "以 Leads status 計入已報價；以 Version Log 顯示議價中。", "現況 KPI 已反映一筆 quoted 案件。", "觀察 V2 後回覆。"],
];
body(timeline, "A4:I11");
timeline.getRange("B4:B11").format.numberFormat = "yyyy-mm-dd hh:mm";
timeline.getRange("D4:D11").format.numberFormat = "#,##0";
timeline.getRange("A:A").format.columnWidth = 18;
timeline.getRange("B:B").format.columnWidth = 18;
timeline.getRange("C:C").format.columnWidth = 16;
timeline.getRange("D:E").format.columnWidth = 13;
timeline.getRange("F:F").format.columnWidth = 18;
timeline.getRange("G:G").format.columnWidth = 38;
timeline.getRange("H:H").format.columnWidth = 36;
timeline.getRange("I:I").format.columnWidth = 30;
timeline.showGridLines = false;
timeline.freezePanes.freezeRows(3);
timeline.tables.add("A3:I11", true, "Pilot003TimelineTable").showFilterButton = true;

// Append only the new evidence to the existing backlog; v2 is still deferred until recurring real patterns are observed.
backlog.getRange("A1:F1").values = [["Data Model Issues / Backlog｜Pilot #001 + #002 + #003"]];
backlog.getRange("A12:F12").values = [[
  "P0", "Quote 缺少 quote_version_id、previous_version、issued_at、version status 與異議／核准欄位。", "#003", "同一案件 V1／V2 只能靠文字與多列判讀，無法安全自動化議價規則。", "先保留 Quote Version Log；下一輪將其最小欄位併入 Quotes 或關聯表。", "否；第三個情境剛驗證需求，先收集 3–5 筆真實模式。",
]];
body(backlog, "A12:F12");

dashboard.getRange("A24:H24").merge();
dashboard.getRange("A24:H24").values = [["DEMO / SIMULATION｜Pilot #003：V1 NT$128,800 → V2 NT$104,800（差額 NT$24,000）。客戶價格異議已記錄；案件維持 quoted，未假設成交或失單；再折扣需人工核准。"]];
dashboard.getRange("A24:H24").format = { fill: sand, font: { color: text }, wrapText: true, verticalAlignment: "center" };
dashboard.getRange("A24:H24").format.rowHeight = 32;
dashboard.getRange("A26:B27").values = [["#003 議價中（DEMO / SIMULATION）", "案件數"], ["quoted 且有 V2", null]];
header(dashboard, "A26:B26");
dashboard.getRange("B27").formulas = [["=COUNTIF('Leads'!$N$4:$N$203,\"quoted\")"]];
dashboard.getRange("A27:B27").format = { fill: rose, font: { color: text }, borders: { preset: "outside", style: "thin", color: lightBorder } };
dashboard.getRange("B27").format.font = { bold: true, color: navy, size: 14 };

const checks = [
  await wb.inspect({ kind: "table", range: "Leads!A3:V8", include: "values,formulas", tableMaxRows: 10, tableMaxCols: 24 }),
  await wb.inspect({ kind: "table", range: "Quotes!A3:K7", include: "values,formulas", tableMaxRows: 10, tableMaxCols: 12 }),
  await wb.inspect({ kind: "table", range: "Quote Version Log!A1:L5", include: "values,formulas", tableMaxRows: 10, tableMaxCols: 12 }),
  await wb.inspect({ kind: "table", range: "Dashboard!A1:H27", include: "values,formulas", tableMaxRows: 30, tableMaxCols: 8 }),
];
for (const check of checks) console.log(check.ndjson);
const errorScan = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
console.log(errorScan.ndjson);
for (const sheetName of ["Leads", "Quotes", "Quote Version Log", "Pilot-003 Timeline", "Data Model Issues - Backlog", "Dashboard"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${workDir}/pilot-003-${sheetName}.png`, new Uint8Array(await preview.arrayBuffer()));
}
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(output);
