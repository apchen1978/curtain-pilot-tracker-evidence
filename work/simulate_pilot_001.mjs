import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const input = path.resolve(__dirname, "..", "outputs", "curtain-soft-furnishing-pilot-tracker.xlsx");
const output = path.resolve(__dirname, "..", "outputs", "curtain-soft-furnishing-pilot-tracker-pilot-001-simulation.xlsx");
const workDir = __dirname;
const simulation = "DEMO / SIMULATION";
const pilotId = "SIM-PILOT-001";
const navy = "#18344A";
const teal = "#1F6F6D";
const mint = "#E6F3F1";
const sand = "#F7F2E8";
const lightBorder = "#D9E2E7";
const text = "#1F2933";

const source = await FileBlob.load(input);
const wb = await SpreadsheetFile.importXlsx(source);
const leads = wb.worksheets.getItem("Leads");
const followUps = wb.worksheets.getItem("Follow-ups");
const quotes = wb.worksheets.getItem("Quotes");
const dashboard = wb.worksheets.getItem("Dashboard");

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

// 1–3 and final Lead record: one row reflects the current terminal state; the full state history lives in Timeline.
leads.getRange("A6:V6").values = [[
  pilotId,
  new Date("2026-08-14T10:10:00"),
  "LINE（SIMULATION）",
  "line-virtual-inquiry",
  "自宅需求",
  "預約丈量",
  "王先生（DEMO / SIMULATION）",
  "SIM-0912-000-001",
  "SIM_LINE_PILOT001",
  "新北市",
  "全室",
  "新成屋全室窗簾：客廳＋主臥＋次臥；預算 NT$80,000–120,000；預計 14 天內決策。",
  "可",
  "won",
  "SIMULATION 業務",
  "安排安裝前確認（SIMULATION）",
  new Date("2026-08-28"),
  "DEMO / SIMULATION。資格評估 86/100，高優先；已完成虛擬丈量、報價與成交模擬。",
  "高",
  null,
  null,
  "DEMO / SIMULATION | Pilot Customer #001 | 已成交（虛擬）",
]];
leads.getRange("T6").formulas = [["=IF(A6=\"\",\"\",Q6-INT(B6))"]];
leads.getRange("U6").formulas = [["=IF(A6=\"\",\"\",IF(Q6=\"\",\"未排程\",\"已排程\"))"]];

// 4, 5 and 9: Follow-up records.
followUps.getRange("A5:I6").values = [
  [new Date("2026-08-14T15:00:00"), pilotId, "王先生（DEMO / SIMULATION）", "LINE（SIMULATION）", "第一次跟進：確認新成屋、三個空間、預算與 14 天決策窗口。", "SIMULATION 回覆：提供平面圖與空間照片，同意安排虛擬丈量／需求確認。", "安排 8/17 虛擬丈量與布料偏好確認", new Date("2026-08-17"), "SIMULATION 業務"],
  [new Date("2026-08-21T17:00:00"), pilotId, "王先生（DEMO / SIMULATION）", "LINE（SIMULATION）", "第二次跟進：確認初步報價、遮光需求與主臥私密性。", "SIMULATION 回覆：報價在預算內；確認選擇客廳紗簾＋三組遮光布簾，8/26 最終確認。", "8/26 虛擬成交確認；安排安裝前確認", new Date("2026-08-26"), "SIMULATION 業務"],
];

// 8 and 11: Quote and won result.
quotes.getRange("A5:K5").values = [[
  new Date("2026-08-18"),
  pilotId,
  "王先生（DEMO / SIMULATION）",
  "客廳紗簾＋客廳／主臥／次臥遮光布簾（SIMULATION）",
  104800,
  "已成交",
  104800,
  new Date("2026-09-05"),
  null,
  0.38,
  "N/A — DEMO / SIMULATION 已成交",
]];

// 7: Minimal measurement record, kept separate from commercial quote data.
const measurements = wb.worksheets.add("Measurements");
title(measurements, "A1:K1", "Measurements｜虛擬丈量與需求確認");
measurements.getRange("A3:K3").values = [["measurement_date", "lead_id", "customer_name", "space", "window_count", "estimated_width_cm", "estimated_height_cm", "light_control", "privacy_need", "fabric_preference", "measurement_note"]];
header(measurements, "A3:K3");
measurements.getRange("A4:K6").values = [
  [new Date("2026-08-17"), pilotId, "王先生（DEMO / SIMULATION）", "客廳", 2, 360, 240, "日間採光＋夜間遮光", "中", "紗簾＋遮光布簾", "DEMO / SIMULATION：依平面圖與照片進行虛擬丈量；正式案仍需現場複測。"],
  [new Date("2026-08-17"), pilotId, "王先生（DEMO / SIMULATION）", "主臥", 1, 220, 240, "全遮光", "高", "全遮光布簾", "DEMO / SIMULATION：需確認窗簾盒深度與軌道位置。"],
  [new Date("2026-08-17"), pilotId, "王先生（DEMO / SIMULATION）", "次臥", 1, 180, 240, "柔和遮光", "中", "遮光布簾", "DEMO / SIMULATION：預留兒童房／客房用途待最後確認。"],
];
body(measurements, "A4:K6");
measurements.getRange("A4:A203").format.numberFormat = "yyyy-mm-dd";
measurements.getRange("E4:G203").format.numberFormat = "#,##0";
measurements.getRange("A:K").format.columnWidth = 18;
measurements.getRange("D:D").format.columnWidth = 14;
measurements.getRange("H:J").format.columnWidth = 20;
measurements.getRange("K:K").format.columnWidth = 40;
measurements.showGridLines = false;
measurements.freezePanes.freezeRows(3);
const measurementsTable = measurements.tables.add("A3:K203", true, "MeasurementsTable");
measurementsTable.showFilterButton = true;

// Complete audit trail of the 11 requested steps.
const timeline = wb.worksheets.add("Pilot-001 Timeline");
title(timeline, "A1:H1", "Pilot Customer #001｜完整流程事件紀錄（DEMO / SIMULATION）");
timeline.getRange("A3:H3").values = [["step", "timestamp", "status_after", "owner", "action", "result", "next_action", "scheduled_for"]];
header(timeline, "A3:H3");
timeline.getRange("A4:H14").values = [
  ["1. 建立 Lead", new Date("2026-08-14T10:10:00"), "new_lead", "SIMULATION 系統", "以虛擬 LINE 詢問建立 Lead。", "完整資料、14 天決策窗口、預算已記錄。", "執行資格評估", new Date("2026-08-14T10:15:00")],
  ["2. AI Lead Qualification", new Date("2026-08-14T10:15:00"), "new_lead", "SIMULATION AI", "解析需求、預算、地區、空間與決策期。", "新成屋全室、預算明確、可提供照片、需求可執行。", "計算 Lead Score", new Date("2026-08-14T10:16:00")],
  ["3. 優先級 / Lead Score", new Date("2026-08-14T10:16:00"), "qualified", "SIMULATION AI", "Lead Score 86/100：預算 25、決策期 20、需求完整 20、可丈量 15、可聯繫 6。", "高優先；應於 24 小時內首次聯繫。", "第一次跟進", new Date("2026-08-14T15:00:00")],
  ["4. 第一次 Follow-up", new Date("2026-08-14T15:00:00"), "contacted", "SIMULATION 業務", "以虛擬 LINE 確認預算、三個空間與交期。", "等待平面圖與照片。", "收取客戶回覆", new Date("2026-08-15T10:00:00")],
  ["5. 模擬客戶回覆", new Date("2026-08-15T10:00:00"), "contacted", "SIMULATION 客戶", "提供平面圖與照片；接受虛擬丈量。", "需求資訊足夠，無需再追問基本資格。", "排定丈量／需求確認", new Date("2026-08-17T14:00:00")],
  ["6. 推進至丈量階段", new Date("2026-08-15T10:05:00"), "measurement_booked", "SIMULATION 業務", "建立丈量預約與需求確認。", "丈量日期已排定。", "執行虛擬丈量", new Date("2026-08-17T14:00:00")],
  ["7. 虛擬丈量資料", new Date("2026-08-17T14:00:00"), "qualified", "SIMULATION 業務", "記錄客廳、主臥、次臥的估算尺寸、遮光與隱私需求。", "可進入初步報價；標記正式案仍需現場複測。", "產生初步報價", new Date("2026-08-18T11:00:00")],
  ["8. 初步報價", new Date("2026-08-18T11:00:00"), "quoted", "SIMULATION 業務", "產生 NT$104,800 初步報價，落在客戶預算內。", "客戶待確認布料與最終組合。", "第二次 Follow-up", new Date("2026-08-21T17:00:00")],
  ["9. 第二次 Follow-up", new Date("2026-08-21T17:00:00"), "quoted", "SIMULATION 業務", "確認報價、遮光與主臥私密性需求。", "客戶口頭接受範圍，約定最終確認。", "最終確認", new Date("2026-08-26T16:00:00")],
  ["10. Dashboard 更新", new Date("2026-08-21T17:05:00"), "quoted", "SIMULATION 系統", "Leads / Follow-ups / Quotes 已同步；Dashboard 由 Leads 公式更新。", "漏斗可看到 1 件模擬成交前案件。", "記錄結果", new Date("2026-08-26T16:00:00")],
  ["11. 最終結果", new Date("2026-08-26T16:00:00"), "won", "SIMULATION 客戶", "虛擬客戶確認 NT$104,800 報價與 9/5 安裝預排。", "成交（DEMO / SIMULATION）；無任何真實訂單或外部通知。", "安裝前確認（僅模擬）", new Date("2026-08-28T10:00:00")],
];
body(timeline, "A4:H14");
timeline.getRange("B4:B14").format.numberFormat = "yyyy-mm-dd hh:mm";
timeline.getRange("H4:H14").format.numberFormat = "yyyy-mm-dd hh:mm";
timeline.getRange("A:H").format.columnWidth = 20;
timeline.getRange("E:E").format.columnWidth = 36;
timeline.getRange("F:F").format.columnWidth = 40;
timeline.getRange("G:G").format.columnWidth = 28;
timeline.showGridLines = false;
timeline.freezePanes.freezeRows(3);
const timelineTable = timeline.tables.add("A3:H14", true, "Pilot001TimelineTable");
timelineTable.showFilterButton = true;

// Dashboard is formula-driven from Leads. Add only a simulation annotation, not hard-coded KPIs.
dashboard.getRange("A15:H15").merge();
dashboard.getRange("A15:H15").values = [["DEMO / SIMULATION｜Pilot Customer #001：王先生，新北市新成屋全室窗簾；Lead Score 86/100；最終結果：已成交 NT$104,800（虛擬）。"]];
dashboard.getRange("A15:H15").format = { fill: mint, font: { bold: true, color: navy }, wrapText: true, verticalAlignment: "center", borders: { preset: "outside", style: "thin", color: lightBorder } };
dashboard.getRange("A15:H15").format.rowHeight = 30;
dashboard.getRange("A17:H17").merge();
dashboard.getRange("A17:H17").values = [["流程警示：Dashboard 只以 Leads 的現況狀態計數；歷程、丈量與成交原因需分別查看 Pilot-001 Timeline、Measurements、Quotes。"]];
dashboard.getRange("A17:H17").format = { fill: sand, font: { color: text }, wrapText: true };

const reconciliation = await wb.inspect({ kind: "table", range: "Dashboard!A1:H17", include: "values,formulas", tableMaxRows: 20, tableMaxCols: 10 });
console.log(reconciliation.ndjson);
const errorScan = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan" });
console.log(errorScan.ndjson);
for (const sheetName of ["Leads", "Follow-ups", "Quotes", "Measurements", "Pilot-001 Timeline", "Dashboard"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${workDir}/after-${sheetName}.png`, new Uint8Array(await preview.arrayBuffer()));
}
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(output);
