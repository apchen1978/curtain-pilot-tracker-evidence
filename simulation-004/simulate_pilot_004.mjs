// simulate_pilot_004.mjs — Virtual Lead #004 (SIM-PILOT-004) portable simulation seed
// ================================================================================
// STATUS: NOT EXECUTED in this workspace.
//   - Requires the `@oai/artifact-tool` dependency, which is NOT installed in this
//     workspace copy (node_modules is empty, no package.json). Per task rules no
//     packages were installed; a Markdown/JSON/CSV audit artifact was produced instead
//     (see ../simulation-004/lead-004-simulation.md).
//   - Run this file in an environment that has @oai/artifact-tool (e.g. the
//     environment where the original build_/simulate_/verify_ scripts were executed).
// PATHS: relative to this project root (no absolute paths).
// INPUT : outputs/curtain-soft-furnishing-pilot-tracker-pilot-003-quote-version-simulation.xlsx
// OUTPUT: outputs/curtain-soft-furnishing-pilot-tracker-pilot-004-simulation.xlsx
// CONVENTIONS: mirrors simulate_pilot_002/003.mjs (explicit-unknown, nurture/72h,
//   quote-versioning, DEMO/SIMULATION labels). No real messages, no webhooks.
// ================================================================================

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputsDir = path.join(projectRoot, "outputs");
const workDir = path.join(projectRoot, "work");
const input = path.join(outputsDir, "curtain-soft-furnishing-pilot-tracker-pilot-003-quote-version-simulation.xlsx");
const output = path.join(outputsDir, "curtain-soft-furnishing-pilot-tracker-pilot-004-simulation.xlsx");

const navy = "#18344A";
const teal = "#1F6F6D";
const mint = "#E6F3F1";
const sand = "#F7F2E8";
const rose = "#FDE2E2";
const lightBorder = "#D9E2E7";
const text = "#1F2933";
const pilotId = "SIM-PILOT-004";

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
const leads = wb.worksheets.getItem("Leads");
const followUps = wb.worksheets.getItem("Follow-ups");
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

// Lead row 9 (rows 4-5 DEMO, 6 #001, 7 #002, 8 #003). Unknowns stay explicit.
// SCORE intentionally left UNKNOWN: no structured scoring model exists in the project
// (Backlog P0). No quote, no measurement scheduled yet.
leads.getRange("A9:V9").values = [[
  pilotId,
  new Date("2026-08-18T10:30:00"),
  "LINE（SIMULATION）",
  "line-virtual-inquiry",
  "自宅需求（SIMULATION）",
  "傳照片初估（SIMULATION）",
  "陳小姐（DEMO / SIMULATION）",
  "未提供（SIMULATION）",
  "SIM_LINE_PILOT004",
  "新北市新莊區（SIMULATION）",
  "客廳＋主臥＋次臥（SIMULATION）",
  "新成屋全室窗簾：客廳需遮光但白天保留自然光；主臥高度遮光；次臥一般遮光。預算 NT$60,000–80,000；期望 30 天內完成。尚未提供窗戶尺寸、現場照片、格局圖；客戶詢問「大概多少錢？可以先估嗎？」",
  "未提供（SIMULATION）",
  "new_lead",
  "待指派（SIMULATION）",
  "24h 內首次聯繫；收集窗戶尺寸／照片／格局圖（SIMULATION）",
  new Date("2026-08-19T10:30:00"),
  "DEMO / SIMULATION。未安排丈量、未報價；草稿僅供 owner review，未發送。SCORE = UNKNOWN（專案無評分模型）。",
  "高（SIMULATION）",
  null,
  null,
  "DEMO / SIMULATION | Pilot Customer #004 | 未丈量／未報價／待收集尺寸與照片",
]];
leads.getRange("T9").formulas = [["=IF(A9=\"\",\"\",Q9-INT(B9))"]];
leads.getRange("U9").formulas = [["=IF(A9=\"\",\"\",IF(Q9=\"\",\"未排程\",\"已排程\"))"]];

// Planned follow-ups (NOT SENT). Per existing Follow-ups schema these would be
// written only after an event actually occurs; here they are planning rows.
followUps.getRange("A10:I11").values = [
  [new Date("2026-08-18T10:45:00"), pilotId, "陳小姐（DEMO / SIMULATION）", "LINE（SIMULATION）", "第一次 Follow-up 草稿：回應價格問題；索取三空間窗戶尺寸、照片、格局圖。", "未回覆（尚未發送 — SIMULATION）", "待 owner 核准草稿後發送；收集尺寸／照片／格局圖", new Date("2026-08-19T10:30:00"), "SIMULATION 業務"],
  [new Date("2026-08-19T10:30:00"), pilotId, "陳小姐（DEMO / SIMULATION）", "LINE（SIMULATION）", "第二次 Follow-up 任務（如客戶未回覆）：低頻提醒補齊最小資訊；不承諾價格、不安排丈量。", "未回覆（尚未發送 — SIMULATION）", "若 72h 未回覆：依 Pilot #002 慣例轉 nurture，不判定 lost", new Date("2026-08-22T10:30:00"), "SIMULATION 業務"],
];

// Timeline — planned steps are marked; lead_score stays null (SCORE = UNKNOWN).
const timeline = wb.worksheets.add("Pilot-004 Timeline");
title(timeline, "A1:I1", "Pilot Customer #004｜初步估價前資訊收集流程（DEMO / SIMULATION）");
timeline.getRange("A3:I3").values = [["step", "timestamp", "status_after", "lead_score", "priority", "owner", "action", "result", "next_action"]];
header(timeline, "A3:I3");
timeline.getRange("A4:I10").values = [
  ["1. 建立 Lead", new Date("2026-08-18T10:30:00"), "new_lead", null, "待判定", "SIMULATION 系統", "記錄虛擬 LINE 詢問「大概多少錢？可以先估嗎？」", "需求、預算、30 天時程、三空間與遮光需求已記錄；尺寸／照片／格局缺失。", "執行初始資格評估"],
  ["2. 初始資格評估", new Date("2026-08-18T10:40:00"), "new_lead", null, "高（INFERRED）", "SIMULATION 系統", "依既有規則檢視：有預算＋時程＋明確空間需求；無尺寸／照片／格局。", "SCORE = UNKNOWN（無評分模型）；狀態維持 new_lead（尚未聯繫）。", "建立第一次聯繫任務（24h 內）"],
  ["3. 第一次 Follow-up（草稿）", new Date("2026-08-18T10:45:00"), "new_lead", null, "高（INFERRED）", "SIMULATION 業務", "草稿已備妥，僅供 owner review。", "未發送；無任何真實訊息。", "待 owner 核准後發送"],
  ["4. 收集補充資料", new Date("2026-08-19T10:30:00"), "contacted", null, "高（INFERRED）", "SIMULATION 業務", "索取三空間窗戶尺寸、照片、格局圖。", "等待客戶回覆。", "依回覆更新 Qualification"],
  ["5. 安排丈量", null, "measurement_booked", null, "高（INFERRED）", "SIMULATION 業務", "依 Pilot #001 模式：先虛擬丈量（照片／平面圖），正式案需現場複測。", "前置條件：需先取得照片或格局。", "執行丈量"],
  ["6. 初步估價", null, "quoted", null, "高（INFERRED）", "SIMULATION 業務", "以實際尺寸與布料選擇提供初步估價範圍。", "前置條件：需 owner 提供定價基準（專案無價格表）。", "產生正式報價（Quote Version Log）"],
  ["7. 正式報價", null, "quoted", null, "高（INFERRED）", "SIMULATION 業務", "依 Pilot #003 模式：版本化報價；折扣／降價需人工核准。", "尚未發生。", "追蹤確認"],
];
body(timeline, "A4:I10");
timeline.getRange("B4:B5").format.numberFormat = "yyyy-mm-dd hh:mm";
timeline.getRange("A:I").format.columnWidth = 20;
timeline.getRange("G:G").format.columnWidth = 38;
timeline.getRange("H:H").format.columnWidth = 40;
timeline.getRange("I:I").format.columnWidth = 30;
timeline.showGridLines = false;
timeline.freezePanes.freezeRows(3);
timeline.tables.add("A3:I10", true, "Pilot004TimelineTable").showFilterButton = true;

// Backlog: one proposed P1 entry (pricing baseline) — proposal only, not applied to a v2.
backlog.getRange("A13:F13").values = [[
  "P1",
  "專案缺少結構化定價／估價基準（unit price list）；「初步估價」無法標準化，只能人工判斷。#004 提案。",
  "#004",
  "無法給出有依據的初步估價；客戶問「可以先估嗎」時只能索取資料。",
  "建立估價基準表（品項 × 單位價格或區間）。",
  "否；先建立定價基準。",
]];
body(backlog, "A13:F13");

// Dashboard: annotation only (KPIs remain formula-driven from Leads).
dashboard.getRange("A28:H28").merge();
dashboard.getRange("A28:H28").values = [["DEMO / SIMULATION｜Pilot #004：陳小姐，新莊新成屋三空間。有預算（NT$60,000–80,000）與 30 天時程；缺尺寸／照片／格局。SCORE = UNKNOWN；未丈量、未報價；第一次 Follow-up 草稿待 owner 核准（未發送）。"]];
dashboard.getRange("A28:H28").format = { fill: sand, font: { color: text }, wrapText: true, verticalAlignment: "center" };
dashboard.getRange("A28:H28").format.rowHeight = 32;

// Inspections mirroring the existing verify pattern.
const checks = [
  await wb.inspect({ kind: "table", range: "Leads!A3:V9", include: "values,formulas", tableMaxRows: 10, tableMaxCols: 24 }),
  await wb.inspect({ kind: "table", range: "Dashboard!A1:H28", include: "values,formulas", tableMaxRows: 30, tableMaxCols: 8 }),
];
for (const check of checks) console.log(check.ndjson);
const errorScan = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
console.log(errorScan.ndjson);

for (const sheetName of ["Leads", "Follow-ups", "Pilot-004 Timeline", "Data Model Issues - Backlog", "Dashboard"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(workDir, `pilot-004-${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
}
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(output);
