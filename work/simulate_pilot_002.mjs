import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const input = path.resolve(__dirname, "..", "outputs", "curtain-soft-furnishing-pilot-tracker-pilot-001-simulation.xlsx");
const output = path.resolve(__dirname, "..", "outputs", "curtain-soft-furnishing-pilot-tracker-pilot-002-simulation.xlsx");
const workDir = __dirname;
const navy = "#18344A";
const teal = "#1F6F6D";
const mint = "#E6F3F1";
const sand = "#F7F2E8";
const rose = "#FDE2E2";
const lightBorder = "#D9E2E7";
const text = "#1F2933";
const pilotId = "SIM-PILOT-002";

const source = await FileBlob.load(input);
const wb = await SpreadsheetFile.importXlsx(source);
const leads = wb.worksheets.getItem("Leads");
const followUps = wb.worksheets.getItem("Follow-ups");
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

// Minimal status extension: nurture is required by this scenario; stale/lost are intentionally not inferred.
leads.getRange("N4:N203").dataValidation = { rule: { type: "list", values: ["new_lead", "contacted", "qualified", "measurement_booked", "quoted", "won", "lost", "nurture"] } };

// Current Lead state. Unknown values remain explicit; no dimensions, budget, decision date, measurement or quote are fabricated.
leads.getRange("A7:V7").values = [[
  pilotId,
  new Date("2026-08-14T11:30:00"),
  "LINE（SIMULATION）",
  "line-virtual-inquiry",
  "未知（SIMULATION）",
  "價格詢問（SIMULATION）",
  "未提供（DEMO / SIMULATION）",
  null,
  "SIM_LINE_PILOT002",
  "未提供",
  "未提供",
  "請問全室窗簾大概多少錢？",
  "未提供",
  "nurture",
  "SIMULATION 業務",
  "7 天後培育提醒：邀請提供地區、空間與照片（SIMULATION）",
  new Date("2026-08-25"),
  "DEMO / SIMULATION。初始 Lead Score 15/100（P4）；僅確認新北新成屋、尚未丈量後更新為 30/100（P3）。72 小時未回覆，不判定 Lost。",
  "低",
  null,
  "NURTURE — follow-up re-scheduled（SIMULATION）",
  "DEMO / SIMULATION | Pilot Customer #002 | 無預算／無決策日／無丈量／無報價",
]];
leads.getRange("T7").formulas = [["=IF(A7=\"\",\"\",Q7-INT(B7))"]];

// Two internal follow-up records. Neither represents a real message.
followUps.getRange("A7:I8").values = [
  [new Date("2026-08-14T11:40:00"), pilotId, "未提供（DEMO / SIMULATION）", "LINE（SIMULATION）", "第一次 Follow-up 任務：詢問地區、房型／空間、預計入住日、照片與是否需初估。未發送真實訊息。", "SIMULATION 回覆：新北新成屋，還沒量尺寸，想先抓大概價位。未提供預算、照片、入住日或產品偏好。", "以價位區間培育，不承諾精確報價；48 小時後追蹤。", new Date("2026-08-16T11:40:00"), "SIMULATION 業務"],
  [new Date("2026-08-16T11:40:00"), pilotId, "未提供（DEMO / SIMULATION）", "LINE（SIMULATION）", "第二次 Follow-up 任務：提供需要補齊的最小資訊清單。未發送真實訊息。", "SIMULATION：截至 2026-08-19 11:40，72 小時未回覆。", "轉入 nurture；7 天後一次低頻提醒。", new Date("2026-08-25"), "SIMULATION 業務"],
];

// End-to-end state history with score changes and the 72h response breach.
const timeline = wb.worksheets.add("Pilot-002 Timeline");
title(timeline, "A1:I1", "Pilot Customer #002｜低資訊培育流程（DEMO / SIMULATION）");
timeline.getRange("A3:I3").values = [["step", "timestamp", "status_after", "lead_score", "priority", "owner", "action", "result", "next_action"]];
header(timeline, "A3:I3");
timeline.getRange("A4:I15").values = [
  ["1. 建立 Lead", new Date("2026-08-14T11:30:00"), "new_lead", 15, "P4 / Low", "SIMULATION 系統", "記錄原始 LINE 詢問：『請問全室窗簾大概多少錢？』", "除來源外，沒有尺寸、照片、預算、日期、產品偏好或客戶姓名。", "執行初始 Qualification"],
  ["2. AI Qualification", new Date("2026-08-14T11:32:00"), "new_lead", 15, "P4 / Low", "SIMULATION AI", "辨識為價格探索型詢問；資訊不足，不能做產品、價格或交期承諾。", "可培育，不可報價，不可安排丈量。", "計算初始 Lead Score"],
  ["3. 初始 Score / Priority", new Date("2026-08-14T11:33:00"), "qualified", 15, "P4 / Low", "SIMULATION AI", "Score：來源可聯繫 10 + 有全室需求 5；其餘證據為 0。", "低優先；最小目標是取得可判斷資料。", "建立第一次 Follow-up 任務"],
  ["4. 第一次 Follow-up", new Date("2026-08-14T11:40:00"), "contacted", 15, "P4 / Low", "SIMULATION 業務", "內部建立提問清單，不實際傳送。", "等待少量補充資訊。", "記錄模擬回覆"],
  ["5. 少量資訊回覆", new Date("2026-08-15T09:30:00"), "contacted", 15, "P4 / Low", "SIMULATION 客戶", "僅提供：新北新成屋、尚未丈量、想先抓大概價位。", "仍無預算、入住日、照片、尺寸或材質偏好。", "更新 Qualification"],
  ["6. 更新 Qualification", new Date("2026-08-15T09:35:00"), "qualified", 30, "P3 / Nurture", "SIMULATION AI", "增加地區 5、專案存在性 5、回覆意願 5；不增加預算／日期／量測分數。", "可低頻培育；仍禁止精確報價或成交預測。", "更新 Score / Priority"],
  ["7. 最終 Score / Priority", new Date("2026-08-15T09:36:00"), "qualified", 30, "P3 / Nurture", "SIMULATION AI", "15 → 30：只因增加新北新成屋與一次回覆，不代表高意向。", "意向維持低；處理優先級由 P4 升至 P3。", "建立第二次 Follow-up 任務"],
  ["8. 第二次 Follow-up", new Date("2026-08-16T11:40:00"), "contacted", 30, "P3 / Nurture", "SIMULATION 業務", "內部建立 48 小時後追蹤任務，不實際傳送。", "等待回覆。", "檢查 72 小時未回覆"],
  ["9. 72h 未回覆", new Date("2026-08-19T11:40:00"), "contacted", 30, "P3 / Nurture", "SIMULATION 系統", "距第二次 Follow-up 72 小時無新資訊。", "Follow-up 事件應標為 overdue；既有主表無事件級 overdue 欄位。", "判斷 lifecycle"],
  ["10. Lifecycle 判斷", new Date("2026-08-19T11:45:00"), "nurture", 30, "P3 / Nurture", "SIMULATION 業務", "不採用 lost：客戶只是不完整且暫時失聯，沒有拒絕或不適配證據。", "轉 nurture；stale 應保留給更長週期、可配置的無回覆門檻。", "建立低頻下一次 Follow-up"],
  ["11. 下一次 Follow-up", new Date("2026-08-19T11:50:00"), "nurture", 30, "P3 / Nurture", "SIMULATION 業務", "將下次培育任務排在 2026-08-25。", "Lead 保持開放，不建立丈量或 Quote。", "更新 Dashboard"],
  ["12. Dashboard 更新", new Date("2026-08-19T11:55:00"), "nurture", 30, "P3 / Nurture", "SIMULATION 系統", "Leads 與 Follow-ups 同步；Dashboard 顯示當前 nurture 名單。", "資料模型問題已記入 Backlog。", "等待下一次培育觸發"],
];
body(timeline, "A4:I15");
timeline.getRange("B4:B15").format.numberFormat = "yyyy-mm-dd hh:mm";
timeline.getRange("D4:D15").format.numberFormat = "#,##0";
timeline.getRange("A:I").format.columnWidth = 20;
timeline.getRange("G:G").format.columnWidth = 38;
timeline.getRange("H:H").format.columnWidth = 42;
timeline.getRange("I:I").format.columnWidth = 32;
timeline.showGridLines = false;
timeline.freezePanes.freezeRows(3);
const timelineTable = timeline.tables.add("A3:I15", true, "Pilot002TimelineTable");
timelineTable.showFilterButton = true;

// Accumulated, not redesigned: records issues from both scenarios before any v2 work.
const backlog = wb.worksheets.add("Data Model Issues - Backlog");
title(backlog, "A1:F1", "Data Model Issues / Backlog｜Pilot #001 + #002");
backlog.getRange("A3:F3").values = [["priority", "issue", "exposed_by", "current impact", "minimal next change", "v2 now?"]];
header(backlog, "A3:F3");
backlog.getRange("A4:F11").values = [
  ["P0", "Lead Score、Priority 與 score_updated_at 沒有結構化欄位。", "#001 + #002", "分數只能散落在 notes／timeline，無法比較或排序。", "在 Leads 增加 lead_score、priority、score_updated_at、score_reason。", "否；可做小幅欄位增補。"],
  ["P0", "Follow-up 沒有 sent_at、response_due_at、outcome、overdue_at 或 closed_at。", "#002", "72h 未回覆只能寫在 Timeline，Dashboard 無法可靠統計。", "為 Follow-ups 加入 event status 與 due／response 欄位。", "否；先補事件欄位。"],
  ["P1", "status 沒有 nurture 的完整定義，stale 與 lost 的門檻未規範。", "#002", "不同業務可能任意結案或無限追蹤。", "定義 state transition 與可配置的無回覆門檻。", "否；先定義規則。"],
  ["P1", "decision_due_date 不存在；未知值只能藏在需求摘要。", "#001 + #002", "無法追蹤 14 天決策期或刻意保留未知。", "增加 decision_due_date 與 decision_date_confidence。", "否；小欄位增補。"],
  ["P1", "丈量缺少 estimate／verified 與現場複測旗標。", "#001", "虛擬估算可能被誤當施工依據。", "Measurements 增加 measurement_mode、verified_at。", "否；小欄位增補。"],
  ["P1", "Quote 缺少版本、明細、有效期限與折扣／議價歷程。", "#001", "無法支援正式報價反覆調整。", "Pilot #003 先以 Quote Version log 測試。", "否；先驗證版本需求。"],
  ["P2", "Dashboard 是現況快照，沒有歷程轉換率、滯留天數或過期事件。", "#001 + #002", "難以評估漏斗與跟進品質。", "先以 Follow-up event 補資料，再做衍生 KPI。", "否；避免現在重構。"],
  ["P2", "未知、未提供、不適用沒有統一表示法。", "#002", "空白與未知混用，AI 評分／報表會失真。", "建立枚舉與資料輸入規則。", "否；先統一規範。"],
];
body(backlog, "A4:F11");
backlog.getRange("A:A").format.columnWidth = 10;
backlog.getRange("B:B").format.columnWidth = 35;
backlog.getRange("C:C").format.columnWidth = 16;
backlog.getRange("D:D").format.columnWidth = 35;
backlog.getRange("E:E").format.columnWidth = 40;
backlog.getRange("F:F").format.columnWidth = 22;
backlog.showGridLines = false;
backlog.freezePanes.freezeRows(3);
const backlogTable = backlog.tables.add("A3:F11", true, "DataModelBacklogTable");
backlogTable.showFilterButton = true;

// Formula-based current-state dashboard extension; past overdue cannot be reconstructed from a single current Lead row.
dashboard.getRange("A19:B19").values = [["#002 目前培育狀態", "名單數"]];
header(dashboard, "A19:B19");
dashboard.getRange("A20:B20").values = [["nurture（DEMO / SIMULATION）", null]];
dashboard.getRange("B20").formulas = [["=COUNTIF('Leads'!$N$4:$N$203,\"nurture\")"]];
dashboard.getRange("A20:B20").format = { fill: rose, font: { color: text }, borders: { preset: "outside", style: "thin", color: lightBorder } };
dashboard.getRange("B20").format.font = { bold: true, color: navy, size: 14 };
dashboard.getRange("A22:H22").merge();
dashboard.getRange("A22:H22").values = [["DEMO / SIMULATION｜Pilot #002：初始 15/100（P4）→ 30/100（P3）；72h 未回覆後轉 nurture，未建立丈量或 Quote。歷史 overdue 無法由現行主表重建，已列入 Backlog。"]];
dashboard.getRange("A22:H22").format = { fill: sand, font: { color: text }, wrapText: true, verticalAlignment: "center" };
dashboard.getRange("A22:H22").format.rowHeight = 32;

const reconciled = await wb.inspect({ kind: "table", range: "Dashboard!A1:H22", include: "values,formulas", tableMaxRows: 24, tableMaxCols: 10 });
console.log(reconciled.ndjson);
const errorScan = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
console.log(errorScan.ndjson);
for (const sheetName of ["Leads", "Follow-ups", "Pilot-002 Timeline", "Data Model Issues - Backlog", "Dashboard"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${workDir}/pilot-002-${sheetName}.png`, new Uint8Array(await preview.arrayBuffer()));
}
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(output);
