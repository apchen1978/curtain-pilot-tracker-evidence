import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.resolve(__dirname, "..", "outputs");
const workDir = __dirname;
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(workDir, { recursive: true });

const wb = Workbook.create();
const guide = wb.worksheets.add("00_使用說明");
const leads = wb.worksheets.add("Leads");
const followUps = wb.worksheets.add("Follow-ups");
const quotes = wb.worksheets.add("Quotes");
const dashboard = wb.worksheets.add("Dashboard");

const navy = "#18344A";
const teal = "#1F6F6D";
const mint = "#E6F3F1";
const sand = "#F7F2E8";
const lightBorder = "#D9E2E7";
const text = "#1F2933";

function title(sheet, range, value) {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[value]];
  sheet.getRange(range).format = {
    fill: navy,
    font: { bold: true, color: "#FFFFFF", size: 16 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  sheet.getRange(range).format.rowHeight = 30;
}

function header(sheet, range) {
  sheet.getRange(range).format = {
    fill: teal,
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: lightBorder },
  };
}

function tableStyle(sheet, range) {
  sheet.getRange(range).format = {
    font: { color: text, size: 10 },
    verticalAlignment: "center",
    borders: { preset: "inside", style: "thin", color: lightBorder },
  };
}

// Guide
title(guide, "A1:H1", "窗簾／軟裝 Lead-to-Quote Pilot｜Google Sheets Ready");
guide.getRange("A3:H3").merge();
guide.getRange("A3:H3").values = [["目的：先穩定接住網站名單，再建立跟進、報價與成交的共同資料底座。此檔為本地範本；匯入 Google Sheets 後再連接 Make。"]];
guide.getRange("A3:H3").format = { fill: mint, font: { color: text }, wrapText: true, verticalAlignment: "center" };
guide.getRange("A3:H3").format.rowHeight = 34;
guide.getRange("A5:D5").values = [["上線順序", "完成條件", "負責人", "狀態"]];
header(guide, "A5:D5");
guide.getRange("A6:D10").values = [
  ["1. 匯入 Google Sheets", "保留工作表名稱與 Leads 第一列欄位", "營運負責人", "待開始"],
  ["2. 建立 Make Webhook", "收到一筆完整的測試 payload", "Make 負責人", "待開始"],
  ["3. 寫入 Leads", "欄位沒有錯位；缺 LINE／照片也不報錯", "Make 負責人", "待開始"],
  ["4. 發送通知", "收件人可在 24 小時內看見名單", "業務／老闆", "待開始"],
  ["5. 五情境驗收", "完整、無 LINE、無照片、設計師、丈量預約皆成功", "營運負責人", "待開始"],
];
tableStyle(guide, "A6:D10");
guide.getRange("D6:D10").dataValidation = { rule: { type: "list", values: ["待開始", "進行中", "完成", "阻塞"] } };
guide.getRange("A12:H12").values = [["Make 欄位契約（Webhook → Leads）", "", "", "", "", "", "", ""]];
guide.getRange("A12:H12").merge();
guide.getRange("A12:H12").format = { fill: sand, font: { bold: true, color: navy } };
guide.getRange("A13:D13").values = [["欄位", "來源", "預設／規則", "必要性"]];
header(guide, "A13:D13");
guide.getRange("A14:D31").values = [
  ["lead_id", "Make Tools", "CURTAIN-YYYYMMDD-HHMMSS", "必要"],
  ["created_at", "Make now", "ISO 日期時間", "必要"],
  ["source_platform", "Webhook", "website", "必要"],
  ["source_page", "Webhook", "例如 home-consultation", "建議"],
  ["customer_type", "Webhook", "自宅需求／設計師合作", "必要"],
  ["inquiry_type", "Webhook", "免費諮詢／傳照片初估／預約丈量", "必要"],
  ["name", "Webhook", "客戶姓名", "必要"],
  ["phone", "Webhook", "電話", "必要"],
  ["line_id", "Webhook", "可空白", "選填"],
  ["area", "Webhook", "服務地區", "必要"],
  ["space_type", "Webhook", "客廳／臥室／書房／全室／商空／其他", "建議"],
  ["need_summary", "Webhook", "需求描述", "必要"],
  ["has_photos", "Webhook", "可／暫時無法", "建議"],
  ["status", "Make 固定值", "new_lead", "必要"],
  ["owner", "Make 固定值", "先指定一位負責人", "必要"],
  ["next_action", "Make 固定值", "24h內首次聯繫", "必要"],
  ["next_followup_date", "Make 或人工", "可空白；預設明日", "建議"],
  ["notes", "人工", "可空白", "選填"],
];
tableStyle(guide, "A14:D31");
guide.getRange("A34:H34").merge();
guide.getRange("A34:H34").values = [["重要：正式上線前，請將 Leads 中標示 DEMO 的兩筆範例資料刪除；本檔未連接 Google、Make、LINE 或 Email。"]];
guide.getRange("A34:H34").format = { fill: "#FFF3CD", font: { bold: true, color: "#6B4E00" }, wrapText: true };
guide.getRange("A34:H34").format.rowHeight = 28;
guide.getRange("A:H").format.columnWidth = 18;
guide.getRange("B:B").format.columnWidth = 24;
guide.getRange("C:C").format.columnWidth = 32;
guide.showGridLines = false;
guide.freezePanes.freezeRows(5);

// Leads
title(leads, "A1:V1", "Leads｜網站名單與案件母表");
const leadHeaders = [["lead_id", "created_at", "source_platform", "source_page", "customer_type", "inquiry_type", "name", "phone", "line_id", "area", "space_type", "need_summary", "has_photos", "status", "owner", "next_action", "next_followup_date", "notes", "intent_level", "days_to_followup", "followup_health", "record_note"]];
leads.getRange("A3:V3").values = leadHeaders;
header(leads, "A3:V3");
leads.getRange("A4:V5").values = [
  ["DEMO-20260814-001", new Date("2026-08-14T09:30:00"), "website", "home-consultation", "自宅需求", "預約丈量", "王小美（DEMO）", "0912345678", "amywang", "台北市大安區", "客廳", "新家遮光窗簾；已提供照片，希望九月前完成", "可", "new_lead", "待指派", "24h內首次聯繫", new Date("2026-08-15"), "刪除前僅供測試", "高", null, null, "DEMO — 上線前刪除"],
  ["DEMO-20260814-002", new Date("2026-08-14T11:00:00"), "website", "home-consultation", "設計師合作", "傳照片初估", "林設計師（DEMO）", "0987654321", "", "新北市板橋區", "全室", "需要先估價；尺寸與照片待補", "暫時無法", "contacted", "待指派", "索取尺寸與照片", new Date("2026-08-18"), "刪除前僅供測試", "中", null, null, "DEMO — 上線前刪除"],
];
leads.getRange("T4").formulas = [["=IF(A4=\"\",\"\",Q4-INT(B4))"]];
leads.getRange("T4:T203").fillDown();
leads.getRange("U4").formulas = [["=IF(A4=\"\",\"\",IF(Q4=\"\",\"未排程\",\"已排程\"))"]];
leads.getRange("U4:U203").fillDown();
tableStyle(leads, "A4:V203");
leads.getRange("B4:B203").format.numberFormat = "yyyy-mm-dd hh:mm";
leads.getRange("Q4:Q203").format.numberFormat = "yyyy-mm-dd";
leads.getRange("T4:T203").format.numberFormat = "#,##0";
leads.getRange("E4:E203").dataValidation = { rule: { type: "list", values: ["自宅需求", "設計師合作"] } };
leads.getRange("F4:F203").dataValidation = { rule: { type: "list", values: ["免費諮詢", "傳照片初估", "預約丈量"] } };
leads.getRange("K4:K203").dataValidation = { rule: { type: "list", values: ["客廳", "臥室", "書房", "全室", "商空", "其他"] } };
leads.getRange("M4:M203").dataValidation = { rule: { type: "list", values: ["可", "暫時無法"] } };
leads.getRange("N4:N203").dataValidation = { rule: { type: "list", values: ["new_lead", "contacted", "qualified", "measurement_booked", "quoted", "won", "lost"] } };
leads.getRange("S4:S203").dataValidation = { rule: { type: "list", values: ["高", "中", "低"] } };
leads.getRange("U4:U203").conditionalFormats.add("containsText", { text: "未排程", format: { fill: "#FDE2E2", font: { color: "#9B1C1C", bold: true } } });
leads.getRange("A:V").format.columnWidth = 14;
leads.getRange("A:A").format.columnWidth = 22;
leads.getRange("D:D").format.columnWidth = 22;
leads.getRange("G:G").format.columnWidth = 18;
leads.getRange("J:J").format.columnWidth = 18;
leads.getRange("L:L").format.columnWidth = 38;
leads.getRange("P:P").format.columnWidth = 20;
leads.getRange("R:R").format.columnWidth = 24;
leads.getRange("V:V").format.columnWidth = 24;
leads.getRange("A3:V203").format.wrapText = true;
leads.showGridLines = false;
leads.freezePanes.freezeRows(3);

// Follow-ups
title(followUps, "A1:I1", "Follow-ups｜每次跟進都保留一筆紀錄");
followUps.getRange("A3:I3").values = [["followup_date", "lead_id", "customer_name", "channel", "summary", "customer_response", "next_action", "next_followup_date", "owner"]];
header(followUps, "A3:I3");
followUps.getRange("A4:I4").values = [[new Date("2026-08-14"), "DEMO-20260814-002", "林設計師（DEMO）", "電話", "確認需要全室初估", "下週補照片與尺寸", "週一追蹤資料", new Date("2026-08-18"), "待指派"]];
tableStyle(followUps, "A4:I203");
followUps.getRange("A4:A203").format.numberFormat = "yyyy-mm-dd";
followUps.getRange("H4:H203").format.numberFormat = "yyyy-mm-dd";
followUps.getRange("D4:D203").dataValidation = { rule: { type: "list", values: ["LINE", "電話", "現場", "Email"] } };
followUps.getRange("A:I").format.columnWidth = 18;
followUps.getRange("E:F").format.columnWidth = 30;
followUps.getRange("G:G").format.columnWidth = 24;
followUps.showGridLines = false;
followUps.freezePanes.freezeRows(3);

// Quotes
title(quotes, "A1:K1", "Quotes｜報價、成交與安裝追蹤");
quotes.getRange("A3:K3").values = [["quote_date", "lead_id", "customer_name", "curtain_type", "quoted_amount", "confirmed", "won_amount", "planned_install_date", "actual_install_date", "gross_margin_estimate", "lost_reason"]];
header(quotes, "A3:K3");
quotes.getRange("A4:K4").values = [[new Date("2026-08-14"), "DEMO-20260814-001", "王小美（DEMO）", "遮光布簾", 48000, "未確認", null, null, null, null, ""]];
tableStyle(quotes, "A4:K203");
quotes.getRange("A4:A203").format.numberFormat = "yyyy-mm-dd";
quotes.getRange("H4:I203").format.numberFormat = "yyyy-mm-dd";
quotes.getRange("E4:E203").format.numberFormat = '"NT$"#,##0';
quotes.getRange("G4:G203").format.numberFormat = '"NT$"#,##0';
quotes.getRange("J4:J203").format.numberFormat = "0.0%";
quotes.getRange("F4:F203").dataValidation = { rule: { type: "list", values: ["未確認", "已確認", "已成交", "未成交"] } };
quotes.getRange("A:K").format.columnWidth = 18;
quotes.getRange("D:D").format.columnWidth = 22;
quotes.getRange("K:K").format.columnWidth = 26;
quotes.showGridLines = false;
quotes.freezePanes.freezeRows(3);

// Dashboard
title(dashboard, "A1:H1", "Dashboard｜每週檢視名單與案件漏斗");
dashboard.getRange("A3:B3").values = [["KPI", "目前值"]];
header(dashboard, "A3:B3");
dashboard.getRange("A4:A10").values = [["全部名單"], ["新名單"], ["已聯繫"], ["已約丈量"], ["已報價"], ["已成交"], ["未排程跟進"]];
dashboard.getRange("B4:B10").formulas = [
  ["=COUNTIF('Leads'!$A$4:$A$203,\"<>\")"],
  ["=COUNTIF('Leads'!$N$4:$N$203,\"new_lead\")"],
  ["=COUNTIF('Leads'!$N$4:$N$203,\"contacted\")"],
  ["=COUNTIF('Leads'!$N$4:$N$203,\"measurement_booked\")"],
  ["=COUNTIF('Leads'!$N$4:$N$203,\"quoted\")"],
  ["=COUNTIF('Leads'!$N$4:$N$203,\"won\")"],
  ["=COUNTIF('Leads'!$U$4:$U$203,\"未排程\")"],
];
tableStyle(dashboard, "A4:B10");
dashboard.getRange("B4:B10").format = { fill: mint, font: { bold: true, color: navy, size: 14 }, horizontalAlignment: "right", borders: { preset: "outside", style: "thin", color: lightBorder } };
dashboard.getRange("D3:E3").values = [["意向", "名單數"]];
header(dashboard, "D3:E3");
dashboard.getRange("D4:D6").values = [["高"], ["中"], ["低"]];
dashboard.getRange("E4:E6").formulas = [["=COUNTIF('Leads'!$S$4:$S$203,D4)"], ["=COUNTIF('Leads'!$S$4:$S$203,D5)"], ["=COUNTIF('Leads'!$S$4:$S$203,D6)"]];
tableStyle(dashboard, "D4:E6");
dashboard.getRange("G3:H3").values = [["本週營運檢查", "狀態"]];
header(dashboard, "G3:H3");
dashboard.getRange("G4:H7").values = [
  ["新名單是否 24h 內首次聯繫", "人工檢查"],
  ["是否有未排程跟進", "看 KPI"],
  ["報價後是否有下一步", "人工檢查"],
  ["DEMO 資料是否已刪除", "上線前必做"],
];
tableStyle(dashboard, "G4:H7");
dashboard.getRange("A13:H13").merge();
dashboard.getRange("A13:H13").values = [["使用方式：所有 KPI 都以 Leads 為來源；新增名單後更新 status、intent_level 與 next_followup_date，即可反映在此頁。"]];
dashboard.getRange("A13:H13").format = { fill: sand, font: { color: text }, wrapText: true };
dashboard.getRange("A:B").format.columnWidth = 24;
dashboard.getRange("D:E").format.columnWidth = 18;
dashboard.getRange("G:G").format.columnWidth = 28;
dashboard.getRange("H:H").format.columnWidth = 18;
dashboard.showGridLines = false;

const leadTable = leads.tables.add("A3:V203", true, "LeadsTable");
leadTable.showFilterButton = true;
const followupTable = followUps.tables.add("A3:I203", true, "FollowupsTable");
followupTable.showFilterButton = true;
const quoteTable = quotes.tables.add("A3:K203", true, "QuotesTable");
quoteTable.showFilterButton = true;

const inspect = await wb.inspect({ kind: "table", range: "Dashboard!A1:H13", include: "values,formulas", tableMaxRows: 20, tableMaxCols: 10 });
console.log(inspect.ndjson);
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
console.log(errors.ndjson);

for (const sheetName of ["00_使用說明", "Leads", "Follow-ups", "Quotes", "Dashboard"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${workDir}/${sheetName}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(`${outputDir}/curtain-soft-furnishing-pilot-tracker.xlsx`);
