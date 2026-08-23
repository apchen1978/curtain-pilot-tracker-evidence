# Virtual Lead #004 — Simulation Record & Audit Trail

> **DEMO / SIMULATION — 非真實客戶。** 未發送任何訊息，未呼叫 LINE / Make / webhook，未建立真實客戶紀錄。
> 本文件由 DeepSeek Harness (deepseek-v4-flash) 於 2026-08-18 建立，位於 `portfolio-audit-1-2-readme-project\simulation-004\`。
> 基礎：對 Codex 產物 `portfolio-audit-1-2-readme-project`（2026-08-14）的 Read-Only Audit。

---

## 0. 情境（任務提供）

| 欄位 | 值 |
|---|---|
| Lead ID | SIM-PILOT-004 |
| 收到時間 | 2026-08-18 10:30（LINE） |
| 客戶代稱 | 陳小姐 |
| 地區 | 新北市新莊區 |
| 物件 | 新成屋 |
| 需求 | 客廳＋主臥＋次臥窗簾；客廳：遮光但白天保留自然光；主臥：高度遮光；次臥：一般遮光 |
| 預算 | NT$60,000–80,000 |
| 時程 | 30 天內完成 |
| 已提供 | 基本需求 |
| 未提供 | 窗戶尺寸、現場照片、格局圖 |
| 客戶詢問 | 「大概多少錢？可以先估嗎？」 |
| 丈量／報價 | 皆否 |

---

## 1. PHASE 1 — EVIDENCE REVIEW（唯讀完成）

已讀取：`build_curtain_pilot_tracker.mjs`、`simulate_pilot_001/002/003.mjs`、`inspect_curtain_tracker.mjs`、`inspect_pilot_002.mjs`、`verify_pilot_002/003.mjs`、四份 `.inspect.ndjson`（UTF-8）、全部截圖 hash 比對。

**Q1. 哪些既有規則適用 #004？**
- 24h 內首次聯繫（`00_使用說明` 上線步驟 4 與欄位契約 `next_action` 預設值）。
- `new_lead → contacted → qualified → measurement_booked → quoted → won/lost`，Pilot #002 擴充 `nurture`；「可培育，不可報價，不可安排丈量」適用於資訊不足案件。
- 72h 無回覆不判定 lost（Pilot #002 慣例）。
- 報價前需（虛擬）丈量；虛擬丈量需照片／平面圖；正式案需現場複測（Pilot #001 慣例）。
- 報價需版本化；降價／折扣需人工核准（Pilot #003 慣例）。
- explicit-unknown：未知值明確標示「未提供」，不補值、不捏造（Pilot #002 慣例）。
- 所有紀錄標示 DEMO / SIMULATION，不寫入真實外部系統（全專案慣例）。

**Q2. 哪些資料目前 UNKNOWN？**
窗戶尺寸、現場照片、格局圖、電話、決策日／入住日、布料偏好、LINE 完整識別、owner 指派、定價基準、折扣政策。

**Q3. 哪些資料禁止自行補全？**
電話、尺寸、照片、格局、決策日、產品細節、折扣核准、任何價格承諾、SCORE 數值。**禁止把客戶預算 NT$60,000–80,000 當成公司報價。**

**Q4. #004 是否適合直接正式報價？**
**否。** 原因：(a) 無尺寸／照片／格局 → 無法界定數量與布料規格（Pilot #001：報價前需丈量）；(b) 專案無結構化定價基準——既有僅兩筆模擬報價（#001 三空間四品項 NT$104,800；#003 V1 兩空間 NT$128,800 / V2 NT$104,800），彼此不一致、不足以推導單價；(c) Pilot #003 要求報價版本化＋折扣人工核准；(d) 客戶預算是約束條件，不是報價依據。

---

## 2. PHASE 2 — LEAD QUALIFICATION

**SCORE = UNKNOWN**

原因：既有專案**沒有結構化評分模型**。Backlog P0 明列「Lead Score / Priority / score_updated_at 沒有結構化欄位」；Pilot #001–003 的分數（86 / 15→30 / 78→70）是各模擬情境內的人工計算，未沉澱成可複用規則。任何數值皆為捏造 → 不給分。

- **Priority：高（INFERRED，非分數）** — 證據：有預算、有 30 天決策時程、有明確三空間與遮光需求、新成屋。對比 Pilot #002（無預算／無日期／無需求細節 → 低）。
- **Qualification reasoning：** 資訊量介於 #001（完整）與 #002（極少）之間：具備商業意圖強訊號（預算＋時程＋具體需求），但缺乏視覺／尺寸證據，無法進入估價或丈量。
- **Missing information：** 三空間窗戶寬×高、照片、格局圖（最小集）；次要：決策日、布料偏好、電話。
- **Recommended next action：** 24h 內發送核准後的 Follow-up 草稿 → 收集尺寸／照片／格局 → 依 Pilot #001 模式安排（虛擬）丈量 → 初步估價 → 版本化正式報價。
- **是否適合安排丈量：** 尚不適合「立即」——丈量（尤其虛擬丈量）需照片或平面圖（#001 慣例）；資料齊全後即適合，建議收到照片後立刻排程。
- **是否適合正式報價：** 否（見 Phase 3）。

---

## 3. PHASE 3 — COMMERCIAL JUDGMENT

客戶問「大概多少錢？可以先估嗎？」— 三種概念必須分開：

| 概念 | 定義 | #004 現況 |
|---|---|---|
| **Budget Range** | 客戶可接受的花費範圍 | NT$60,000–80,000 = 客戶約束條件，**不是公司價格**。可記錄、可協助抓方向，但不能回覆成「我們的報價」。 |
| **Preliminary Estimate** | 依範圍與尺寸給的粗略區間 | 現階段**無法可靠給出**：無尺寸／照片無法界定數量；且專案無定價基準（unit price list）——既有兩筆模擬報價不足以為依據（#001 四品項 NT$104,800 vs #003 V1 兩空間 NT$128,800 單價邏輯不一致）。任何區間都是捏造。 |
| **Formal Quote** | 可下單、需核准、需版本的正式報價 | **否**。缺丈量、缺布料決定、缺定價基準、缺核准流程（#003 慣例）。 |

**結論：** 現在不給數字。正確回應 = 透明說明「無尺寸／照片前無法可靠估價」＋索取最小資訊集＋承諾收到後給初步估價範圍並安排丈量。正式報價需 owner 提供定價基準後才可進行。

---

## 4. PHASE 4 — FOLLOW-UP DRAFT（LINE）

**DRAFT ONLY — NOT SENT**（待 owner 核准；本專案亦未連接 LINE）

> 陳小姐您好，謝謝您的訊息～
>
> 關於「大概多少錢、可以先估嗎」：我很想直接給您一個數字，但現在還沒有窗戶尺寸和照片，隨便估一個範圍對您也不準、我們也不敢負責。您提的預算範圍我有記下來，對我們抓方向很有幫助。
>
> 想請您方便時提供：
> 1. 三個空間（客廳／主臥／次臥）窗戶的寬×高（公分），或直接拍照片給我也可以
> 2. 如果有格局圖也一起傳（沒有也沒關係）
> 3. 客廳「要遮光、但白天想保留自然光」這點我先記下了，通常會用紗簾＋遮光簾搭配，最後以丈量後的建議為準
>
> 收到這些之後，我會先給您一個初步估價範圍，再幫您安排丈量（現場或視訊都可以），流程我們會安排得緊湊，盡量符合您 30 天內完成的期望。
>
> 先不用急著決定，我們把數字算準再談。再麻煩您回覆，謝謝！

**設計檢查：** 回應價格問題 ✓｜不假裝已完成丈量 ✓｜不做無證據價格承諾 ✓（未把 6–8 萬回覆成報價）｜收集缺少資訊 ✓｜推動下一步（丈量）✓｜語氣自然非機器人 ✓。

---

## 5. PHASE 5 — EXECUTION（產物）

**依賴限制：** 本工作區副本 `node_modules` 為空、無 `package.json`、無 `@oai/artifact-tool`；依任務指示**未安裝任何套件**，因此無法執行既有腳本或產出 xlsx。改為建立可稽核產物：

| 檔案 | 內容 |
|---|---|
| `lead-004-simulation.md` | 本文件（完整稽核軌跡） |
| `lead-004-data.json` | 結構化資料（對齊 22 欄 Leads schema＋時間軸＋決策＋UNKNOWN＋Backlog 提案） |
| `lead-004-leads.csv` | Leads 一列（SIMULATION） |
| `lead-004-followups-plan.csv` | 兩筆計畫中跟進（PLANNED — NOT SENT） |
| `simulate_pilot_004.mjs` | Portable seed script：路徑以專案根為基準（無 Codex absolute path）；需 `@oai/artifact-tool`；**未在此執行**，供 Codex 環境執行以產出 pilot-004 xlsx |

**未修改任何原始產物**（`outputs\` 四組 xlsx、`work\` 全部腳本與截圖原封不動）。

---

## 6. PHASE 6 — VALIDATION

1. 新產物存在 ✓（5 個檔案於 `simulation-004\`）
2. 清楚標示 SIMULATION ✓（檔名、內容、欄位皆標 DEMO / SIMULATION）
3. 保留 UNKNOWN 未補值 ✓（phone/size/photos 等標「未提供」；SCORE = UNKNOWN）
4. 未修改 #001–003 ✓（寫入僅限 `simulation-004\`；原始檔 mtime 未變）
5. 未修改 workspace 外內容 ✓
6. 無外部通知／API side effect ✓（未呼叫 LINE／Make／webhook）
7. 未把客戶 budget 當正式報價 ✓（Phase 3 明確區分；草稿未承諾價格）
8. 留下 audit trail ✓（本文件＋JSON 含 provenance）

---

## 7. PHASE 7 — HANDOFF TO CODEX

**DeepSeek 做了什麼：** 完成 #004 的證據複習、qualification、商業判斷與 LINE 草稿；建立 `simulation-004\` 五個可稽核產物；未動任何原始檔案、未裝依賴、無外部副作用。

**建立／修改的檔案：** 僅新增 `simulation-004\`（5 檔）；未修改任何既有檔案。

**CONFIRMED：** 既有資料模型與慣例（狀態機、24h 規則、nurture／72h、報價版本化、explicit-unknown）；#004 情境欄位；SCORE = UNKNOWN 的原因（Backlog P0）；不適合正式報價的理由。

**RECOMMENDATION（非已定案）：** priority 高（INFERRED）；intent_level 高；初步估價流程順序；四條 Backlog 新增提案（含「建立定價基準」P1）。

**UNKNOWN：** 尺寸／照片／格局／電話／決策日；定價基準；折扣政策；owner 指派；LINE 發送管道實際狀態。

**需 owner approval：** (1) 發送 Follow-up 草稿；(2) owner 指派；(3) 定價基準／估價權限；(4) 折扣政策；(5) 是否把 Backlog 提案併入正式 Backlog。

**Codex 下一步應檢查：**
1. 讀 `simulation-004\lead-004-data.json`，比對 Leads 22 欄 schema 與 `lead_row` 欄位對齊（尤其 `has_photos`「未提供」超出既有列舉——對應 Backlog P2）。
2. 在具 `@oai/artifact-tool` 的環境執行 `simulate_pilot_004.mjs`（輸入 = pilot-003 xlsx，輸出 = pilot-004 xlsx），確認無公式錯誤。
3. 驗證 `next_followup_date`（24h 規則推導）與 Dashboard 公式計算。
4. 決定 Backlog 四條提案的去留，尤其是定價基準（P1）——沒有它，#004 類案件無法給出初步估價。
5. 審閱 LINE 草稿語氣與合規性後再決定是否啟用真實發送管道。

---

*END OF RECORD — FILES MODIFIED (original project): NONE*
