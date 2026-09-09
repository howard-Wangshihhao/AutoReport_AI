# AI Security Test Report Generator

此工具用於依據 Excel 中已由檢測人員填寫的「判定結果」與「目前情況」，自動產生資安檢測報告文字、CVSS v3.1 Base Vector / Score，並依檢測大分類產生整體改善建議。

目前支援 OpenAI 與 Gemini API，並可於 Windows、Linux、macOS 執行。程式只使用 Python 套件，不需要安裝 Microsoft Excel 或 LibreOffice。

---

## 1. 主要功能

### 測項層級

針對 `IOT Device` 工作表逐列處理：

- `Report`：英文正式檢測報告
- `Report_ch`：繁體中文正式檢測報告
- `cvss`：CVSS v3.1 Base Vector
- `score`：CVSS v3.1 Base Score
- 依測項名稱套用較適合的報告敘述模板
- 可解析「參考 P06」等跨測項引用資訊
- 人工「判定結果」與 CVSS Severity 不一致時，`cvss` 與 `score` 會標示為紅字

### 分類層級

程式會依 Excel 的 `分類` 欄位彙整測項，並自動建立：

```text
Security Recommendation
```

此工作表欄位如下：

| 欄位 | 說明 |
|---|---|
| 分類 | Static、Dynamic、Fuzzing、PT 等大分類 |
| 最高風險 | 該分類目前最高人工判定風險 |
| 觸發測項 | 該分類中 Low Risk 以上的測項編號 |
| Recommendation | 英文整體改善建議 |
| Recommendation_ch | 繁體中文整體改善建議 |

只有分類中存在 `Low` 以上的測項時，才會呼叫 AI 產生分類改善建議。

---

## 2. 建議專案結構

```text
AutoReport_AI/
├─ generate_report_ai.py
├─ .env
├─ README.md
└─ check_iot_ai_template.xlsx
```

執行後預設產生：

```text
check_iot_ai_template_report.xlsx
```

原始 Excel 不會被覆蓋。

---

## 3. 安裝環境

建議使用 Python 3.10 以上版本。

### 建立虛擬環境

Windows PowerShell：

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

Linux / macOS：

```bash
python3 -m venv venv
source venv/bin/activate
```

### 安裝套件

```bash
pip install openpyxl python-dotenv openai google-genai
```

若只使用其中一個 AI Provider，也可只安裝對應 SDK。

---

## 4. `.env` 設定

`.env` 必須與 `generate_report_ai.py` 放在同一個資料夾。

### OpenAI

```env
AI_PROVIDER=openai
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=your_model_name_here
OPENAI_REASONING=

SHEET_NAME=IOT Device
REPORT_LANGUAGE=English
OVERWRITE_REPORT=true
```

### Gemini

```env
AI_PROVIDER=gemini
GEMINI_API_KEY=your_api_key_here
GEMINI_MODEL=your_model_name_here

SHEET_NAME=IOT Device
REPORT_LANGUAGE=English
OVERWRITE_REPORT=true
```

說明：

- `AI_PROVIDER`：`openai` 或 `gemini`
- `OPENAI_API_KEY` / `GEMINI_API_KEY`：API Key
- `OPENAI_MODEL` / `GEMINI_MODEL`：要使用的模型名稱
- `OPENAI_REASONING`：選填；留白時不額外指定 reasoning effort
- `SHEET_NAME`：主要檢測工作表，預設為 `IOT Device`
- `REPORT_LANGUAGE`：目前保留供相容性使用；實際輸出固定為英文與繁體中文兩欄
- `OVERWRITE_REPORT=true`：允許重新產生既有 Report / CVSS 欄位

---

## 5. Excel 必要欄位

`IOT Device` 工作表至少需要以下欄位：

| 欄位 | 用途 |
|---|---|
| 分類 | 用於 Static / Dynamic / Fuzzing / PT 等大項彙整 |
| 編號 | 測項編號，例如 P06 |
| 測項 | 測項名稱 |
| 判定結果 | 檢測人員人工決定的風險結果 |
| 目前情況 | 實際測試結果、觀察內容與風險降低因素 |
| Report | AI 產生的英文正式報告 |
| Report_ch | AI 產生的繁體中文正式報告 |
| cvss | CVSS v3.1 Base Vector |
| score | CVSS v3.1 Base Score |

以下欄位為選用：

```text
修補建議
```

若有填寫，AI 可將其整理至報告與分類改善建議中；若未提供，AI 不應自行創造具體修補措施。

---

## 6. 支援的判定結果

程式目前可辨識下列寫法。

### Not Applicable

```text
Not Applicable
not applicable
N/A
NA
不適用
```

輸出：

```text
cvss  = -
score = -
```

### No Risk

```text
None
none
no risk
no finding
無風險
未發現風險
```

輸出：

```text
cvss  = CVSS:3.1/AV:-/AC:-/PR:-/UI:-/S:-/C:N/I:N/A:N
score = 0.0
```

> 注意：上述 `AV:-/AC:-/...` 為本報告工具的自訂「無風險」表示方式，不是正式 CVSS v3.1 Base Vector 的合法值。

### Risk

```text
Low / low risk / 低風險
Medium / medium risk / 中風險
High / high risk / 高風險
Critical / critical risk / 嚴重風險 / 重大風險
```

Low 以上會由 AI 根據測項與「目前情況」判斷以下 CVSS v3.1 Base Metrics：

```text
AV / AC / PR / UI / S / C / I / A
```

Python 再依 CVSS v3.1 Base Score 公式計算分數。

### 不處理的狀態

以下狀態會直接略過：

```text
TBD
Testing
```

---

## 7. Report 產生邏輯

### Not Applicable

結構：

```text
不適用原因
→ 不在本次測試範圍
→ 固定結論
```

例如：

```text
此測項因廠商未提供原始碼，故不在本次測試範圍內，因此，此測項評估為不適用。
```

### No Risk

結構：

```text
測試方式 / 觀察方式
→ 測試結果
→ 固定結論
```

### Low / Medium / High / Critical

結構：

```text
發現的風險
→ 目前情況中已存在的降低或限制風險因素
→ 必要時整理修補建議
→ 固定風險結論
```

降低風險的原因只能來自 Excel 的 `目前情況` 或明確引用的參考測項，不可由 AI 自行創造。

---

## 8. 測項類型模板

除了判定結果之外，程式也會依 `測項` 名稱選擇適合的報告描述風格，目前包含：

- 軟體 / 原始碼 / SBOM / 元件弱點分析
- 應用程式與輸入處理安全
- 身分驗證與密碼管理
- 網路與通訊安全
- 敏感資料與隱私保護
- 安全事件日誌與通知
- 更新與軟體交付保護
- 實體與一般裝置安全
- General fallback

此分類只影響文字描述方式，不會自行新增不存在的測試事實。

---

## 9. 跨測項引用

若 `目前情況` 中寫：

```text
參考 P06
```

或：

```text
參考P06
refer to P06
```

程式會自動取得 P06 的：

- 測項名稱
- 判定結果
- 目前情況

並一併提供給 AI 作為補充上下文。

此機制不會改變本列人工判定結果。

---

## 10. CVSS 與人工判定不一致

人工判定結果與 CVSS v3.1 Base Severity 可能不同。

例如：

```text
人工判定：Low
CVSS Score：4.3
CVSS Severity：Medium
```

程式仍會保留：

```text
判定結果 = Low
```

但會：

1. 在命令列顯示 `CVSS WARNING`
2. 將該列 `cvss` 字體標紅
3. 將該列 `score` 字體標紅

例如：

```text
[CVSS WARNING] P06: CVSS v3.1 計算結果為 4.3 (medium)，與人工判定結果 low risk 不一致；已保留人工判定並輸出 CVSS Base Score，請人工複核
```

CVSS Base Severity 區間：

| Severity | Score |
|---|---:|
| None | 0.0 |
| Low | 0.1–3.9 |
| Medium | 4.0–6.9 |
| High | 7.0–8.9 |
| Critical | 9.0–10.0 |

---

## 11. Security Recommendation

完成逐項 Report / CVSS 後，程式會再依 `分類` 進行彙整。

例如：

```text
PT
├─ P01 None
├─ P06 Low
├─ P08 Low
├─ P09 Low
└─ P10 None
```

則會產生：

```text
分類：PT
最高風險：low risk
觸發測項：P06, P08, P09
```

並將 P06、P08、P09 的資訊一起交由 AI 產生一段「分類層級」的整體改善建議。

### 分類建議方向

#### Static

優先考量：

- SSDLC
- Source Code Security Testing
- Code Review
- SAST / SCA
- SBOM
- 第三方元件弱點管理
- 修補追蹤
- Release Security Gate

#### Dynamic

優先考量：

- Dynamic Security Testing
- Runtime Behavior
- Release Regression Testing
- Abnormal Behavior Validation

#### Fuzzing

優先考量：

- Fuzz Testing
- Protocol / Parser
- API
- Boundary Input
- Abnormal Input
- Regression Testing

#### PT

依實際 Low 以上測項聚焦：

- Attack Surface
- Network / Communication Security
- Authentication / Authorization
- Sensitive Data Protection
- Update Security
- Physical Security

不會將與實際 Finding 無關的控制措施全部列入。

---

## 12. 執行方式

### 一般執行

```bash
python generate_report_ai.py "check_iot_ai_template.xlsx"
```

預設輸出：

```text
check_iot_ai_template_report.xlsx
```

### 指定輸出檔

```bash
python generate_report_ai.py "check_iot_ai_template.xlsx" --output "result.xlsx"
```

### 只做 Excel 標準化

```bash
python generate_report_ai.py "check_iot_ai.xlsx" --normalize-only
```

此模式：

- 不呼叫 AI
- 不產生 Report
- 將 x14 Extended Data Validation 轉為標準 DataValidation

預設輸出：

```text
check_iot_ai_report.xlsx
```

---

## 13. 執行結果範例

```text
[11] P06 Network Security Test (Adjacent) (Low) ...
  [CVSS WARNING] P06: CVSS v3.1 計算結果為 4.3 (medium)，與人工判定結果 low risk 不一致；已保留人工判定並輸出 CVSS Base Score，請人工複核
[13] P08 Sensitive Data Leakage (Low) ...
[14] P09 API Authentication (Low) ...

完成
產生 Report：21 筆
略過：3 筆
失敗：0 筆
產生分類建議：1 類
輸出檔：...\check_iot_ai_template_report.xlsx
```

---

## 14. 常見問題

### Q1. 為什麼判定結果是 Low，但 CVSS 是 Medium？

人工「判定結果」屬於檢測人員的整體風險判定；CVSS Base Score 則描述漏洞本身的 Base Severity。兩者可能因產品情境、既有控制措施或風險接受因素而不同。

程式不會為了配合人工判定而刻意修改 CVSS Metrics，而是以紅字提醒人工複核。

### Q2. 為什麼 Report 沒有產生？

確認：

- `判定結果` 是否有值
- `目前情況` 是否有值
- 判定結果是否為 `TBD` 或 `Testing`
- API Key / Model 是否正確設定
- 命令列是否出現 `[ERROR]`

### Q3. 為什麼 CVSS 空白但 Report 有內容？

Report 與 CVSS 已分開處理。若 AI 報告成功，但 CVSS Base Metrics 無法完整判斷，程式仍會保留 Report，並在命令列顯示：

```text
[CVSS ERROR]
```

### Q4. `Security Recommendation` 為什麼沒有建議？

只有該分類中存在 `Low / Medium / High / Critical` 測項時才產生建議。

若分類全部為 `None / Not Applicable / TBD / Testing`，Recommendation 會保持空白。

### Q5. 為什麼每次執行 Security Recommendation 都重新產生？

目前程式會重建 `Security Recommendation` 工作表，以確保內容與最新人工判定及 Report 結果一致。

---

## 15. 使用原則

本工具定位為「檢測報告撰寫與 CVSS 輔助判讀工具」。

建議維持以下原則：

1. `判定結果` 由檢測人員人工決定。
2. `目前情況` 應記錄可被報告引用的實際測試事實。
3. AI 不應自行改變人工風險判定。
4. AI 產出的 Report、CVSS Metrics 與改善建議仍應由檢測人員複核。
5. CVSS 與人工判定不一致時，應確認 Base Metrics 是否正確，而不是單純為配合風險等級調整分數。
6. 分類改善建議屬於後續改善方向，不代表所有建議措施都一定適用於受測產品。

---

## 16. 目前輸出成果

一般執行完成後，主要會得到：

```text
IOT Device
├─ Report
├─ Report_ch
├─ cvss
└─ score

Security Recommendation
├─ 分類
├─ 最高風險
├─ 觸發測項
├─ Recommendation
└─ Recommendation_ch
```

此架構可同時保留「逐測項 Finding」與「分類層級改善建議」，方便後續整理為正式資安檢測報告。
