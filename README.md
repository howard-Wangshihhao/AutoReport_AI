# AutoReport AI

以 Excel 中人工填寫的「判定結果」與「目前情況」為基礎，呼叫 OpenAI 或 Gemini 產生資安檢測報告文字，並自動回填英文 `Report`、中文 `Report_ch`、CVSS v3.1 Vector、Score，以及分類層級的 `Security Recommendation`。

本工具設計原則為：**人工判定結果優先、AI 僅協助撰寫與 CVSS Base Metrics 判讀，不自行改變風險等級，也不應編造未提供的測試事實。**

---

## 1. 功能摘要

- 讀取既有 `.xlsx` 檢測表。
- 依 `判定結果`、`目前情況`、`修補建議` 產生：
  - 英文 `Report`
  - 繁體中文 `Report_ch`
- 支援 OpenAI 與 Google Gemini。
- 對 Low / Medium / High / Critical 項目產生 CVSS v3.1 Base Metrics。
- CVSS 第一輪資訊不完整時，自動進行一次專項重試。
- 若兩次都無法合理判定 CVSS，不硬編 Metrics：
  - `cvss = N/A`
  - `score = -`
- 對 None / No Risk 項目固定輸出 0.0。
- 人工風險等級與 CVSS Severity 不一致時：
  - 保留人工判定結果
  - `cvss` 與 `score` 以紅字顯示，提醒人工複核
- `Report / Report_ch / cvss / score` 統一：
  - 字級 12 pt
  - 垂直置中
  - 自動換行
- 依內容長度自動調整 Excel 列高。
- 可建立 `Security Recommendation` 工作表，依大分類彙整 Low 以上項目並產生整體改善建議。
- 支援 Excel x14 清單型 Data Validation 的轉換，避免 openpyxl 讀寫後下拉選單遺失。
- 不需要安裝 Microsoft Excel 或 LibreOffice。

---

## 2. 執行環境

建議使用 Python 3.10 以上版本。

安裝共用套件：

```powershell
pip install openpyxl python-dotenv
```

使用 OpenAI：

```powershell
pip install openai
```

使用 Gemini：

```powershell
pip install google-genai
```

也可以一次安裝：

```powershell
pip install openpyxl python-dotenv openai google-genai
```

---

## 3. 專案檔案

最簡單的目錄結構如下：

```text
AutoReport_AI/
├─ generate_report_ai.py
├─ .env
└─ check_iot_ai.xlsx
```

---

## 4. `.env` 設定

`.env` 請放在 Python 程式相同目錄。

### 使用 Gemini

```env
AI_PROVIDER=gemini
GEMINI_API_KEY=你的_API_Key
GEMINI_MODEL=你的_Gemini_模型名稱

SHEET_NAME=IOT Device
REPORT_LANGUAGE=English
OVERWRITE_REPORT=true
```

### 使用 OpenAI

```env
AI_PROVIDER=openai
OPENAI_API_KEY=你的_API_Key
OPENAI_MODEL=你的_OpenAI_模型名稱

# 選填，例如 medium
OPENAI_REASONING=

SHEET_NAME=IOT Device
REPORT_LANGUAGE=English
OVERWRITE_REPORT=true
```

> `REPORT_LANGUAGE` 為相容既有設定保留。程式目前固定同時產生英文 `Report` 與繁體中文 `Report_ch`。

### OVERWRITE_REPORT

```env
OVERWRITE_REPORT=true
```

表示重新執行時可覆寫既有 AI 產出的內容。

若設為：

```env
OVERWRITE_REPORT=false
```

當該列的 Report、Report_ch、cvss、score 已完整存在時，會略過該列。

---

## 5. Excel 必要欄位

主工作表預設名稱：

```text
IOT Device
```

至少需要以下欄位：

| 欄位 | 用途 |
|---|---|
| 編號 | 測項編號，例如 P06、S01 |
| 測項 | 測試項目名稱 |
| 判定結果 | 人工決定的最終風險結果 |
| 目前情況 | 測試方式、測試結果、攻擊條件等事實 |
| Report | AI 產生的英文報告 |
| Report_ch | AI 產生的繁體中文報告 |
| cvss | CVSS v3.1 Vector |
| score | CVSS Base Score |

建議另外保留：

| 欄位 | 用途 |
|---|---|
| 分類 | 用於 Security Recommendation 分類彙整 |
| 修補建議 | 人工提供的改善方向，AI 可納入報告 |

---

## 6. 建議的「目前情況」格式

為了讓 AI 更容易區分測試事實、攻擊條件與降低風險因素，建議使用固定結構：

```text
測試方式：
- ...

測試結果：
- ...

攻擊條件：
- ...

可能影響：
- ...

既有防護／限制條件：
- ...
```

原則：

- 一個 bullet 只描述一個事實。
- 不要把人工風險結論混在「目前情況」裡。
- 不確定的資訊不要推測。
- 沒有資料可寫「原始資料未提供」或「本次未進一步評估」。

例如：

```text
測試方式：
- 針對 Bluetooth 配對與資料傳輸流程進行測試。

測試結果：
- 裝置採用 Just Works 配對方式，未使用 PIN/Passkey。

攻擊條件：
- 攻擊者需位於 Bluetooth 通訊範圍內。
- 配對功能需由合法使用者按下實體按鈕後才會啟用。

可能影響：
- 配對期間可能存在未經充分身分驗證之連線風險。

既有防護／限制條件：
- Bluetooth 非持續開啟。
- 敏感資料傳輸已有加密保護。
```

---

## 7. 判定結果處理

程式目前可辨識：

- Critical
- High
- Medium
- Low
- None / No Risk
- Not Applicable / N/A
- TBD
- Testing

其中：

- `TBD`、`Testing`：略過，不產生 Report。
- `Not Applicable`：產生 N/A 類型報告，不要求 CVSS Base Metrics。
- `None` / `No Risk`：產生無風險報告，score 固定為 0.0。
- `Low` 以上：產生 Report 並進行 CVSS v3.1 Base Metrics 判讀。

人工填寫的 `判定結果` 是最終檢測結論，AI 不得自行改成其他風險等級。

---

## 8. CVSS v3.1 規則

程式僅處理 CVSS v3.1 Base Metrics：

```text
AV / AC / PR / UI / S / C / I / A
```

分數由 Python 依 Base Metrics 計算，不採用 AI 直接提供的 score。

### 8.1 Low 以上

AI 第一輪會嘗試輸出完整 Metrics。

若第一輪 Metrics 缺漏或格式錯誤：

```text
[CVSS RETRY] Pxx: 第一輪 CVSS 不完整，重新判讀一次
```

程式會再做一次專門的 CVSS 判讀。

### 8.2 資訊仍不足

若第二次仍無法合理辨識具體漏洞或攻擊情境，不會硬湊 CVSS：

```text
[CVSS N/A] S01: 目前資訊不足以辨識可供 CVSS v3.1 Base Metrics 評估的具體漏洞或攻擊情境
```

Excel 會寫入：

```text
cvss  = N/A
score = -
```

這種情況常見於：

- 僅表示測試資料尚未提供完整。
- 僅表示測試範圍不完整。
- 沒有具體漏洞。
- 沒有攻擊路徑。
- 沒有可以判斷 C / I / A 的實際影響。

### 8.3 None / No Risk

程式目前使用：

```text
CVSS:3.1/AV:-/AC:-/PR:-/UI:-/S:-/C:N/I:N/A:N
```

score：

```text
0.0
```

> 注意：上述 `-` 並不是 CVSS v3.1 官方 Base Vector 的標準 metric value，而是本工具為報告呈現所採用的自訂「無風險」表示方式。

### 8.4 Not Applicable

```text
cvss  = -
score = -
```

---

## 9. 人工風險與 CVSS 不一致

人工判定結果與 CVSS Base Severity 是兩個不同概念。

例如人工判定：

```text
Low
```

但 CVSS Base Score：

```text
4.3 Medium
```

程式不會為了讓 CVSS 配合人工 Low 而修改 Metrics，而會顯示：

```text
[CVSS WARNING] P06: CVSS v3.1 計算結果為 4.3 (medium)，與人工判定結果 low risk 不一致；已保留人工判定並輸出 CVSS Base Score，請人工複核
```

Excel 中：

- 人工 `判定結果` 不變。
- `cvss` 與 `score` 以紅字提醒。

這是預期行為，不代表程式錯誤。

---

## 10. CVSS 與環境控制的區分

CVSS Base Metrics 應描述漏洞本身的固有特性。

例如以下環境控制通常不應直接拿來降低 Base Metrics：

- Firewall
- ACL
- IP 白名單
- 額外網路隔離
- 特定部署限制

這些因素可以作為人工整體風險判定的降低因素，但不一定改變 CVSS Base Score。

另外，如果漏洞利用前必須由「攻擊者以外的合法使用者」先做必要操作，例如：

```text
必須由合法使用者按下實體按鈕後，才開啟 Bluetooth pairing
```

程式的 CVSS 重試 Prompt 會優先考慮：

```text
UI:R
```

而不是用 `AC:H` 取代使用者互動條件。

---

## 11. Report 格式

程式完成後會統一下列四個欄位：

```text
Report
Report_ch
cvss
score
```

格式為：

- 12 pt
- 垂直置中
- 自動換行

其中：

- `Report` 會先沿用 `Report_ch` 的儲存格格式，保持中英文版面一致。
- `cvss` / `score` 僅調整字級與對齊，不覆蓋原本字色。
- 因此 CVSS mismatch 的紅字會保留。

列高會依 `目前情況`、`修補建議`、`Report`、`Report_ch` 的內容長度自動增加。

---

## 12. 參考其他測項

若「目前情況」中寫：

```text
參考 P06
```

或：

```text
參考P06
refer to P06
```

程式會把 P06 的：

- 測試項目
- 判定結果
- 目前情況

一併提供給 AI 作為補充背景。

但參考測項只提供上下文，不得改變本列人工判定結果。

---

## 13. Security Recommendation

程式會建立／重建：

```text
Security Recommendation
```

工作表欄位：

| 欄位 | 說明 |
|---|---|
| 分類 | 例如 Static、Dynamic、Fuzzing、PT |
| 最高風險 | 此分類中目前最高人工風險 |
| 觸發測項 | Low 以上的測項編號 |
| Recommendation | 英文整體改善建議 |
| Recommendation_ch | 中文整體改善建議 |

只有 Low / Medium / High / Critical 的測項會觸發 AI 產生分類建議。

建議是「分類層級」的整體改善方向，而不是逐筆複製 Report。

目前主要方向：

- **Static**：SSDLC、Code Review、SAST/SCA、SBOM、第三方元件弱點管理、修補追蹤、release security gate。
- **Dynamic**：動態測試、Runtime 行為、版本發布前回歸測試、異常行為驗證。
- **Fuzzing**：Protocol、Parser、API、邊界輸入、Malformed Data、持續 Fuzzing 與回歸測試。
- **PT**：依實際弱點聚焦 Attack Surface、Communication Security、Authentication、Authorization、Sensitive Data、Update、Physical Security 等。

---

## 14. 執行方式

一般執行：

```powershell
python generate_report_ai.py check_iot_ai.xlsx
```

若你的檔名為：

```text
generate_report_ai.py
```

則執行：

```powershell
python generate_report_ai.py check_iot_ai.xlsx
```

預設輸出：

```text
Navi_report.xlsx
```

### 自訂輸出檔名

```powershell
python generate_report_ai.py check_iot_ai.xlsx --output Navi_result.xlsx
```

### 只處理 Excel x14 下拉選單

```powershell
python generate_report_ai.py check_iot_ai.xlsx --normalize-only
```

此模式：

- 不呼叫 AI
- 不產生 Report
- 只將 x14 Data Validation 轉為標準 Excel Data Validation

預設輸出：

```text
Navi_template.xlsx
```

---

## 15. 執行畫面範例

正常處理：

```text
[2] S01 Source code scanning (Low) ...
[3] S02 SBOM Vulnerability Scanning (High) ...
[11] P06 Network Security Test (Adjacent) (Low) ...
```

CVSS 重試：

```text
[CVSS RETRY] S01: 第一輪 CVSS 不完整，重新判讀一次
```

CVSS 不適用／資訊不足：

```text
[CVSS N/A] S01: 目前資訊不足以辨識可供 CVSS v3.1 Base Metrics 評估的具體漏洞或攻擊情境
```

CVSS 與人工風險不一致：

```text
[CVSS WARNING] P06: CVSS v3.1 計算結果為 4.3 (medium)，與人工判定結果 low risk 不一致；已保留人工判定並輸出 CVSS Base Score，請人工複核
```

完成時會顯示類似：

```text
完成
產生 Report：21 筆
略過：3 筆
失敗：0 筆
CVSS 不適用／資訊不足：1 筆
產生分類建議：2 類
```

---

## 16. 常見問題

### Q1. 為什麼 Low Risk 的 CVSS 是 N/A？

因為人工 Low Risk 不一定代表已發現一個可以用 CVSS 評分的具體漏洞。

例如：

```text
尚未取得完整 plugin source code，因此目前檢測範圍不完整
```

這是一個人工評估上的風險或限制，但未必存在可以判定 AV / AC / PR / UI / S / C / I / A 的具體漏洞。

因此程式會保留人工 `Low`，但將：

```text
cvss = N/A
score = -
```

避免產生假的 CVSS。

### Q2. 為什麼人工 Low，但 CVSS 是 Medium？

人工整體風險可以考慮實際產品情境、既有控制、暴露程度與測試限制；CVSS Base Score 則主要描述漏洞本身的固有嚴重度。

因此兩者不一致是允許的，程式會以紅字提醒人工複核。

### Q3. 為什麼某列完全沒產生？

請確認：

- `判定結果` 是否有值。
- `目前情況` 是否有值。
- 判定結果是否為 `TBD` 或 `Testing`。
- `OVERWRITE_REPORT=false` 時，該列是否已完整存在 Report / Report_ch / cvss / score。

### Q4. Gemini 為什麼曾出現 AFC warning？

程式目前已針對 Google GenAI SDK 的：

```text
Direct use of automatic function calling (AFC) in Models.generate_content...
```

提示進行過濾，避免在一般文字生成情境下干擾命令列輸出。

### Q5. Excel 下拉選單為什麼需要 normalize？

部分 Excel 使用 x14 Extended Data Validation。openpyxl 無法完整保留該延伸格式，因此程式會先讀取規則，再重建成標準 Data Validation，提高跨環境相容性。

---

## 17. 使用原則

本工具定位為：

> **資安檢測報告文字與 CVSS Base Metrics 的輔助產生工具，而不是自動取代檢測工程師判定的工具。**

建議正式交付前仍由檢測人員確認：

- 判定結果是否正確。
- Report 是否忠實反映實際測試。
- AI 是否誤解「目前情況」。
- CVSS Metrics 是否符合實際攻擊條件。
- CVSS 與人工風險不一致是否有合理依據。
- Recommendation 是否符合產品實際改善方向。

尤其是 High / Critical、醫療器材、重要系統或其他高風險產品，CVSS 與報告內容應進行人工複核。

---

## 18. 資料安全提醒

若使用雲端 AI Provider，Excel 中送往模型的內容可能包含：

- 測項名稱
- 判定結果
- 目前情況
- 修補建議
- 被引用測項內容

因此不建議直接放入：

- 密碼
- API Key
- Token
- 私鑰
- 未遮罩的個資
- 尚未核准上傳雲端的機敏程式碼或產品資訊

若資料涉及客戶機密或受限制資訊，應先依組織政策確認是否可送至所選 AI Provider。

