# AutoReport AI

AutoReport AI 是一套以 Excel 為輸入的資安檢測報告輔助產生工具，目前分為 **IoT 裝置檢測** 與 **OWASP MASTG 行動 App 檢測（Android / iOS）** 兩條流程。

工具以檢測人員人工填寫的 **「判定結果」**、**「目前情況」** 與 **「修補建議」** 為主要依據，呼叫 AI 產生英文 `Report`、繁體中文 `Report_ch`、CVSS v3.1 Base Vector / Score，以及分類層級的 `Security Recommendation`。

> **設計原則：人工判定結果優先。AI 僅協助撰寫報告與判讀 CVSS Base Metrics，不自行改變人工風險等級，也不得編造未提供的測試事實。**

---

## 1. 專案結構

目前目錄結構如下：

```text
AutoReport_AI/
├─ templates/
│  ├─ check_iot_ai.xlsx
│  ├─ check_android_ai.xlsx
│  └─ check_ios_ai.xlsx
├─ .env
├─ .gitignore
├─ generate_iot_report_ai.py
├─ generate_mastg_report_ai.py
└─ README.md
```

各檔案用途：

| 檔案 / 目錄 | 用途 |
|---|---|
| `generate_iot_report_ai.py` | 處理 IoT 裝置檢測 Excel |
| `generate_mastg_report_ai.py` | 處理 OWASP MASTG Android / iOS Excel |
| `templates/check_iot_ai.xlsx` | IoT 檢測範本 |
| `templates/check_android_ai.xlsx` | Android MASTG 檢測範本 |
| `templates/check_ios_ai.xlsx` | iOS MASTG 檢測範本 |
| `.env` | AI Provider、模型與執行參數設定 |


> `.env` 可能包含 API Key 或內部環境資訊，**不要提交到 GitHub**。

---

## 2. 支援的 AI Provider

目前兩支程式可依 `.env` 切換 AI Provider：

- **Ollama**：地端模型，適合機敏檢測資料
- **OpenAI**
- **Google Gemini**

目前地端環境建議模型：

```text
gemma4:31b-it-q8_0
```

在目前 RTX 6000 Ada 48 GB 環境下，建議搭配：

```text
OLLAMA_NUM_CTX=8192
OLLAMA_THINK=false
```

以避免不必要的大型 Context 佔用 VRAM，並使模型盡量維持 100% GPU 推論。

---

## 3. 執行環境

建議使用 Python 3.10 以上版本。

### 共用套件

```powershell
pip install openpyxl python-dotenv requests
```

### 使用 OpenAI 時

```powershell
pip install openai
```

### 使用 Gemini 時

```powershell
pip install google-genai
```

如需一次安裝全部 Provider：

```powershell
pip install openpyxl python-dotenv requests openai google-genai
```

---

## 4. `.env` 設定

`.env` 請放在兩支 Python 程式相同目錄。

### 4.1 建議的 Ollama 設定

```env
AI_PROVIDER=ollama

# =========================
# Ollama
# =========================
OLLAMA_HOST=http://192.168.50.241:11434
OLLAMA_MODEL=gemma4:31b-it-q8_0
OLLAMA_TIMEOUT=300
OLLAMA_KEEP_ALIVE=10m
OLLAMA_NUM_CTX=8192
OLLAMA_THINK=false
OLLAMA_STRUCTURED_OUTPUT=true

# =========================
# Excel
# =========================
SHEET_NAME=IOT Device
REPORT_LANGUAGE=English
OVERWRITE_REPORT=true
```

### 4.2 IoT 與 MASTG 可共用同一份 `.env`

`SHEET_NAME=IOT Device` **可以保留，不需要在 IoT / MASTG 之間反覆修改**。

- `generate_iot_report_ai.py` 會使用 `SHEET_NAME=IOT Device`
- `generate_mastg_report_ai.py` 不使用此設定，會自動辨識：
  - `Android App`
  - `iOS App`

因此同一天要跑 IoT、Android、iOS 時，可以直接共用同一份 `.env`。

### 4.3 OpenAI 範例

```env
AI_PROVIDER=openai
OPENAI_API_KEY=你的_API_Key
OPENAI_MODEL=你的_OpenAI_模型名稱
OPENAI_REASONING=

SHEET_NAME=IOT Device
REPORT_LANGUAGE=English
OVERWRITE_REPORT=true
```

### 4.4 Gemini 範例

```env
AI_PROVIDER=gemini
GEMINI_API_KEY=你的_API_Key
GEMINI_MODEL=你的_Gemini_模型名稱

SHEET_NAME=IOT Device
REPORT_LANGUAGE=English
OVERWRITE_REPORT=true
```

> `REPORT_LANGUAGE` 為相容既有設定保留。目前程式固定同時產生英文 `Report` 與繁體中文 `Report_ch`。

---

## 5. 執行方式

建議從專案根目錄執行。

### 5.1 IoT

```powershell
python generate_iot_report_ai.py templates\check_iot_ai.xlsx
```

預設輸出：

```text
templates\check_iot_ai_report.xlsx
```

### 5.2 Android MASTG

```powershell
python generate_mastg_report_ai.py templates\check_android_ai.xlsx
```

預設輸出：

```text
templates\check_android_ai_report.xlsx
```

### 5.3 iOS MASTG

```powershell
python generate_mastg_report_ai.py templates\check_ios_ai.xlsx
```

預設輸出：

```text
templates\check_ios_ai_report.xlsx
```

### 5.4 自訂輸出檔名

```powershell
python generate_iot_report_ai.py templates\check_iot_ai.xlsx --output result_iot.xlsx
```

```powershell
python generate_mastg_report_ai.py templates\check_android_ai.xlsx --output result_android.xlsx
```

### 5.5 MASTG 指定平台

一般情況下程式會依工作表名稱自動判斷平台；如 Excel 同時含 Android 與 iOS 工作表，可手動指定：

```powershell
python generate_mastg_report_ai.py templates\check_android_ai.xlsx --platform android
```

或：

```powershell
python generate_mastg_report_ai.py templates\check_ios_ai.xlsx --platform ios
```

---

## 6. IoT 與 MASTG 的差異

| 項目 | IoT | MASTG |
|---|---|---|
| Python | `generate_iot_report_ai.py` | `generate_mastg_report_ai.py` |
| 範本 | `check_iot_ai.xlsx` | `check_android_ai.xlsx` / `check_ios_ai.xlsx` |
| 工作表 | `IOT Device` | `Android App` / `iOS App` |
| 平台判斷 | `.env` 的 `SHEET_NAME` | 自動辨識 Android / iOS |
| Android/iOS 測項過濾 | 不適用 | Android：S/D/F + MAxx；iOS：S/D/F + MIxx |
| MASTG Hyperlink 測項名稱解析 | 不適用 | 支援 |
| CVSS / Recommendation | 支援 | 支援 |

---

## 7. Excel 必要欄位

IoT 與 MASTG 主表至少需要下列欄位：

| 欄位 | 用途 |
|---|---|
| `編號` | 測項編號 |
| `測項` | 測試項目名稱 |
| `判定結果` | 檢測人員人工決定的風險結果 |
| `目前情況` | 測試方法、結果、攻擊條件、限制條件等事實 |
| `Report` | AI 產生的英文報告 |
| `Report_ch` | AI 產生的繁體中文報告 |
| `cvss` | CVSS v3.1 Base Vector |
| `score` | CVSS Base Score |

建議另外保留：

| 欄位 | 用途 |
|---|---|
| `分類` | 用於 `Security Recommendation` 分類彙整 |
| `修補建議` | 人工提供的改善方向，AI 可納入報告 |

---

## 8. 「目前情況」建議格式

為了讓模型更容易區分事實、攻擊條件與降低風險因素，可視需要使用：

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

此格式是 **可選**，不是每一段都必填。

填寫原則：

- 只寫實際取得的測試事實。
- 沒有資訊的段落可直接省略。
- 不需要為了格式完整填入「未提供」或「不適用」。
- 不確定的產品能力、攻擊條件或防護措施不要推測。
- Low 以上若希望 AI 產生較合理的 CVSS，建議盡量提供實際攻擊位置、權限需求、使用者互動與 C/I/A 影響。

---

## 9. 判定結果處理

程式可辨識：

```text
Critical
High
Medium
Low
None / No Risk
Not Applicable / N/A
TBD
Testing
```

處理原則：

- `TBD` / `Testing`：略過，不產生 Report。
- `Not Applicable`：產生 N/A 類型報告，不要求標準 CVSS Base Metrics。
- `None / No Risk`：產生無風險報告，Score 固定為 `0.0`。
- `Low / Medium / High / Critical`：產生 Report 並進行 CVSS v3.1 Base Metrics 判讀。

人工填寫的 `判定結果` 為最終檢測結論，AI 不會自行改變。

---

## 10. CVSS v3.1

程式僅處理 CVSS v3.1 Base Metrics：

```text
AV / AC / PR / UI / S / C / I / A
```

AI 負責判讀 Metrics，**Score 由 Python 依 CVSS v3.1 Base Score 公式計算**，不直接採用模型自行輸出的分數。

### Low 以上

第一輪若 Metrics 缺漏或格式錯誤，程式會自動再進行一次 CVSS 專項重試。

```text
[CVSS RETRY] Pxx: 第一輪 CVSS 不完整，重新判讀一次
```

若資訊仍不足以辨識具體漏洞或攻擊情境：

```text
cvss  = N/A
score = -
```

避免硬編不存在的 CVSS Metrics。

### None / No Risk

目前工具使用自訂的無風險表示方式：

```text
CVSS:3.1/AV:-/AC:-/PR:-/UI:-/S:-/C:N/I:N/A:N
score = 0.0
```

> `-` 不是 CVSS v3.1 官方 Base Vector 的標準 metric value，而是本工具為報告呈現所採用的自訂表示方式。

### Not Applicable

```text
cvss  = -
score = -
```

---

## 11. 人工風險與 CVSS Severity 不一致

人工風險判定與 CVSS Base Severity 是不同概念。

例如人工判定：

```text
Low
```

但 Python 依 AI 提供的 Base Metrics 計算為：

```text
4.3 Medium
```

程式會：

- 保留人工 `判定結果`
- 保留 CVSS 計算結果
- 將 `cvss` 與 `score` 以紅字顯示
- 在命令列輸出警告，提醒人工複核

這是預期行為，不代表程式錯誤。

---

## 12. Security Recommendation

程式會建立或重建：

```text
Security Recommendation
```

工作表主要欄位：

| 欄位 | 說明 |
|---|---|
| 分類 | IoT 分類或 MASTG / MASVS Domain |
| 最高風險 | 該分類目前最高人工風險 |
| 觸發測項 | Low 以上的測項編號 |
| Recommendation | 英文整體改善建議 |
| Recommendation_ch | 中文整體改善建議 |

只有 Low / Medium / High / Critical 項目會觸發 AI 產生分類建議。

建議內容以「分類層級」的整體改善方向為主，不逐筆重複 Report。

---

## 13. MASTG 特別處理

`generate_mastg_report_ai.py` 會自動依工作表名稱辨識平台：

```text
Android App
```

或：

```text
iOS App
```

並依平台過濾測項：

- Android：共通 `Sxx / Dxx / Fxx` + `MAxx`
- iOS：共通 `Sxx / Dxx / Fxx` + `MIxx`

因此 Android 檢測不會誤處理 `MIxx`，iOS 檢測也不會誤處理 `MAxx`。

若 MASTG 測項名稱儲存在 Excel `HYPERLINK()` 公式中，程式會取出實際顯示名稱後再送給 AI，而不直接將整段公式作為測項名稱。

---

## 14. Ollama 執行狀態確認

使用地端 Ollama 時，可在模型主機確認：

```bash
ollama ps
```

目前建議看到類似：

```text
PROCESSOR   100% GPU
CONTEXT     8192
```

搭配：

```bash
nvidia-smi
```

確認 GPU Utilization、VRAM 與功耗是否正常。

程式執行時會顯示目前 Provider / Model / Host，例如：

```text
[AI] Provider: ollama
[AI] Model: gemma4:31b-it-q8_0
[AI] Host: http://192.168.50.241:11434
[AI] Context: 8192
[AI] Thinking: false
```

每次 Ollama 呼叫也會輸出效能資訊，例如：

```text
[OLLAMA PERF] P06 | total: 8.4s | load: 0.3s | prompt: 1620 tokens | output: 180 tokens | 25.1 tok/s
```

---

## 15. `.gitignore` 建議

建議至少包含：

```gitignore
.env
venv/
__pycache__/
*.pyc
*_report.xlsx
backup/
```

`templates/` 內若放的是乾淨範本，可以提交 GitHub；若範本含客戶資料，請先去識別化或不要提交。

---

## 16. 資料安全提醒

若使用 OpenAI / Gemini 等雲端 Provider，送往模型的內容可能包含：

- 測項名稱
- 判定結果
- 目前情況
- 修補建議
- 被引用測項內容

因此不要直接放入：

- 密碼
- API Key
- Token
- 私鑰
- 未遮罩個資
- 未核准上傳雲端的機敏程式碼或產品資訊

若資料涉及客戶機密或受限制資訊，建議使用組織核准的地端 Ollama 環境，或依內部資料處理政策確認後再使用雲端模型。

---

## 17. 使用原則

本工具定位為：

> **資安檢測報告文字與 CVSS Base Metrics 的輔助產生工具，而不是取代檢測工程師判定的自動化決策工具。**

正式交付前仍應由檢測人員確認：

- 人工判定結果是否正確。
- Report 是否忠實反映實際測試。
- AI 是否誤解「目前情況」。
- CVSS Metrics 是否符合實際攻擊條件。
- CVSS 與人工風險不一致是否有合理依據。
- Recommendation 是否符合產品實際改善方向。

尤其是 High / Critical、醫療器材、重要系統或其他高風險產品，應進行完整人工複核。
