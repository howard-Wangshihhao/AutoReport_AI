# AutoReport AI

AutoReport AI 是一套以 Excel 為輸入的資安檢測報告輔助產生工具，目前分為 **IoT 裝置檢測** 與 **OWASP MASTG 行動 App 檢測（Android / iOS）** 兩條流程。

工具以檢測人員人工填寫的 **「判定結果」**、**「目前情況」** 與 **「修補建議」** 為主要依據，呼叫 AI 產生英文 `Report`、繁體中文 `Report_ch`、CVSS v3.1 Base Vector / Score，以及分類層級的 `Security Recommendation`。

> **設計原則：人工判定結果優先。AI 僅協助撰寫報告與判讀 CVSS Base Metrics，不自行改變人工風險等級，也不得編造未提供的測試事實。**

---

## 1. 專案結構

目前目錄結構建議如下：

```text
AutoReport_AI/
├─ templates/
│  ├─ check_iot_ai.xlsx
│  ├─ check_android_ai.xlsx
│  └─ check_ios_ai.xlsx
├─ venv/
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
| `venv/` | Python 虛擬環境 |
| `.env` | AI Provider、模型與執行參數設定 |

> `.env` 可能包含 API Key 或內部環境資訊，**不要提交到 GitHub**。

---

## 2. 支援的 AI Provider

目前兩支程式皆可依 `.env` 切換 AI Provider：

- **Ollama**：地端模型，適合機敏檢測資料
- **OpenAI**
- **Google Gemini**

目前實際環境可穩定使用：

```text
qwen2.5:14b
```

若希望提高報告措辭、CVSS 判讀與複雜情境的穩定度，也可使用：

```text
gemma4:31b-it-q8_0
```

但目前環境曾觀察到 Ollama Server 載入 `gemma4:31b-it-q8_0` 時自動使用 `262144` context，造成 31B Q8 模型出現 CPU/GPU 混合 offload，推論速度大幅下降甚至 timeout。因此在 Server context 尚未能固定為較小值前，建議日常批次作業先使用 `qwen2.5:14b`。

程式端仍建議設定：

```text
OLLAMA_NUM_CTX=8192
OLLAMA_THINK=false
```

> `OLLAMA_NUM_CTX` 會隨 `/api/chat` request 傳送，但實際 runner context 仍應以 `ollama ps` 顯示結果為準。

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

`.env` 請放在兩支 Python 程式相同目錄。程式會使用該檔案，且以 `.env` 內容覆蓋同名的既有環境變數。

### 4.1 目前建議的 Ollama 設定

```env
AI_PROVIDER=ollama

# =========================
# Ollama
# =========================
OLLAMA_HOST=http://192.168.50.241:11434
OLLAMA_MODEL=qwen2.5:14b
# OLLAMA_MODEL=gemma4:31b-it-q8_0
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

程式不允許輸出路徑與原始 Excel 相同，以避免覆蓋原始檢測資料。

### 5.5 MASTG 指定平台

一般情況下程式會依工作表名稱自動判斷平台；如 Excel 同時含 Android 與 iOS 工作表，可手動指定：

```powershell
python generate_mastg_report_ai.py templates\check_android_ai.xlsx --platform android
```

或：

```powershell
python generate_mastg_report_ai.py templates\check_ios_ai.xlsx --platform ios
```

### 5.6 只標準化 Excel，不呼叫 AI

兩支程式皆支援：

```powershell
--normalize-only
```

例如：

```powershell
python generate_mastg_report_ai.py templates\check_android_ai.xlsx --normalize-only
```

此模式只將來源 Excel 中的 x14 清單型 Data Validation 轉為較適合 openpyxl 維護的標準 Data Validation，不產生 Report，也不呼叫 AI。

預設輸出會改為：

```text
<input>_template.xlsx
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
| MASTG HYPERLINK 測項名稱解析 | 不適用 | 支援 |
| MASVS Domain | 不適用 | 支援分類彙整 |
| CVSS / Recommendation | 支援 | 支援 |

---

## 7. Excel 必要欄位

IoT 與 MASTG 主表至少需要下列欄位：

| 欄位 | 用途 |
|---|---|
| `編號` | 測項編號 |
| `測項` | 測試項目名稱 |
| `判定結果` | 檢測人員人工決定的風險結果 |
| `目前情況` | 測試方式、結果、攻擊條件、既有防護／限制條件等事實 |
| `Report` | AI 產生的英文報告 |
| `Report_ch` | AI 產生的繁體中文報告 |
| `cvss` | CVSS v3.1 Base Vector |
| `score` | CVSS Base Score |

建議另外保留：

| 欄位 | 用途 |
|---|---|
| `分類` | 用於 `Security Recommendation` 分類彙整 |
| `修補建議` | 人工提供的改善方向，AI 可納入報告 |

若 `判定結果` 或 `目前情況` 為空，該列不會進行 AI Report 產生。

---

## 8. 「目前情況」建議格式

目前建議統一使用以下四段：

```text
測試方式：
- ...

測試結果：
- ...

攻擊條件：
- ...

既有防護／限制條件：
- ...
```

填寫原則：

- **測試方式**：實際如何檢查、觀察、掃描、操作或驗證；只有真的使用的工具才寫工具名稱。
- **測試結果**：只寫實際觀察到的結果、弱點、符合情形或不適用原因。
- **攻擊條件**：Low 以上建議盡量補上攻擊者需要的存取位置、權限、使用者互動、實體接觸或其他必要條件。
- **既有防護／限制條件**：只寫產品或場域已確認存在的控制、限制、使用情境或降低風險因素。
- 沒有資料的段落可以保留標題後空白，不要為了格式完整而自行推測。
- 不確定的產品能力、攻擊方式或防護措施不要補寫。
- Low 以上若希望 AI 產生較合理的 CVSS，建議把實際攻擊條件與 C/I/A 影響寫清楚。

---

## 9. Report 產生規則

目前 IoT 與 MASTG 兩支程式已統一 Report 寫作方式。

### 9.1 第一個句子先交代測項目的／範圍

英文 Report 第一個句子會優先使用類似：

```text
This test item evaluated ... for security vulnerabilities / security controls.
```

繁體中文 Report 第一個句子會優先使用類似：

```text
本測試項目針對……進行安全弱點評估／安全性評估。
```

第一句只用來交代「這個測項在評估什麼」，**不代表已執行某個特定工具、payload、掃描或攻擊**。

第一句之後，才依 `目前情況` 描述：

- 實際測試方式
- 確認方式
- 觀察結果
- 弱點或風險
- 攻擊條件
- 既有防護／限制條件
- 已提供的修補方向

### 9.2 不得自行增加測試事實

MASTG 測項名稱只代表「要檢查什麼」，不是已觀察到的弱點證據。IoT 亦同樣遵守此原則。

AI 不得因測項名稱自行假設：

- 使用了某套工具
- 執行了某個 payload
- 存在某種加密演算法或通訊協定
- 已有某項安全控制
- 成功完成某種攻擊

### 9.3 固定結論

Report 最後會保留與人工 `判定結果` 對應的結論；AI 不得自行升級或降低人工風險等級。

---

## 10. 判定結果處理

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

## 11. CVSS v3.1

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

## 12. 人工風險與 CVSS Severity 不一致

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

## 13. Security Recommendation

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

## 14. MASTG 特別處理

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

對 MASVS 測項，程式可依工作表中的 MASVS Domain 標題進行分類彙整，並將平台與 Domain 資訊一起提供給 AI，以避免 Android / iOS 專有機制混用。

若 `目前情況` 中明確寫有例如：

```text
參考 MI1
參考 MA5
Refer to S01
```

程式可把被引用測項的測項名稱、判定結果與目前情況一併提供給 AI 作為補充上下文，但不會用參考測項改變本列的人工判定結果。

---

## 15. Excel 相容性與輸出格式

兩支程式皆會盡量保留來源 Excel 的既有格式，並處理以下項目：

- x14 清單型 Data Validation 轉為標準 Data Validation
- `Report` / `Report_ch` / `cvss` / `score` 文字自動換行
- Report 欄位沿用中文 Report 欄位的基礎格式
- Report 與 CVSS 輸出採 12pt
- 依長文字內容自動調整列高
- 人工風險與 CVSS Severity 不一致時，`cvss` / `score` 以紅字提醒

---

## 16. Ollama 執行狀態確認與問題排查

### 16.1 確認 API 可連線

Windows PowerShell：

```powershell
curl.exe http://192.168.50.241:11434/api/tags
```

若可正常回傳模型清單，代表 Windows 到 Ollama API 的基本網路連線正常。

### 16.2 確認模型實際載入狀態

Ollama 主機：

```bash
ollama ps
```

搭配：

```bash
nvidia-smi
```

檢查：

- `PROCESSOR` 是否大量 offload 到 CPU
- `CONTEXT` 是否符合預期
- VRAM 是否接近上限
- GPU Utilization 是否正常

### 16.3 目前環境的已知現象

目前環境曾觀察：

```text
gemma4:31b-it-q8_0
PROCESSOR  21%/79% CPU/GPU
CONTEXT    262144
```

此時 31B Q8 推論速度會明顯下降，甚至造成 Python API timeout。

同一套程式與 Ollama Server 改用：

```text
qwen2.5:14b
```

可正常且快速執行，因此目前日常批次作業建議先使用 `qwen2.5:14b`。

若未來 Ollama 管理端能將 Server context 固定為較小值，再切回 `gemma4:31b-it-q8_0` 做品質優先的產出。

### 16.4 程式啟動資訊

程式執行時會顯示目前 Provider / Model / Host，例如：

```text
[CONFIG] .env: D:\AutoReport_AI\.env
[AI] Provider: ollama
[AI] Model: qwen2.5:14b
[AI] Host: http://192.168.50.241:11434
[AI] Context: 8192
[AI] Thinking: false
```

每次 Ollama 呼叫也會輸出效能資訊，例如：

```text
[OLLAMA PERF] S01 | total: 6.5s | load: 0.3s | prompt: 1500 tokens | output: 160 tokens | 24.0 tok/s
```

> 啟動畫面中的 `[AI] Context: 8192` 代表程式準備送出的 request 設定；Ollama runner 實際採用的 context 請以 `ollama ps` 為準。

---

## 17. `.gitignore` 建議

建議至少包含：

```gitignore
.env
venv/
__pycache__/
*.pyc
*_report.xlsx
*_template.xlsx
backup/
```

`templates/` 內若放的是乾淨範本，可以提交 GitHub；若範本含客戶資料，請先去識別化或不要提交。

---

## 18. 資料安全提醒

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

## 19. 使用原則

本工具定位為：

> **資安檢測報告文字與 CVSS Base Metrics 的輔助產生工具，而不是取代檢測工程師判定的自動化決策工具。**

正式交付前仍應由檢測人員確認：

- 人工判定結果是否正確。
- Report 第一個句子是否正確描述該測項的評估目的／範圍。
- Report 是否忠實反映 `目前情況` 中的實際測試方式與結果。
- AI 是否誤解攻擊條件或既有防護／限制條件。
- CVSS Metrics 是否符合實際攻擊條件。
- CVSS 與人工風險不一致是否有合理依據。
- Recommendation 是否符合產品實際改善方向。

尤其是 High / Critical、醫療器材、重要系統或其他高風險產品，應進行完整人工複核。
