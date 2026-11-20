# 試卷審題小幫手

讓老師上傳 Word 試卷（.docx），依照 TOML 規則檔自動檢查題號、配分、選項與文字。

- **純前端**：試卷只在瀏覽器中解析，不會上傳到任何伺服器
- **免費部署**：Next.js 靜態輸出，部署在 GitHub Pages

## 本機開發

```bash
npm install
npm run dev        # http://localhost:3000
npm test           # 執行測試
npm run sample     # 重新產生範例試卷 public/samples/sample-exam.docx
```

## 部署到 GitHub Pages

1. 把專案 push 到 GitHub 的 `main` 分支
2. 到 repo 的 **Settings → Pages → Build and deployment → Source** 選 **GitHub Actions**
3. 之後每次 push 到 `main`，`.github/workflows/deploy.yml` 會自動測試、建置並部署

網址會是 `https://<帳號>.github.io/<repo 名稱>/`。`basePath` 由 workflow 自動帶入，不需要手動設定。

## 試卷格式

解析器會辨識以下格式（Word 自動編號也可以）：

| 項目 | 範例 |
|---|---|
| 大題標題 | `一、選擇題（每題 4 分，共 40 分）`、`貳、非選擇題`、`第二部分 計算題`、`Part 2計算題和應用題 / Calculation and word problems (每題5分，共60分)`、`Part II: Word Problems`、`Section B - Multiple Choice` |
| 大題配分（英文） | `5 points each`、`60 points total`、`Total: 60`、`(60 marks)` |
| 每格配分 | `每格 2 分、共 30 分｜2 points each, 30 points total`、`每空 2 分`、`2 points per blank`；**不數格子**，大題分數採用「共 Y 分」（所以寫每格時一定要寫共幾分）；全卷總分仍會檢查 |
| 配分寫在大題標題下一行 | 標題下、第一題之前的配分說明（例如上面那行）也讀得到 |
| 題目 | `1.`、`1、`、`( )1.`、`（　）12．` |
| 選擇題選項 | `(A) … (B) …`（可同一行或分行）、`A. …`（需在行首） |
| 單題配分 | 題目中的 `（8 分）`、`(5 points)`、`[5 marks]` |

大題配分的判斷順序：「共 Y 分」→「每題 X 分 × 題數」→ 各題「（X 分）」加總（「每格」不換算）。只檢查各大題分數加起來是否等於全卷總分，**不檢查**大題內的題數 × 每題配分是否等於「共 Y 分」。英文的「points each」沒說是每題還是每格，會跟著中文描述的單位。中英文配分的分數、總分或單位互相矛盾時會回報。

## 規則檔

規則檔是 [TOML](https://toml.io/cn/) 格式，預設規則在 `src/lib/rules/default-rules.ts`，目前只啟用兩條：**試卷標題**（`exam-title`）與**總分是否為 100 分**（`total-score`）。下表其他規則類型都還能用，把它們加進規則檔即可；完整範例見 `src/lib/__tests__/fixtures/all-rules.toml`。網頁上可以直接編輯、上傳、下載，修改會存在瀏覽器（localStorage）。

```toml
name = "我的規則"

[[rules]]                  # 每一條規則以 [[rules]] 開頭
id = "total-score"         # 自訂名稱，不可重複
type = "score-sum"         # 規則類型（見下表）
expect = 100
severity = "error"         # "error" 不通過 / "warning" 警告 / "info" 提醒
message = "自訂訊息（選填）"
enabled = true             # false 可暫時停用

[[rules]]
id = "no-answer"
type = "regex"
pattern = '答案\s*[:：]'   # 正規表示式請用單引號，反斜線不用寫兩次
```

| type | 用途 | 參數 |
|---|---|---|
| `score-sum` | 各大題配分加總是否等於全卷總分 | `expect`（預設 100） |
| `numbering` | 題號連續、不重複 | `restart`: `auto` / `always` / `never` |
| `choice-options` | 選擇題選項數量、順序、重複 | `count`: `4` 或 `[4, 5]` |
| `answer-leftover` | 題號前的作答括號不可有答案 | — |
| `bracket-balance` | 括號、引號成對 | — |
| `exam-title` | 標題的學校名稱、學年度、學期、年級、考試類型，以及中英文標題是否一致 | `schoolZh`、`schoolEn`、`academicYear`（`"auto"`／`"2026-2027"`／`"off"`）、`semester`（`"auto"`／`1`／`2`／`"off"`）、`examType`（`midterm`／`final`／`quiz`／`mock`）、`grade`、`within` |
| `subject-name` | 試卷開頭的中英文科目名稱是否寫錯、拼錯（G1–G12 常見科目，清單在 `src/lib/rules/subjects.ts`） | `within`（預設 5 段）、`extraSubjects`、`ignore`、`requireSubject` |
| `regex` | 自訂正規表示式 | `pattern` 或 `patterns`、`mode`: `forbid` / `require`、`within`、`flags`、`maxReports` |

## 專案結構

```
src/lib/docx/       解析 .docx（JSZip + XML，含 Word 自動編號還原）
src/lib/structure.ts 辨識大題、題目、配分
src/lib/rules/      規則引擎、內建規則、預設規則檔
src/components/     網頁介面
scripts/            產生範例試卷
```

## 已知限制

- 不支援舊版 `.doc`，請先另存為 `.docx`
- 圖片中的文字、文字方塊的版面位置不會檢查
- 方程式只擷取文字內容
