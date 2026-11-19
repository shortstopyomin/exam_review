/**
 * 預設規則檔（TOML）。網頁上的「下載規則檔」會下載這份內容，
 * 老師可以修改後再上傳。
 *
 * 用 String.raw 是為了讓正規表示式裡的反斜線原樣保留；
 * TOML 的單引號字串（'...'）同樣不處理跳脫字元，所以寫起來和 JavaScript 正規表示式一樣。
 */
export const DEFAULT_RULES_TOML = String.raw`# 試卷審題規則檔（TOML 格式）
# severity（嚴重程度）："error" = 不通過、"warning" = 警告、"info" = 提醒
# 不想用的規則可以加上 enabled = false
# 每一條規則以 [[rules]] 開頭；其他可用的規則類型請見 README

name = "預設試卷檢查規則"

# ── 1. 試卷標題 ────────────────────────────
# 每次出題前，請確認 examType（考試類型）是否正確

[[rules]]
id = "exam-title"
type = "exam-title"
within = 5                                   # 只看前 5 個段落（試卷標題）
schoolZh = "越華國際學校"
schoolEn = "Viet Hoa International School"
academicYear = "auto"     # "auto" 依今天日期推算（8 月起為新學年），或寫 "2026-2027"，"off" 不檢查
semester = "auto"         # "auto"：8 月～1 月為第 1 學期、2～7 月為第 2 學期；或寫 1、2；"off" 不檢查
examType = "final"        # midterm 期中 / final 期末 / quiz 小考 / mock 模擬考；刪掉這行則只檢查中英文是否一致
# grade = 10              # 年級；不寫則只檢查中英文年級是否一致
severity = "error"

# ── 2. 總分 ────────────────────────────────

[[rules]]
id = "total-score"
type = "score-sum"
expect = 100              # 全卷總分
severity = "error"
`;
