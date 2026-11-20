import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { parseDocx } from "../docx/parse";
import { formatNumber } from "../docx/numbering";
import { lint, parseRuleSet, RuleSetError } from "../rules/engine";
import { DEFAULT_RULES_TOML } from "../rules/default-rules";
import { extractOptions } from "../rules/builtin";
import type { Paragraph, ParsedDocument, RuleSet } from "../types";

// 固定「今天」，讓學年度／學期的自動推算不會隨測試執行日期改變
function fixToday(set: RuleSet): RuleSet {
  for (const r of set.rules) if (r.type === "exam-title") r.today = "2026-11-19";
  return set;
}

/** 網頁上的預設規則（目前只有標題與總分兩條） */
const defaults = fixToday(parseRuleSet(DEFAULT_RULES_TOML));
/** 完整規則（所有內建規則類型），用來測試各規則本身 */
const allRules = fixToday(
  parseRuleSet(readFileSync(new URL("./fixtures/all-rules.toml", import.meta.url), "utf-8")),
);
/** 不含學校標題檢查的規則，用於與標題無關的測試 */
const generic = { ...allRules, rules: allRules.rules.filter((r) => r.type !== "exam-title") };

function docFrom(lines: string[]): ParsedDocument {
  const paragraphs: Paragraph[] = lines.map((text, index) => ({
    index, text, inTable: false, hasMath: false, imageCount: 0,
  }));
  return { fileName: "t.docx", paragraphs };
}

const ids = (doc: ParsedDocument) => lint(doc, allRules).issues.map((i) => i.ruleId);

describe("範例試卷（端對端）", async () => {
  const buf = readFileSync("public/samples/sample-exam.docx");
  const doc = await parseDocx(buf);
  const result = lint(doc, allRules);
  const msgs = result.issues.map((i) => `${i.ruleId}: ${i.message}`);

  it("還原 Word 自動編號", () => {
    const tf = result.structure.sections[1];
    expect(tf.kind).toBe("truefalse");
    expect(tf.questions.map((q) => q.number)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
  });

  it("找出埋入的問題", () => {
    expect(msgs).toEqual(
      expect.arrayContaining([
        expect.stringMatching(/^question-numbering: 題號由 6 跳到 8/),
        expect.stringMatching(/^choice-options: 第 3 題有 3 個選項/),
        expect.stringMatching(/^choice-options: 第 6 題選項 \(A\) 與 \(C\) 內容相同/),
        expect.stringMatching(/^answer-leftover: 第 5 題/),
        expect.stringMatching(/^halfwidth-punctuation:/),
        expect.stringMatching(/^repeated-punctuation:/),
        expect.stringMatching(/^bracket-balance:/),
        expect.stringMatching(/^answer-in-paper:/),
        expect.stringMatching(/^subject-name: 科目名稱「Mathmatics」疑似拼錯，應為「Mathematics」/),
        expect.stringMatching(/^exam-title: 英文學校名稱「Viet Hoa Internatinal School」疑似寫錯/),
      ]),
    );
    expect(result.passed).toBe(false);
  });

  it("計算題標示共 16 分、各題 8 + 6 = 14 分：不分題檢查，全卷 100 分通過", () => {
    expect(msgs.some((m) => m.startsWith("total-score"))).toBe(false);
  });

  it("不誤報基本資料欄位", () => {
    expect(msgs.some((m) => m.startsWith("header-fields"))).toBe(false);
  });
});

describe("規則", () => {
  it("乾淨的試卷通過", () => {
    const doc = docFrom([
      "七年級數學科 小考", "班級： 座號： 姓名：",
      "一、選擇題（每題 50 分，共 100 分）",
      "( )1. 一加一？", "(A) 1 (B) 2 (C) 3 (D) 4",
      "( )2. 二加二？", "(A) 1 (B) 2 (C) 3 (D) 4",
    ]);
    expect(lint(doc, generic).issues).toEqual([]);
  });

  it("總分不等於 100", () => {
    const doc = docFrom(["班級座號姓名", "一、是非題（每題 10 分）", "( )1. 甲", "( )2. 乙"]);
    expect(lint(doc, allRules).issues.map((i) => i.message)).toContain("全卷配分加總為 20 分，應為 100 分");
  });

  it("小數不會被當成題號", () => {
    const doc = docFrom(["班級座號姓名", "一、計算題（共 100 分）", "1. 計算", "3.5 + 2 = ?"]);
    expect(ids(doc)).not.toContain("question-numbering");
  });

  it("選項可寫在同一行或分行", () => {
    expect(extractOptions("題目\n(A) 甲 (B) 乙\n(C) 丙 (D) 丁").map((o) => o.letter).join("")).toBe("ABCD");
    expect(extractOptions("題目\nA. 甲\nB. 乙\nC. 丙").map((o) => o.text)).toEqual(["甲", "乙", "丙"]);
    expect(extractOptions("（Ａ）甲（Ｂ）乙").map((o) => o.letter)).toEqual(["A", "B"]);
  });
});

describe("規則檔驗證", () => {
  it("未知的 type", () => {
    expect(() => parseRuleSet('[[rules]]\nid = "a"\ntype = "nope"')).toThrow(RuleSetError);
  });
  it("錯誤的正規表示式", () => {
    expect(() => parseRuleSet("[[rules]]\nid = \"a\"\ntype = \"regex\"\npattern = '('")).toThrow(/正規表示式/);
  });
  it("id 重複", () => {
    expect(() => parseRuleSet('[[rules]]\nid = "a"\ntype = "numbering"\n[[rules]]\nid = "a"\ntype = "numbering"')).toThrow(/重複/);
  });
  it("TOML 語法錯誤", () => {
    expect(() => parseRuleSet("[[rules]\nid = ")).toThrow(/TOML 格式錯誤/);
  });
  it("沒有規則", () => {
    expect(() => parseRuleSet('name = "x"')).toThrow(/\[\[rules\]\]/);
  });
});

describe("編號格式", () => {
  it.each([
    [3, "decimal", "3"], [3, "upperLetter", "C"], [4, "lowerRoman", "iv"],
    [3, "taiwaneseCountingThousand", "三"], [12, "taiwaneseCountingThousand", "十二"],
    [25, "chineseCounting", "二十五"], [2, "ideographTraditional", "乙"], [3, "decimalEnclosedCircle", "③"],
  ])("%i %s → %s", (n, fmt, out) => expect(formatNumber(n, fmt)).toBe(out));
});

describe("科目名稱", () => {
  const subjectIssues = (...header: string[]) =>
    lint(docFrom([...header, "一、選擇題（共 100 分）"]), allRules)
      .issues.filter((i) => i.ruleId === "subject-name")
      .map((i) => i.message);

  it.each([
    ["七年級數子科 第一次段考", "「數子」疑似寫錯，應為「數學」"],
    ["高二物裡科期末考", "「物裡」疑似寫錯，應為「物理」"],
    ["科目：地裡", "「地裡」疑似寫錯，應為「地理」"],
    ["五年級 自然科厂 期中評量", "應為「自然科學」"],
    ["Grade 8 Biolgy Final Exam", "「Biolgy」疑似拼錯，應為「Biology」"],
    ["Grade 9 Histroy Quiz", "「Histroy」疑似拼錯，應為「History」"],
    ["Grade 10 Socail Studies Midterm", "「Socail Studies」疑似拼錯，應為「Social Studies」"],
    ["Grade 11 Earht Science Test", "「Earht Science」疑似拼錯，應為「Earth Science」"],
    ["G12 Chemestry Exam", "「Chemestry」疑似拼錯，應為「Chemistry」"],
  ])("%s", (header, expected) => {
    const msgs = subjectIssues(header);
    expect(msgs).toHaveLength(1);
    expect(msgs[0]).toContain(expected);
  });

  it.each([
    "臺北市立和平國民中學 114 學年度第一學期 七年級數學科 第一次段考",
    "新北市立○○高級中學 113學年度第二學期 高二 公民與社會 期末考",
    "三年級 國語 第二次定期評量",
    "Grade 7 Mathematics Midterm Exam",
    "Grade 10 Earth Sciences Final",
    "G5 English Language Arts Quiz",
    "七年級 英聽 測驗",
    "高三 數甲 複習卷 Grade 12 Physics Review",
  ])("正確不誤報：%s", (header) => {
    expect(subjectIssues(header)).toEqual([]);
  });

  it("找不到科目時提醒", () => {
    expect(subjectIssues("期中考")).toEqual([expect.stringContaining("找不到科目名稱")]);
  });

  it("extraSubjects 與 ignore", () => {
    const set = parseRuleSet(
      '[[rules]]\nid = "s"\ntype = "subject-name"\nextraSubjects = ["數學丙"]\nignore = ["Chemestry"]',
    );
    const run = (h: string) => lint(docFrom([h]), set).issues.map((i) => i.message);
    expect(run("數學丙 段考")).toEqual([]);
    expect(run("Chemestry Exam")).toEqual([expect.stringContaining("找不到科目名稱")]);
  });
});

describe("試卷標題", () => {
  const USER_TITLE = [
    "越華國際學校 2026-2027 學年度 第 1學期 10年級 中文數學 期中評量試卷",
    "Viet Hoa International School Grade 10 Math C Midterm Exam of Semester 1 Year 2026-2027",
    "班級 Class ：                     座號 No.：                    姓名 Name:                        總分 Score:",
  ];
  const FIXED_TITLE = [
    "越華國際學校 2026-2027 學年度 第 1 學期 10年級 中文數學 期末評量試卷",
    "Viet Hoa International School Grade 10 Math C Final Exam of Semester 1 Year 2026-2027",
    USER_TITLE[2],
  ];
  const titleIssues = (lines: string[], extra = "") => {
    const set = parseRuleSet(
      `[[rules]]
id = "t"
type = "exam-title"
schoolZh = "越華國際學校"
schoolEn = "Viet Hoa International School"
today = "2026-11-19"
${extra}`,
    );
    return lint(docFrom(lines), set).issues.map((i) => i.message);
  };

  it("使用者提供的標題：期中應為期末", () => {
    expect(titleIssues(USER_TITLE, 'examType = "final"')).toEqual([
      "考試類型「期中」應為「期末」",
      "考試類型「Midterm」應為「Final」",
    ]);
  });

  it("修正後的標題通過", () => {
    expect(titleIssues(FIXED_TITLE, 'examType = "final"\ngrade = 10')).toEqual([]);
  });

  it("沒設定 examType 時檢查中英文是否一致", () => {
    const lines = [FIXED_TITLE[0], USER_TITLE[1]]; // 中文期末、英文 Midterm
    expect(titleIssues(lines)).toEqual([expect.stringContaining("中英文標題的考試類型不一致")]);
  });

  it("學校名稱寫錯", () => {
    const lines = [
      "越華國際學 2026-2027 學年度 第 1 學期 10年級 數學 期末評量試卷",
      "Viet Hoa Internatonal Schol Grade 10 Math Final Exam of Semester 1 Year 2026-2027",
    ];
    expect(titleIssues(lines)).toEqual([
      expect.stringContaining("中文學校名稱「越華國際學」疑似寫錯"),
      expect.stringContaining("英文學校名稱「Viet Hoa Internatonal Schol」疑似寫錯"),
    ]);
  });

  it("英文學校名稱大小寫", () => {
    const lines = [FIXED_TITLE[0], FIXED_TITLE[1].replace("International School", "international school")];
    expect(titleIssues(lines)).toEqual([expect.stringContaining("大小寫不正確")]);
  });

  it("學年度錯誤與不連續", () => {
    const lines = [
      FIXED_TITLE[0].replace("2026-2027", "2025-2026"),
      FIXED_TITLE[1].replace("2026-2027", "2026-2028"),
    ];
    expect(titleIssues(lines)).toEqual([
      "學年度「2025-2026」應為「2026-2027」",
      "學年度「2026-2028」前後年份不連續",
    ]);
  });

  it("學期錯誤", () => {
    const lines = [FIXED_TITLE[0].replace("第 1 學期", "第 2 學期"), FIXED_TITLE[1]];
    expect(titleIssues(lines)).toEqual(["「第 2 學期」應為第 1 學期"]);
  });

  it("中英文年級不一致", () => {
    const lines = [FIXED_TITLE[0], FIXED_TITLE[1].replace("Grade 10", "Grade 11")];
    expect(titleIssues(lines)).toEqual([expect.stringContaining("中英文標題的年級不一致")]);
  });

  it("2 月後自動推算為第 2 學期", () => {
    const set = parseRuleSet('[[rules]]\nid = "t"\ntype = "exam-title"\ntoday = "2027-03-01"');
    const issues = lint(docFrom(FIXED_TITLE), set).issues.map((i) => i.message);
    expect(issues).toContain("「第 1 學期」應為第 2 學期");
    expect(issues).toContain("「Semester 1」應為第 2 學期");
  });

  it("民國學年度", () => {
    const set = parseRuleSet('[[rules]]\nid = "t"\ntype = "exam-title"\ntoday = "2026-11-19"\nsemester = "off"');
    const run = (h: string) => lint(docFrom([h]), set).issues.map((i) => i.message);
    expect(run("115 學年度 數學 期末考")).toEqual([]);
    expect(run("114 學年度 數學 期末考")).toEqual(["「114 學年度」應為「115 學年度」"]);
  });
});

describe("預設規則（標題＋總分）", () => {
  it("只有兩條規則", () => {
    expect(defaults.rules.map((r) => r.type)).toEqual(["exam-title", "score-sum"]);
  });

  const header = [
    "越華國際學校 2026-2027 學年度 第 1 學期 10年級 中文數學 期末評量試卷",
    "Viet Hoa International School Grade 10 Math C Final Exam of Semester 1 Year 2026-2027",
    "班級 Class ：  座號 No.：  姓名 Name:  總分 Score:",
  ];

  it("標題正確、總分 100 → 通過", () => {
    const doc = docFrom([...header, "一、選擇題（每題 50 分，共 100 分）", "( )1. 甲", "( )2. 乙"]);
    const r = lint(doc, defaults);
    expect(r.issues).toEqual([]);
    expect(r.passed).toBe(true);
  });

  it("標題寫期中、總分 90 → 不通過", () => {
    const doc = docFrom([
      header[0].replace("期末", "期中"),
      header[1],
      "一、選擇題（每題 30 分）", "( )1. 甲", "( )2. 乙", "( )3. 丙",
    ]);
    const r = lint(doc, defaults);
    expect(r.issues.map((i) => i.message).sort()).toEqual(
      ["考試類型「期中」應為「期末」", "全卷配分加總為 90 分，應為 100 分"].sort(),
    );
    expect(r.passed).toBe(false);
  });
});
