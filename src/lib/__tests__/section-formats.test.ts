/**
 * 大題標題格式與配分：中英雙語試卷（Part 1 / Part 2）。
 */
import { describe, expect, it } from "vitest";
import { analyzeStructure, sectionScore } from "../structure";
import { lint, parseRuleSet } from "../rules/engine";
import type { ParsedDocument } from "../types";

const PART2 = "Part 2計算題和應用題 / Calculation and word problems (每題5分，共60分)";

function doc(lines: string[]): ParsedDocument {
  return {
    fileName: "t.docx",
    paragraphs: lines.map((text, index) => ({ index, text, inTable: false, hasMath: false, imageCount: 0 })),
  };
}

const range = (from: number, n: number) => Array.from({ length: n }, (_, i) => from + i);

/** 雙語數學試卷：Part 1 選擇題 20 題 × 2 分 + Part 2 計算與應用題 n 題 × 5 分 */
function mathExam(part2Count: number, part2Header = PART2) {
  return [
    "Part 1選擇題 / Multiple choice (每題2分，共40分)",
    ...range(1, 20).flatMap((n) => [`( )${n}. 題目 ${n}`, "(A) 1 (B) 2 (C) 3 (D) 4"]),
    part2Header,
    ...range(1, part2Count).flatMap((n) => [
      `${n}. 小明有 ${n * 3} 顆蘋果，分給 3 個人，每人幾顆？`,
      "(1) 列出算式",
      "(2) 求答案",
      "答：______",
    ]),
  ];
}

const scoreRule = parseRuleSet('[[rules]]\nid = "total-score"\ntype = "score-sum"\nexpect = 100');
const scoreIssues = (lines: string[]) => lint(doc(lines), scoreRule).issues.map((i) => i.message);

describe("使用者提供的 Part 2 標題", () => {
  it("辨識為大題，讀出每題 5 分、共 60 分", () => {
    const s = analyzeStructure(doc([PART2, "1. 計算 3 + 4", "2. 應用題"]).paragraphs).sections;
    expect(s).toHaveLength(1);
    expect(s[0]).toMatchObject({ title: PART2, kind: "short", perQuestionScore: 5, totalScore: 60 });
    expect(s[0].questions.map((q) => q.number)).toEqual([1, 2]);
  });

  it("小題「(1)(2)」和「答：」不會被當成新題目", () => {
    const s = analyzeStructure(doc(mathExam(12)).paragraphs).sections;
    expect(s.map((x) => x.questions.length)).toEqual([20, 12]);
    expect(s.map((x) => sectionScore(x).total)).toEqual([40, 60]);
  });

  it("12 題 × 5 分 = 60 分，加上 Part 1 的 40 分 → 總分 100，通過", () => {
    expect(scoreIssues(mathExam(12))).toEqual([]);
  });

  it.each([11, 13])("題數不符（%i 題 × 5 分）不再檢查，只看大題標示的「共 60 分」", (n) => {
    expect(scoreIssues(mathExam(n))).toEqual([]);
  });

  it("標示的分數本身加起來不是 100", () => {
    const header = PART2.replace("共60分", "共50分").replace("每題5分", "每題5分");
    const msgs = scoreIssues(mathExam(10, header));
    expect(msgs).toEqual(["全卷配分加總為 90 分，應為 100 分"]);
  });
});

describe("各種大題標題格式", () => {
  it.each([
    ["一、選擇題（每題2分，共40分）", "choice", 2, 40],
    ["貳、非選擇題（共60分）", "short", undefined, 60],
    ["第二部分 計算題（每題5分，共60分）", "short", 5, 60],
    ["第2大題：填充題（共20分）", "fill", undefined, 20],
    ["Part 2計算題和應用題 / Calculation and word problems (每題5分，共60分)", "short", 5, 60],
    ["Part II: Word Problems (5 points each, 60 points total)", "short", 5, 60],
    ["Part 1 - Multiple Choice (2 points each, total 40 points)", "choice", 2, 40],
    ["Section B: True or False (each question is worth 2 points)", "truefalse", 2, undefined],
    ["Section A Fill in the blanks (20 marks)", "fill", undefined, 20],
    ["PART 3 Short Answer (Total: 30)", "short", undefined, 30],
  ])("%s", (title, kind, per, total) => {
    const s = analyzeStructure(doc([title, "1. 題目"]).paragraphs).sections;
    expect(s).toHaveLength(1);
    expect(s[0]).toMatchObject({ kind, perQuestionScore: per, totalScore: total });
  });

  it.each([
    "Part of the answer is shown below.",
    "Section of a circle",
    "Particle physics",
  ])("一般句子不會被當成大題：%s", (text) => {
    expect(analyzeStructure(doc([text]).paragraphs).sections).toEqual([]);
  });

  it("英文單題配分 (5 points)", () => {
    const s = analyzeStructure(doc(["Part 2 Word problems", "1. Solve x. (5 points)", "2. Find y. [5 marks]"]).paragraphs);
    expect(s.sections[0].questions.map((q) => q.score)).toEqual([5, 5]);
    expect(sectionScore(s.sections[0]).total).toBe(10);
  });
});

describe("作答說明不算題目、不算大題", () => {
  const TITLE = [
    "越華國際學校 2026-2027 學年度 第 1 學期 10年級 中文數學 期末評量試卷",
    "Viet Hoa International School Grade 10 Math C Final Exam of Semester 1 Year 2026-2027",
    "班級 Class ：  座號 No.：  姓名 Name:  總分 Score:",
  ];
  const BODY = mathExam(12);
  const structureOf = (lines: string[]) => analyzeStructure(doc(lines).paragraphs);

  it("單獨一行 Complete all the questions.", () => {
    const lines = [...TITLE, "Complete all the questions.", ...BODY];
    const s = structureOf(lines);
    expect(s.sections.map((x) => x.title)).toEqual([BODY[0], PART2]);
    expect(s.instructions).toEqual([3]);
    expect(scoreIssues(lines)).toEqual([]);
  });

  it("Word 自動編號的作答說明（1. 2. 3.）", () => {
    const lines = [
      ...TITLE,
      "作答說明 / Instructions",
      "1. Complete all the questions.",
      "2. Write your answers on the answer sheet.",
      "3. 本試卷共 4 頁，請檢查是否缺頁。",
      "4. Calculators are not allowed.",
      ...BODY,
    ];
    const s = structureOf(lines);
    expect(s.sections.map((x) => x.title)).toEqual([BODY[0], PART2]);
    expect(s.sections.map((x) => x.questions.length)).toEqual([20, 12]);
    expect(s.instructions).toEqual([3, 4, 5, 6, 7]);
    expect(scoreIssues(lines)).toEqual([]);
  });

  it("說明沒有標題、直接用編號列在第一個大題之前", () => {
    const lines = [...TITLE, "1. Complete all the questions.", "2. Show all your work.", ...BODY];
    expect(structureOf(lines).sections.map((x) => x.questions.length)).toEqual([20, 12]);
    expect(scoreIssues(lines)).toEqual([]);
  });

  it("說明夾在大題底下或題目之間：不算題目，也不併進前一題", () => {
    const lines = [
      "Part 1 Multiple choice (每題40分，共40分)",
      "Complete all the questions.",
      "1. 題目一",
      "Part 2 Calculation (每題30分，共60分)",
      "Directions: Show all your work.",
      "1. Complete all the questions.",
      "1. 計算 3 + 4",
      "Write your answers on the answer sheet.",
      "2. 計算 5 + 6",
    ];
    const s = structureOf(lines);
    expect(s.sections.map((x) => x.questions.map((q) => q.number))).toEqual([[1], [1, 2]]);
    expect(s.sections[1].questions[0].body.trim()).toBe("計算 3 + 4");
    expect(s.instructions).toEqual([1, 4, 5, 7]);
    expect(scoreIssues(lines)).toEqual([]);
  });

  it.each([
    "1. Complete the table below.",
    "1. Answer the following: what is 3 + 4?",
    "1. 說明：請說明光合作用的過程。",
    "1. Read the passage and answer questions 2–4.",
  ])("真正的題目不會被誤判成說明：%s", (q) => {
    const s = structureOf(["Part 1 Short answer (共10分)", q]);
    expect(s.instructions).toEqual([]);
    expect(s.sections[0].questions).toHaveLength(1);
  });

  it("沒有任何大題標題時，編號題目照常計算", () => {
    const s = structureOf(["Complete all the questions.", "1. 題目一", "2. 題目二"]);
    expect(s.sections[0].questions.map((q) => q.number)).toEqual([1, 2]);
    expect(s.instructions).toEqual([0]);
  });
});

describe("每格配分：每格 2 分、共 30 分｜2 points each, 30 points total（不數格子）", () => {
  const SCORE = "每格 2 分、共 30 分｜2 points each, 30 points total";
  const rule30 = parseRuleSet('[[rules]]\nid = "total-score"\ntype = "score-sum"\nexpect = 30');
  const issues30 = (lines: string[]) => lint(doc(lines), rule30).issues.map((i) => i.message);
  const fill = (n: number, boxes = 1) =>
    range(1, n).map((k) => `${k}. ${range(1, boxes).map((b) => `${k} + ${b} = ( )`).join("，")}。`);

  it("寫在大題標題裡：讀成每格 2 分、共 30 分", () => {
    const s = analyzeStructure(doc([`三、填充題（${SCORE}）`, ...fill(1)]).paragraphs).sections[0];
    expect(s).toMatchObject({ kind: "fill", perBlankScore: 2, perQuestionScore: undefined, totalScore: 30 });
    expect(sectionScore(s).total).toBe(30);
  });

  it("寫在大題標題的下一行：也讀得到，且那一行不算題目", () => {
    const s = analyzeStructure(doc(["Part 3 填充題 / Fill in the blanks", SCORE, ...fill(1)]).paragraphs).sections[0];
    expect(s).toMatchObject({ kind: "fill", perBlankScore: 2, totalScore: 30, scoreLines: [1] });
    expect(s.questions).toHaveLength(1);
  });

  it.each([
    ["1 題", fill(1)],
    ["14 題各 1 格", fill(14)],
    ["8 題各 2 格", fill(8, 2)],
    ["沒有任何格子（用 ____）", ["1. 3 + 4 = ____。", "2. 5 + 6 = ____。"]],
  ])("不數格子，大題分數一律採用「共 30 分」：%s", (_, questions) => {
    expect(issues30([`三、填充題（${SCORE}）`, ...(questions as string[])])).toEqual([]);
  });

  it("「2 points each」不會被當成每題 2 分 × 題數", () => {
    const s = analyzeStructure(doc([`三、填充題（${SCORE}）`, ...fill(3)]).paragraphs).sections[0];
    expect(s.perQuestionScore).toBeUndefined();
    expect(sectionScore(s)).toEqual({ total: 30 });
  });

  it("全卷總分仍然要檢查：40 + 30 + 20 = 90 ≠ 100", () => {
    const lines = [
      "Part 1 選擇題 / Multiple choice (每題2分，共40分)",
      ...range(1, 20).map((n) => `( )${n}. 題目`),
      `Part 2 填充題 / Fill in the blanks (${SCORE})`,
      ...fill(15),
      "Part 3 計算題 / Calculation (每題5分，共20分)",
      ...range(1, 4).map((n) => `${n}. 計算`),
    ];
    expect(scoreIssues(lines)).toEqual(["全卷配分加總為 90 分，應為 100 分"]);
  });

  it("全卷總分 40 + 30 + 30 = 100 → 通過", () => {
    const lines = [
      "Part 1 選擇題 / Multiple choice (每題2分，共40分)",
      ...range(1, 20).map((n) => `( )${n}. 題目`),
      `Part 2 填充題 / Fill in the blanks (${SCORE})`,
      ...fill(15),
      "Part 3 計算題 / Calculation (每題5分，共30分)",
      ...range(1, 6).map((n) => `${n}. 計算`),
    ];
    expect(scoreIssues(lines)).toEqual([]);
  });

  it("只寫每格、沒寫共幾分：無法判斷大題配分", () => {
    expect(issues30(["三、填充題（每格 2 分）", ...fill(15)])).toEqual([
      expect.stringContaining("無法判斷「三、填充題（每格 2 分）」的配分"),
    ]);
  });

  it.each(["每格 2 分，共 30 分", "每空 2 分，共 30 分", "每個空格 2 分，共 30 分", "2 points per blank, 30 points total"])(
    "其他寫法：%s",
    (score) => {
      const s = analyzeStructure(doc([`三、填充題（${score}）`, "1. ( )"]).paragraphs).sections[0];
      expect(s).toMatchObject({ perBlankScore: 2, totalScore: 30 });
    },
  );

  it("每題的寫法照舊算題數：每題 5 分、共 60 分｜5 points each, 60 points total", () => {
    const s = analyzeStructure(doc(["Part 2 計算題", "每題 5 分、共 60 分｜5 points each, 60 points total", "1. 計算"]).paragraphs)
      .sections[0];
    expect(s).toMatchObject({ perQuestionScore: 5, perBlankScore: undefined, totalScore: 60 });
  });

  it.each([
    ["每格 2 分、共 30 分｜3 points each, 30 points total", "分數不一致"],
    ["每格 2 分、共 30 分｜2 points each, 40 points total", "總分不一致"],
    ["每格 2 分、共 30 分｜2 points per question, 30 points total", "單位不一致"],
  ])("中英文配分說明不一致會回報：%s", (score, expected) => {
    const msgs = issues30(["Part 3 填充題 / Fill in the blanks", score, ...fill(15)]);
    expect(msgs.some((m) => m.includes("配分說明中英文不一致") && m.includes(expected))).toBe(true);
  });
});

describe("阿拉伯數字編號的大題：「1、 計算題：每題3分，共15分」", () => {
  const CALC = "1、 計算題：每題3分，共15分";
  const DRAW = "2、 畫畫看、做做看: 每題5分，共15分";

  it.each([
    [CALC, "short", 3, 15],
    [DRAW, "other", 5, 15],
    ["3、填充題（每格2分，共20分）", "fill", undefined, 20],
  ])("辨識為大題並讀出配分：%s", (title, kind, per, total) => {
    const s = analyzeStructure(doc([title, "1. 題目"]).paragraphs).sections;
    expect(s).toHaveLength(1);
    expect(s[0]).toMatchObject({ title, kind, perQuestionScore: per, totalScore: total });
    expect(s[0].questions.map((q) => q.number)).toEqual([1]);
  });

  it("兩個大題各自收題目，總分 15 + 15 = 30", () => {
    const lines = [CALC, ...range(1, 5).map((n) => `${n}. 計算 ${n} + ${n}`), DRAW, ...range(1, 3).map((n) => `${n}. 畫出圖形 ${n}`)];
    const s = analyzeStructure(doc(lines).paragraphs).sections;
    expect(s.map((x) => [x.title, x.questions.length])).toEqual([
      [CALC, 5],
      [DRAW, 3],
    ]);
    const rule30 = parseRuleSet('[[rules]]\nid = "total-score"\ntype = "score-sum"\nexpect = 30');
    expect(lint(doc(lines), rule30).issues).toEqual([]);
  });

  it.each([
    "1. 小明有 3 顆蘋果，每顆 5 元，共幾元？",
    "1、 計算：3 + 4 = ?",
    "2. 下列敘述何者正確：共有幾個？",
    "1. 小華說：每天跑步 5 分鐘。",
  ])("一般題目不會被當成大題：%s", (text) => {
    const s = analyzeStructure(doc(["一、選擇題（共 10 分）", text]).paragraphs).sections;
    expect(s).toHaveLength(1);
    expect(s[0].questions).toHaveLength(1);
  });
});

describe("大題標題段落裡有手動換行（Shift+Enter）", () => {
  const TITLE = "一、在正確的□ 中打勾。每題4分，共20分\nChoose the correct answer. Put a ✓ in the correct box.";

  it("中文標題換行接英文說明：仍辨識為大題並讀出配分", () => {
    const s = analyzeStructure(doc([TITLE, "1. 題目", "2. 題目"]).paragraphs).sections;
    expect(s).toHaveLength(1);
    expect(s[0]).toMatchObject({ title: TITLE, perQuestionScore: 4, totalScore: 20 });
    expect(s[0].questions).toHaveLength(2);
  });

  it("阿拉伯數字編號的大題也一樣", () => {
    const title = "1、 計算題：每題3分，共15分\nCalculation: 3 points each, 15 points total";
    const s = analyzeStructure(doc([title, "1. 計算"]).paragraphs).sections;
    expect(s[0]).toMatchObject({ title, kind: "short", perQuestionScore: 3, totalScore: 15 });
  });
});

describe("題號後面直接接數字：「8.3盆仙人掌…」", () => {
  const lines = [
    "四、應用題",
    "7.有7200顆糖果，每75顆裝1罐，每24罐裝1箱，共可裝幾箱？ （4分）",
    "8.3盆仙人掌小盆栽賣270元，定禾買9盆付了1000元，可以找回幾元？（5分）",
    "9.1塊香草蛋糕賣48元，1塊栗子蛋糕賣72元，媽媽各買了18塊，共花幾元？（5分）",
    "10.長方體所有表面的面積共是幾平方公分？（5分）",
  ];

  it("剛好是下一題的題號：當成新題目", () => {
    const s = analyzeStructure(doc(lines).paragraphs).sections[0];
    expect(s.questions.map((q) => q.number)).toEqual([7, 8, 9, 10]);
    expect(s.questions[1].body.startsWith("3盆仙人掌")).toBe(true);
  });

  it("大題第一題也可以：1.5公尺長的繩子…", () => {
    const s = analyzeStructure(doc(["四、應用題", "1.5公尺長的繩子剪成3段，每段幾公尺？", "2.算算看"]).paragraphs).sections[0];
    expect(s.questions.map((q) => q.number)).toEqual([1, 2]);
  });

  it.each([
    ["題目內文換行後的小數", ["四、應用題", "1. 繩子的長度是多少？", "3.5 公尺減掉 1 公尺"]],
    ["不連號的小數", ["四、應用題", "1. 繩子的長度是多少？", "2. 計算", "8.5 + 1 = ?"]],
  ])("不是下一題的題號：仍當成小數（%s）", (_, ls) => {
    const s = analyzeStructure(doc(ls as string[]).paragraphs).sections[0];
    expect(s.questions.length).toBe((ls as string[]).length - 2);
  });
});
