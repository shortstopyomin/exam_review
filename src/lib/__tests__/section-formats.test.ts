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

  it("少一題：11 題 × 5 分 = 55 分，與「共 60 分」不符", () => {
    expect(scoreIssues(mathExam(11))).toEqual([
      expect.stringContaining("標示共 60 分，但每題 5 分 × 11 題 = 55 分"),
    ]);
  });

  it("多一題：13 題 × 5 分 = 65 分，與「共 60 分」不符", () => {
    expect(scoreIssues(mathExam(13))).toEqual([
      expect.stringContaining("標示共 60 分，但每題 5 分 × 13 題 = 65 分"),
    ]);
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
