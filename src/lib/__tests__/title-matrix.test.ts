/**
 * 標題檢查的格式矩陣：確認各種中英文寫法都能正確辨識
 * 「第 1 學期」與「期末 / Final」。
 * 設定：今天 2026-11-19（第 1 學期）、examType = "final"。
 */
import { describe, expect, it } from "vitest";
import { lint, parseRuleSet } from "../rules/engine";
import type { ParsedDocument } from "../types";

const set = parseRuleSet(`[[rules]]
id = "t"
type = "exam-title"
schoolZh = "越華國際學校"
schoolEn = "Viet Hoa International School"
examType = "final"
today = "2026-11-19"`);

const ZH = (semester: string, exam: string) =>
  `越華國際學校 2026-2027 學年度 ${semester} 10年級 中文數學 ${exam}`;
const EN = (semester: string, exam: string) =>
  `Viet Hoa International School Grade 10 Math C ${exam} ${semester} Year 2026-2027`;

function issues(lines: string[]) {
  const doc: ParsedDocument = {
    fileName: "t.docx",
    paragraphs: lines.map((text, index) => ({ index, text, inTable: false, hasMath: false, imageCount: 0 })),
  };
  return lint(doc, set).issues.map((i) => i.message);
}

const OK_ZH_SEMESTER = ["第1學期", "第 1 學期", "第 1學期", "第一學期", "第１學期", "上學期"];
const OK_EN_SEMESTER = [
  "of Semester 1", "Semester 1", "Semester I", "Semester One", "1st Semester", "First Semester",
  "first semester", "SEMESTER 1", "Fall Semester", "Sem 1", "Sem. 1", "Term 1", "1st Term", "First Term", "S1",
];
const OK_ZH_EXAM = ["期末評量試卷", "期末考", "期末考試", "期末測驗", "期末評量"];
const OK_EN_EXAM = [
  "Final Exam", "Final Examination", "Finals", "final exam", "FINAL EXAM", "Final Test", "Final Assessment",
  "End-of-Semester Exam", "End of Semester Exam", "End-of-Term Exam", "End of Term Exam",
];

const BAD_ZH_SEMESTER = ["第2學期", "第 2 學期", "第二學期", "第２學期", "下學期"];
const BAD_EN_SEMESTER = [
  "Semester 2", "Semester II", "Semester Two", "2nd Semester", "Second Semester", "Spring Semester",
  "Sem 2", "Term 2", "2nd Term", "S2",
];
const BAD_ZH_EXAM = ["期中評量試卷", "期中考", "期中測驗"];
const BAD_EN_EXAM = ["Midterm Exam", "Mid-term Exam", "Mid Term Exam", "MIDTERM", "Mid-Semester Exam", "Mid-Term Test"];

describe("正確寫法不誤報（第 1 學期、期末）", () => {
  it.each(OK_ZH_SEMESTER)("中文學期：%s", (s) => {
    expect(issues([ZH(s, "期末評量試卷"), EN("of Semester 1", "Final Exam")])).toEqual([]);
  });
  it.each(OK_EN_SEMESTER)("英文學期：%s", (s) => {
    expect(issues([ZH("第 1 學期", "期末評量試卷"), EN(s, "Final Exam")])).toEqual([]);
  });
  it.each(OK_ZH_EXAM)("中文考試類型：%s", (e) => {
    expect(issues([ZH("第 1 學期", e), EN("of Semester 1", "Final Exam")])).toEqual([]);
  });
  it.each(OK_EN_EXAM)("英文考試類型：%s", (e) => {
    expect(issues([ZH("第 1 學期", "期末評量試卷"), EN("of Semester 1", e)])).toEqual([]);
  });
});

/**
 * 上面「不誤報」的測試，即使某種寫法完全沒被辨識也會通過。
 * 這裡讓另一種語言留白，證明每種寫法都真的有被認出來：
 * 若沒認出來，會多出「找不到學期／考試類型」的訊息。
 */
describe("正確寫法確實有被辨識", () => {
  it.each(OK_ZH_SEMESTER)("中文學期：%s", (s) => {
    expect(issues([ZH(s, "期末評量試卷"), EN("", "Final Exam")])).toEqual([
      "英文標題找不到學期（例如「Semester 1」）",
    ]);
  });
  it.each(OK_EN_SEMESTER)("英文學期：%s", (s) => {
    expect(issues([ZH("", "期末評量試卷"), EN(s, "Final Exam")])).toEqual([
      "中文標題找不到學期（例如「第 1 學期」）",
    ]);
  });
  it.each(OK_ZH_EXAM)("中文考試類型：%s", (e) => {
    expect(issues([ZH("第 1 學期", e), EN("of Semester 1", "Exam")])).toEqual(["英文標題找不到考試類型「Final」"]);
  });
  it.each(OK_EN_EXAM)("英文考試類型：%s", (e) => {
    expect(issues([ZH("第 1 學期", "評量試卷"), EN("of Semester 1", e)])).toEqual(["中文標題找不到考試類型「期末」"]);
  });
});

describe("不會誤判", () => {
  it("Mid Term 2 Exam 的 Term 2 不是學期", () => {
    const msgs = issues([ZH("第 1 學期", "期末評量試卷"), EN("of Semester 1", "Mid Term 2 Exam")]);
    expect(msgs).toEqual(["考試類型「Mid Term」應為「Final」"]);
  });
  it("中英文都寫錯時，兩邊都回報", () => {
    expect(issues([ZH("第 2 學期", "期中考"), EN("Semester 2", "Midterm Exam")])).toEqual([
      "「第 2 學期」應為第 1 學期",
      "考試類型「期中」應為「期末」",
      "「Semester 2」應為第 1 學期",
      "考試類型「Midterm」應為「Final」",
    ]);
  });
});

describe("錯誤寫法都抓得到", () => {
  it.each(BAD_ZH_SEMESTER)("中文學期：%s", (s) => {
    expect(issues([ZH(s, "期末評量試卷"), EN("of Semester 1", "Final Exam")])).toEqual([
      expect.stringMatching(/應為第 1 學期$/),
    ]);
  });
  it.each(BAD_EN_SEMESTER)("英文學期：%s", (s) => {
    expect(issues([ZH("第 1 學期", "期末評量試卷"), EN(s, "Final Exam")])).toEqual([
      expect.stringMatching(/應為第 1 學期$/),
    ]);
  });
  it.each(BAD_ZH_EXAM)("中文考試類型：%s", (e) => {
    expect(issues([ZH("第 1 學期", e), EN("of Semester 1", "Final Exam")])).toEqual(["考試類型「期中」應為「期末」"]);
  });
  it.each(BAD_EN_EXAM)("英文考試類型：%s", (e) => {
    expect(issues([ZH("第 1 學期", "期末評量試卷"), EN("of Semester 1", e)])).toEqual([
      expect.stringMatching(/^考試類型「.+」應為「Final」$/),
    ]);
  });
});

describe("缺少資訊時提醒", () => {
  it("中文標題沒寫期末", () => {
    expect(issues([ZH("第 1 學期", "評量試卷"), EN("of Semester 1", "Final Exam")])).toEqual([
      "中文標題找不到考試類型「期末」",
    ]);
  });
  it("英文標題沒寫 Final", () => {
    expect(issues([ZH("第 1 學期", "期末評量試卷"), EN("of Semester 1", "Exam")])).toEqual([
      "英文標題找不到考試類型「Final」",
    ]);
  });
  it("中英文都沒寫學期", () => {
    expect(issues([ZH("", "期末評量試卷"), EN("", "Final Exam")])).toEqual([
      "標題中找不到學期（例如「第 1 學期」）",
    ]);
  });
  it("只有中文寫學期、英文沒寫", () => {
    expect(issues([ZH("第 1 學期", "期末評量試卷"), EN("", "Final Exam")])).toEqual([
      "英文標題找不到學期（例如「Semester 1」）",
    ]);
  });
});
