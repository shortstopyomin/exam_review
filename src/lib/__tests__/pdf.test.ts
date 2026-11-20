/**
 * PDF 解析：用假的 pdf.js 餵文字項目，測試組行、頁碼、換行接續的邏輯。
 */
import { describe, expect, it } from "vitest";
import { isPdf, parsePdf, PdfParseError, type PdfLib } from "../pdf/parse";
import { analyzeStructure, sectionScore } from "../structure";

interface Item {
  str: string;
  x: number;
  y: number;
  w?: number;
  h?: number;
  eol?: boolean;
}

/** 每頁 A4（595 × 842），一行一個項目；[文字, y] 或完整 Item */
function fakeLib(pages: (Item | [string, number])[][]): PdfLib {
  return {
    getDocument: () => ({
      promise: Promise.resolve({
        numPages: pages.length,
        getPage: async (n: number) => ({
          view: [0, 0, 595, 842],
          getTextContent: async () => ({
            items: pages[n - 1].map((it) => {
              const i: Item = Array.isArray(it) ? { str: it[0], x: 35, y: it[1], eol: true } : it;
              const h = i.h ?? 14;
              return { str: i.str, transform: [h, 0, 0, h, i.x, i.y], width: i.w ?? i.str.length * h, height: h, hasEOL: i.eol ?? false };
            }),
          }),
        }),
      }),
      destroy: async () => {},
    }),
  };
}

const parse = (pages: (Item | [string, number])[][]) => parsePdf(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), "t.pdf", fakeLib(pages));
const texts = async (pages: (Item | [string, number])[][]) => (await parse(pages)).paragraphs.map((p) => p.text);

describe("PDF 解析", () => {
  it("用檔頭判斷 PDF", () => {
    expect(isPdf(new TextEncoder().encode("%PDF-1.7\n..."))).toBe(true);
    expect(isPdf(new Uint8Array([0x50, 0x4b, 3, 4]))).toBe(false); // docx（zip）
  });

  it("同一行的多個項目接成一行，有空隙時補空白", async () => {
    const lines = await texts([
      [
        { str: "", x: 35, y: 565, eol: true },
        { str: "一、", x: 35, y: 565, w: 28 },
        { str: "在正確的", x: 64, y: 565, w: 56 },
        { str: "每題", x: 184, y: 565, w: 20, h: 10 },
        { str: "4", x: 207, y: 565, w: 5, h: 10 },
        { str: "分", x: 216, y: 565, w: 10, h: 10, eol: true },
        { str: "1.", x: 35, y: 489, w: 11 },
        { str: "九千零八萬", x: 50, y: 489, w: 70, eol: true },
      ],
    ]);
    expect(lines).toEqual(["一、在正確的 每題 4 分", "1. 九千零八萬"]);
  });

  it("頁首、頁尾的頁碼不算內容；內文裡的數字行保留", async () => {
    const lines = await texts([
      [["1", 45], ["一、選擇題（共 10 分）", 700], ["1. 題目", 680], ["10", 600]],
      [["- 2 -", 820], ["2. 題目", 700], ["第 2 頁", 40]],
    ]);
    expect(lines).toEqual(["一、選擇題（共 10 分）", "1. 題目", "10", "2. 題目"]);
  });

  it("配分被換行切開：「共 20」⏎「分 Word Problems」接回同一行", async () => {
    const doc = await parse([[["六、應用題:把做法寫出來。每題 5 分，共 20", 700], ["分 Word Problems: Show your work.", 682], ["⑴ 計算", 650]]]);
    const s = analyzeStructure(doc.paragraphs).sections;
    expect(s[0]).toMatchObject({ perQuestionScore: 5, totalScore: 20 });
    expect(sectionScore(s[0]).total).toBe(20);
  });

  it("換行剛好以數字開頭（「…number is⏎34. Her classmate…」）不算新題目", async () => {
    const doc = await parse([
      [
        ["三、回答問題。每題 1 分，共 10 分", 760],
        ["1. 小兔的車票號碼是 34 號，Rabbit's ticket number is", 740],
        ["34. Her classmate, Monkey, is sitting", 722],
        ["2. 下一題？", 690],
      ],
    ]);
    const s = analyzeStructure(doc.paragraphs).sections[0];
    expect(s.questions.map((q) => q.number)).toEqual([1, 2]);
    expect(s.questions[0].body).toContain("34. Her classmate");
  });

  it("上一行沒有句尾標點，但這行是下一題：仍是新題目", async () => {
    const doc = await parse([[["一、選擇題（共 10 分）", 760], ["1. 題目", 740], ["□ 70 度 □ 110 度", 722], ["2. ∠ㄆ是幾度？", 704]]]);
    expect(analyzeStructure(doc.paragraphs).sections[0].questions.map((q) => q.number)).toEqual([1, 2]);
  });

  it("沒有文字（掃描檔）：回報錯誤", async () => {
    await expect(parse([[]])).rejects.toThrow(PdfParseError);
  });
});
