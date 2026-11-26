import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { AlignmentType, Document, LevelFormat, Packer, Paragraph, TextRun } from "docx";
import {
  composeExam,
  defaultSectionFor,
  detectExamInfo,
  fillScores,
  sectionTotal,
  detectExamType,
  inspectTemplate,
  loadDocxSource,
  markDuplicates,
  outputFileName,
  versionedFileName,
  questionPrefixLength,
} from "../compose/compose";
import { formatEqField, parseXml, previewParagraph, replaceInParagraph, W, paragraphText } from "../compose/xml";
import { parseDocx } from "../docx/parse";

const p = (text: string, opts = {}) => new Paragraph({ children: [new TextRun(text)], ...opts });
const q = (text: string, ref = "q") => new Paragraph({ children: [new TextRun(text)], numbering: { reference: ref, level: 0 } });

async function build(children: Paragraph[]): Promise<Uint8Array> {
  const doc = new Document({
    numbering: {
      config: [
        { reference: "q", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "(　　)%1.", alignment: AlignmentType.LEFT }] },
        { reference: "c", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT }] },
      ],
    },
    sections: [{ children }],
  });
  return new Uint8Array(await Packer.toBuffer(doc));
}

const template = () =>
  build([
    // 「Midterm」拆在兩個 run，測試跨 run 取代
    // 年級的「8」是紅字、獨立一個 run（跟真實版型一樣），換掉後「年級」不能變紅
    new Paragraph({
      children: [
        new TextRun("越華國際學校 2026-2027 學年度 第 1學期 "),
        new TextRun({ text: "8", color: "FF0000" }),
        new TextRun("年級 數學 期中評量試卷"),
      ],
    }),
    new Paragraph({
      children: [new TextRun("Grade 8 Math Mid"), new TextRun("term Exam of Semester 1 Year 2026-2027")],
    }),
    p("Instructions: Show your final answer clearly."),
    p("Part 1選擇題 / Multiple choice (每題4分，共40分)"),
    q("版型選擇題一"),
    p("(A)1 (B)2 (C)3 (D)4"),
    p(""),
    p("Part 2計算題和應用題 / Calculation and word problems (每題5分，共60分)"),
    q("版型計算題一", "c"),
    // 計算題留 3 行作答空間
    p(""),
    p(""),
    p(""),
  ]);

const source = () =>
  build([
    p("某某學校 期末考"),
    p("一、選擇題（每題 2 分，共 4 分）"),
    p("( B )1. 下列哪一個數是質數？"),
    p("(A) 21 (B) 27 (C) 31 (D) 33"),
    p("( )2. 若 x + 3 = 7，則 x = ？"),
    p("(A) 3 (B) 4 (C) 7 (D) 10"),
    p("二、填充題（每格 2 分，共 4 分）"),
    p("1. 15 與 20 的最大公因數為 (  )。"),
    p("2. |-7| = (  )。"),
    p("三、計算題"),
    p("1. 計算 (-3)² + 4 × (-2) 的值。（8 分）"),
    p(""),
    p(""),
  ]);

async function texts(bytes: Uint8Array) {
  const doc = await parseDocx(bytes);
  return doc.paragraphs.map((x) => x.text);
}

describe("compose", () => {
  it("reads template sections and exam type", async () => {
    const info = await inspectTemplate(await template(), "t.docx");
    expect(info.sections.map((s) => s.kind)).toEqual(["choice", "short"]);
    expect(info).toMatchObject({ examType: "midterm", startYear: 2026, semester: 1, grade: 8 });
  });

  it("extracts questions and strips typed numbers from preview", async () => {
    const src = await loadDocxSource(await source(), "s.docx", 0);
    expect(src.questions.map((x) => [x.kind, x.number])).toEqual([
      ["choice", 1],
      ["choice", 2],
      ["fill", 1],
      ["fill", 2],
      ["short", 1],
    ]);
    expect(src.questions[0].text).toBe("下列哪一個數是質數？\n(A) 21 (B) 27 (C) 31 (D) 33");
    // 作答空白帶進題目
    expect(src.questions[4].blocks).toHaveLength(3);
  });

  it("reads Word files whose XML parts start with a UTF-8 BOM", async () => {
    // 有些題庫系統匯出的 .docx，每個 XML 開頭都有 BOM，Word 開得了但 xmldom 會報錯
    const zip = await JSZip.loadAsync(await source());
    for (const f of Object.values(zip.files)) {
      if (f.dir || !/\.(xml|rels)$/.test(f.name)) continue;
      zip.file(f.name, "﻿" + (await f.async("string")));
    }
    const bom = await zip.generateAsync({ type: "uint8array" });
    const src = await loadDocxSource(bom, "bom.docx", 0);
    expect(src.questions).toHaveLength(5);
  });

  it("marks duplicates across exams, keeping the first", async () => {
    const a = await loadDocxSource(await source(), "a.docx", 0);
    const b = await loadDocxSource(await source(), "b.docx", 1);
    markDuplicates([a, b]);
    expect(a.questions.every((x) => !x.duplicateOf)).toBe(true);
    expect(b.questions.map((x) => x.duplicateOf)).toEqual(a.questions.map((x) => x.id));
  });

  it("builds the exam into the template layout", async () => {
    const tpl = await template();
    const info = await inspectTemplate(tpl, "t.docx");
    const src = await loadDocxSource(await source(), "s.docx", 0);
    const sections = info.sections.map(() => [] as typeof src.questions);
    for (const x of src.questions) sections[defaultSectionFor(x.kind, info.sections)].push(x);
    const out = await composeExam({ template: tpl, templateName: "t.docx", info: { examType: "final", startYear: 2027, semester: 2, grade: 9 }, sections, sources: [src] });
    const t = (await texts(out)).filter((x) => x.trim());
    expect(t).toEqual([
      "越華國際學校 2027-2028 學年度 第 2學期 9年級 數學 期末評量試卷",
      "Grade 9 Math Final Exam of Semester 2 Year 2027-2028",
      // 「final answer」不是考試名稱，不能改
      "Instructions: Show your final answer clearly.",
      "Part 1選擇題 / Multiple choice (每題____分，共____分)",
      "(　　)1. 下列哪一個數是質數？",
      "(A) 21 (B) 27 (C) 31 (D) 33",
      "(　　)2. 若 x + 3 = 7，則 x = ？",
      "(A) 3 (B) 4 (C) 7 (D) 10",
      "Part 2計算題和應用題 / Calculation and word problems (每題____分，共____分)",
      "1. 15 與 20 的最大公因數為 (  )。",
      "2. |-7| = (  )。",
      "3. 計算 (-3)² + 4 × (-2) 的值。（____ 分）",
    ]);
    // 每題後面的空白行數至少和版型一樣：選擇題 1 行、計算題 3 行
    const all = (await parseDocx(out)).paragraphs.map((x) => x.text.trim());
    const blanksAfter = (start: string) => {
      let i = all.findIndex((x) => x.startsWith(start));
      let n = 0;
      while (++i < all.length && !all[i]) n++;
      return n;
    };
    expect(blanksAfter("(A) 21")).toBe(1);
    expect(blanksAfter("1. 15 與 20")).toBe(3);
    expect(blanksAfter("2. |-7|")).toBe(3);
    // 原題已有 2 行空白，補到 3 行
    expect(blanksAfter("3. 計算")).toBe(3);

    // 版型自己的題目不保留
    expect(t.join("\n")).not.toContain("版型");
    // 每個大題用自己的編號清單（從 1 重新開始），不沿用來源試卷的
    const xml = await (await JSZip.loadAsync(out)).file("word/document.xml")!.async("string");
    const numIds = [...xml.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((m) => m[1]);
    expect(numIds).toHaveLength(5);
    expect(new Set(numIds.slice(0, 2)).size).toBe(1);
    expect(new Set(numIds.slice(2)).size).toBe(1);
    expect(numIds[0]).not.toBe(numIds[2]);
    // 紅字只留在數字上
    expect(xml).toMatch(/<w:color w:val="FF0000"\/><\/w:rPr><w:t[^>]*>9<\/w:t>/);
    expect(xml).toMatch(/<w:t[^>]*>年級 數學 期末評量試卷<\/w:t>/);
  });

  it("switches back to midterm", async () => {
    const tpl = await template();
    const info = await inspectTemplate(tpl, "t.docx");
    const out = await composeExam({
      template: tpl,
      templateName: "t.docx",
      info: { examType: "midterm" },
      sections: info.sections.map(() => []),
      sources: [],
    });
    const t = await texts(out);
    expect(t[0]).toContain("期中評量");
    // 沒給的欄位維持版型原樣
    expect(t[0]).toBe("越華國際學校 2026-2027 學年度 第 1學期 8年級 數學 期中評量試卷");
    expect(t[1]).toBe("Grade 8 Math Midterm Exam of Semester 1 Year 2026-2027");
  });

  it("keeps the template's typed answer bracket after an auto number", async () => {
    // 版型的自動編號只有「1.」，作答括號「(   )」是手打在後面的文字
    const tpl = await build([
      p("一、單一選擇題（每題8分，共96分）"),
      q("(   ) 版型選擇題一", "c"),
      p("二、計算題"),
      // 手打題號、括號在題號後面
      p("1. (  ) 版型計算題一"),
    ]);
    const info = await inspectTemplate(tpl, "t.docx");
    const src = await loadDocxSource(await source(), "s.docx", 0);
    const choice = src.questions.filter((x) => x.kind === "choice");
    const out = await composeExam({
      template: tpl,
      templateName: "t.docx",
      info: { examType: "final" },
      sections: [choice, src.questions.filter((x) => x.kind === "short")],
      sources: [src],
    });
    expect(info.sections).toHaveLength(2);
    const t = (await texts(out)).filter((x) => x.trim());
    // 原題的「( B )」答案不留，換成版型的空括號；題號重新從 1 開始
    expect(t).toContain("1. (   ) 下列哪一個數是質數？");
    expect(t).toContain("2. (   ) 若 x + 3 = 7，則 x = ？");
    expect(t.find((x) => x.includes("計算 (-3)²"))).toBe("1. (  ) 計算 (-3)² + 4 × (-2) 的值。（____ 分）");
  });

  it("adds the score to section titles that have no score text", async () => {
    const tpl = await build([p("一、選擇題（每格2分，共58分）"), q("版型選擇題一"), p("二、計算題"), p("1. 版型計算題一")]);
    const src = await loadDocxSource(await source(), "s.docx", 0);
    const out = await composeExam({
      template: tpl,
      templateName: "t.docx",
      info: { examType: "final" },
      sections: [src.questions.filter((x) => x.kind === "choice"), src.questions.filter((x) => x.kind === "short")],
      sources: [src],
      scores: [8, 12.5],
    });
    const t = (await texts(out)).filter((x) => x.trim());
    // 「每格」改成「每題」並填入分數；標題沒寫配分的補上
    expect(t).toContain("一、選擇題（每題8分，共16分）");
    expect(t).toContain("二、計算題（每題12.5分，共12.5分）");
  });

  it("keeps zhuyin fonts off the score text", async () => {
    // 低年級版型：整個標題用注音字型，配分文字不要注音
    const zy = (text: string) => new TextRun({ text, font: "ToneOZ-Zhuyin-Kai-Traditional" });
    const tpl = await build([
      new Paragraph({ children: [zy("一、計算題："), zy("每"), zy("題3"), zy("分，共15分。")] }),
      q("版型計算題一", "c"),
      new Paragraph({ children: [zy("二、畫畫看")] }),
      q("版型畫畫看一", "c"),
    ]);
    const src = await loadDocxSource(await source(), "s.docx", 0);
    const out = await composeExam({
      template: tpl,
      templateName: "t.docx",
      info: { examType: "final" },
      sections: [src.questions.filter((x) => x.kind === "short"), src.questions.filter((x) => x.kind === "fill")],
      sources: [src],
      scores: [4, 2],
    });
    const xml = await (await JSZip.loadAsync(out)).file("word/document.xml")!.async("string");
    const fontOf = (text: string) => {
      const run = [...xml.matchAll(/<w:r>.*?<\/w:r>/g)].map((m) => m[0]).find((r) => r.includes(`>${text}<`));
      return /w:eastAsia="([^"]+)"/.exec(run ?? "")?.[1];
    };
    expect(fontOf("一、計算題：")).toBe("ToneOZ-Zhuyin-Kai-Traditional");
    expect(fontOf("每題4分")).toBe("標楷體");
    expect(fontOf("共4分")).toBe("標楷體");
    expect(fontOf("二、畫畫看")).toBe("ToneOZ-Zhuyin-Kai-Traditional");
    // 標題沒寫配分，補上的文字也不要注音
    expect(fontOf("（每題2分，共4分）")).toBe("標楷體");
  });

  it("fills per-question scores and section totals from the picked count", async () => {
    const tpl = await template();
    const info = await inspectTemplate(tpl, "t.docx");
    // 版型寫的「每題4分」「每題5分」當預設值
    expect(info.sections.map((s) => s.perQuestionScore)).toEqual([4, 5]);
    const src = await loadDocxSource(await source(), "s.docx", 0);
    const sections = info.sections.map(() => [] as typeof src.questions);
    for (const x of src.questions) sections[defaultSectionFor(x.kind, info.sections)].push(x);
    const out = await composeExam({
      template: tpl,
      templateName: "t.docx",
      info: { examType: "final" },
      sections,
      sources: [src],
      scores: [2.5, 10],
    });
    const t = (await texts(out)).filter((x) => x.trim());
    // 選擇題 2 題 × 2.5，非選 3 題 × 10
    expect(t).toContain("Part 1選擇題 / Multiple choice (每題2.5分，共5分)");
    expect(t).toContain("Part 2計算題和應用題 / Calculation and word problems (每題10分，共30分)");
    // 題目後面的單題配分換成這次的每題配分
    expect(t).toContain("3. 計算 (-3)² + 4 × (-2) 的值。（10 分）");
  });
});

describe("compose helpers", () => {
  it("fills score text in section titles", () => {
    const fill = (text: string, per?: number, total?: number) => {
      const para = parseXml(`<w:p xmlns:w="${W}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).documentElement!;
      fillScores(para, per, total);
      return paragraphText(para);
    };
    expect(fill("一、 填充題 （每格2分，共58分）", 3, 72)).toBe("一、 填充題 （每題3分，共72分）");
    expect(fill("一、 填充題 （每格2分，共58分）")).toBe("一、 填充題 （每格____分，共____分）");
    expect(fill("Part 1 (5 points each, 60 points total)", 4, 40)).toBe("Part 1 (4 points each, 40 points total)");
    expect(fill("2 points per blank, Total: 30", 3, 9)).toBe("3 points per question, Total: 9");
    expect(fill("二、選擇題（40％，每題4分）", 5, 50)).toBe("二、選擇題（50％，每題5分）");
    expect(sectionTotal(0.1, 3)).toBe(0.3);
    expect(sectionTotal(undefined, 3)).toBeUndefined();
  });

  it("detects exam type", () => {
    expect(detectExamType("Grade 8 Final Exam")).toBe("final");
    expect(detectExamType("write your final answer")).toBeUndefined();
    expect(detectExamType("期中評量")).toBe("midterm");
  });

  it("names the output file", () => {
    expect(outputFileName("2026_2027_S1_Midterm_G8_Math數學考卷v2.docx", { examType: "final" })).toBe(
      "2026_2027_S1_Final_G8_Math數學考卷v2.docx",
    );
    expect(
      outputFileName("2026_2027_S1_Midterm_G8_Math數學考卷v2.docx", {
        examType: "final",
        startYear: 2027,
        semester: 2,
        grade: 10,
      }),
    ).toBe("2027_2028_S2_Final_G10_Math數學考卷v2.docx");
    expect(outputFileName("2026_S1_Final_G8.docx", { examType: "midterm" })).toBe("2026_S1_Midterm_G8.docx");
    // 檔名跟版型一樣時，版本號往上加：沒有的加 v1，vN 改成 v(N+1)
    expect(outputFileName("數學考卷.docx", { examType: "final" })).toBe("數學考卷v1.docx");
    expect(outputFileName("數學考卷v1.docx", { examType: "final" })).toBe("數學考卷v2.docx");
    expect(outputFileName("G10_Math_Final_V3.docx", { examType: "final" })).toBe("G10_Math_Final_V4.docx");
    expect(outputFileName("20261125_G10_Math_Final_Choice-題目卷.docx", { examType: "final" })).toBe(
      "20261125_G10_Math_Final_Choice-題目卷v1.docx",
    );
  });

  it("adds a version number when the file name was already used", () => {
    expect(versionedFileName("考卷.docx", [])).toBe("考卷.docx");
    expect(versionedFileName("考卷.docx", ["考卷.docx"])).toBe("考卷v2.docx");
    expect(versionedFileName("考卷.docx", ["考卷.docx", "考卷v2.docx"])).toBe("考卷v3.docx");
    expect(versionedFileName("G8_Math.docx", ["g8_math.docx"])).toBe("G8_Math_v2.docx");
    // 原檔名已有 vN：改成 v(N+1)
    expect(versionedFileName("Final_G8_數學考卷v2.docx", ["Final_G8_數學考卷v2.docx"])).toBe("Final_G8_數學考卷v3.docx");
    expect(versionedFileName("Math_V3.docx", ["Math_V3.docx", "Math_V4.docx"])).toBe("Math_V5.docx");
    expect(versionedFileName("考卷v2-題目卷.docx", ["考卷v2-題目卷.docx"])).toBe("考卷v3-題目卷.docx");
    // G10、S1 不是版本號
    expect(versionedFileName("S1_G10.docx", ["S1_G10.docx"])).toBe("S1_G10_v2.docx");
  });

  it("strips typed numbers and answer brackets", () => {
    const cut = (t: string) => t.slice(questionPrefixLength(t));
    expect(cut("( B )1. 下列哪一個數是質數？")).toBe("下列哪一個數是質數？");
    expect(cut("（　）將一塊邊長為a的正方形")).toBe("將一塊邊長為a的正方形");
    expect(cut("1.（ C ）若 x = 2")).toBe("若 x = 2");
    expect(cut("12．計算 3 + 4")).toBe("計算 3 + 4");
    // 選項開頭、只有括號的段落不動
    expect(cut("(A) 21 (B) 27")).toBe("(A) 21 (B) 27");
    expect(cut("（　）")).toBe("（　）");
  });

  it("detects title fields", () => {
    expect(detectExamInfo("越華國際學校 2026-2027 學年度 第 1學期 8年級 數學 期中評量試卷")).toEqual({
      examType: "midterm",
      startYear: 2026,
      semester: 1,
      grade: 8,
    });
    expect(detectExamInfo("Grade 10 Math Final Exam of Semester 2 Year 2025 - 2026")).toEqual({
      examType: "final",
      startYear: 2025,
      semester: 2,
      grade: 10,
    });
    expect(detectExamInfo("數學測驗")).toEqual({});
  });

  it("formats EQ fields", () => {
    expect(formatEqField(" EQ \\R(, 147 )")).toBe("√(147)");
    expect(formatEqField("EQ \\F(\\R(,11),3)")).toBe("(√(11))/3");
    expect(formatEqField("EQ \\R(3,x)")).toBe("³√x");
  });

  it("replaces text across runs and previews superscripts", () => {
    const doc = parseXml(
      `<w:p xmlns:w="${W}"><w:r><w:t>每題</w:t></w:r><w:r><w:t>4</w:t></w:r><w:r><w:t>分 x</w:t></w:r>` +
        `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r></w:p>`,
    );
    const para = doc.documentElement!;
    expect(previewParagraph(para)).toBe("每題4分 x²");
    replaceInParagraph(para, /(每題)(\d+)(分)/, (m) => `${m[1]}__${m[3]}`);
    expect(paragraphText(para)).toBe("每題__分 x2");
  });
});
