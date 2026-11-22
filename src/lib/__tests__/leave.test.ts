import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from "docx";
import { estimateDuration, fillLeaveForm, formatDuration, formatPeriod, inspectLeaveForm, LeaveFormError, leaveFileName, type LeaveInfo } from "../leave/leave";

const box = () => new TextRun({ text: "¨", font: "Wingdings" });
const para = (...runs: (string | TextRun)[]) =>
  new Paragraph({ children: runs.map((r) => (typeof r === "string" ? new TextRun(r) : r)) });
const cell = (paras: Paragraph[], columnSpan?: number) => new TableCell({ children: paras, columnSpan });

/** 跟真實請假單一樣的表格結構（簡化版） */
async function form(): Promise<Uint8Array> {
  const table = new Table({
    rows: [
      new TableRow({
        children: [
          cell([para("Họ tên"), para("姓名"), para("Full name:")]),
          cell([para()]),
          cell([para("Đơn vị"), para("單位"), para("Dept：")]),
          cell([para()]),
          cell([para("Chức vụ"), para("職務"), para("Position:")]),
          cell([para()]),
        ],
      }),
      new TableRow({
        children: [
          cell([para(box(), " Việc riêng 事假 Personal leave"), para(box(), "Nghỉ bù", "補假", "Compensatory leave  ", box(), "Hôn lễ", "婚假")], 2),
          cell([para(box(), " Phép bệnh", "病假", "Sick-leave")], 2),
          cell([para(box(), " Khác ", "其他", " Other                ")], 2),
        ],
      }),
      new TableRow({
        children: [cell([para("請假時間與日數", "time and days", "：")], 3), cell([para("縂", "total", "：")], 3)],
      }),
      new TableRow({ children: [cell([para("請假事由", "Reason absent", "：")], 3), cell([para("Việc riêng ")], 3)] }),
      new TableRow({ children: [cell([para("Người xin phép"), para("申請人"), para("Applicant")], 2), cell([para("代理人")], 4)] }),
      new TableRow({ children: [cell([para()], 2), cell([para()], 4)] }),
    ],
  });
  const doc = new Document({ sections: [{ children: [para("請假單 APPLICATION FOR LEAVE"), table] }] });
  return new Uint8Array(await Packer.toBuffer(doc));
}

const info: LeaveInfo = {
  name: "王小明",
  dept: "國小部",
  position: "教師",
  type: "其他",
  otherText: "防疫假",
  period: "2026/11/23 08:00 ~ 2026/11/24 17:00",
  total: "2 天",
  reason: "感冒就醫\n附診斷證明",
  applicant: "王小明",
  date: "2026/11/22",
};

async function documentXml(bytes: Uint8Array) {
  return (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
}

async function rows(bytes: Uint8Array) {
  const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
  const text = (s: string) => [...s.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");
  return [...xml.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)].map((r) =>
    [...r[0].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((c) => ({
      text: text(c[0]),
      paras: [...c[0].matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map((p) => text(p[0])),
    })),
  );
}

describe("請假單", () => {
  it("讀出版型裡的假別與預填值", async () => {
    const t = await inspectLeaveForm(await form());
    expect(t.types).toEqual(["事假", "病假", "補假", "婚假", "其他"]);
    expect(t.prefill.reason).toBe("Việc riêng");
    expect(t.prefill.name).toBe("");
  });

  it("填入欄位、勾選假別、簽名下方押日期", async () => {
    const r = await rows(await fillLeaveForm(await form(), info));
    expect(r[0][1].text).toBe("王小明");
    expect(r[0][3].text).toBe("國小部");
    expect(r[0][5].text).toBe("教師");
    // 只有「其他」打勾，其他假別維持空框
    expect(r[1][2].text).toBe("þ Khác 其他 Other：防疫假");
    expect(r[1][1].text).toMatch(/^¨ Phép bệnh病假/);
    expect(r[1][0].text.match(/¨/g)).toHaveLength(3);
    // 時間與天數都寫在 total 格，「請假時間與日數」格只留標籤
    expect(r[2][0].paras).toEqual(["請假時間與日數time and days："]);
    expect(r[2][1].paras).toEqual(["縂total： 2 天", "2026/11/23 08:00 ~ 2026/11/24 17:00"]);
    expect(r[3][1].paras).toEqual(["感冒就醫", "附診斷證明"]);
    expect(r[5][0].paras).toEqual(["王小明", "2026/11/22"]);
    expect(r[5][1].text).toBe("");
  });

  it("重複填寫不會累積舊值", async () => {
    const once = await fillLeaveForm(await form(), info);
    const twice = await fillLeaveForm(once, { ...info, type: "事假", total: "1 天", period: "2026/12/01" });
    const r = await rows(twice);
    expect(r[2][0].paras).toEqual(["請假時間與日數time and days："]);
    expect(r[2][1].paras).toEqual(["縂total： 1 天", "2026/12/01"]);
    expect(r[1][0].text).toMatch(/^þ/);
    expect(r[1][1].text).toMatch(/^¨/);
    const again = await inspectLeaveForm(twice);
    expect(again.prefill.applicant).toBe("王小明");
    expect(again.prefill.type).toBe("事假");
    // 換成事假後，「其他」後面補的說明要拿掉
    expect(r[1][2].text).toBe("¨ Khác 其他 Other");
  });

  it("維持一頁：日期字小、行距緊湊、簽名列扣回多出來的高度", async () => {
    const blank = await documentXml(await form());
    const xml = await documentXml(await fillLeaveForm(await form(), info));
    const sigRow = (x: string) => [...x.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)][5][0];
    const sigCell = [...sigRow(xml).matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)][0][0];
    const [nameP, dateP] = [...sigCell.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map((m) => m[0]);
    expect(dateP).toContain('<w:sz w:val="18"/>');
    expect(nameP).not.toContain("<w:sz ");
    expect(nameP).toContain('<w:snapToGrid w:val="0"/>');
    // pPr 子元素順序要對（snapToGrid 在 spacing 前、jc 在 rPr 前），Word 才不會說檔案損毀
    expect(nameP).toMatch(/<w:snapToGrid[^>]*\/><w:spacing[^>]*\/>.*<w:jc /);
    // 測試版型的列沒有設定高度，不會調整
    const h = (x: string) => Number(/<w:trHeight w:val="(\d+)"/.exec(sigRow(x))?.[1] ?? 0);
    expect(h(xml)).toBe(h(blank));
  });

  it("不是請假單時給清楚的錯誤", async () => {
    const doc = new Document({ sections: [{ children: [para("一般文件")] }] });
    const bytes = new Uint8Array(await Packer.toBuffer(doc));
    await expect(inspectLeaveForm(bytes)).rejects.toThrow(LeaveFormError);
    await expect(inspectLeaveForm(new Uint8Array([1, 2, 3]))).rejects.toThrow(/無法開啟/);
  });

  it("依上班時間 07:30–16:30 計算天數與小時", () => {
    const est = (sd: string, st: string, ed: string, et: string) => {
      const d = estimateDuration(sd, st, ed, et);
      return d && formatDuration(d);
    };
    // 2026-11-23 是星期一
    expect(est("2026-11-23", "07:30", "2026-11-23", "16:30")).toBe("1 天");
    expect(est("2026-11-23", "07:30", "2026-11-23", "08:30")).toBe("1 小時");
    expect(est("2026-11-23", "07:30", "2026-11-23", "07:45")).toBe("1 小時"); // 最少 1 小時
    expect(est("2026-11-23", "13:30", "2026-11-23", "16:30")).toBe("3 小時");
    expect(est("2026-11-23", "06:00", "2026-11-23", "18:00")).toBe("1 天"); // 上班時間以外不算
    expect(est("2026-11-23", "07:30", "2026-11-25", "10:30")).toBe("2 天 3 小時");
    expect(est("2026-11-23", "13:30", "2026-11-24", "13:30")).toBe("1 天"); // 3 + 6 小時 = 9 小時 = 1 天
    expect(est("2026-11-27", "07:30", "2026-11-30", "16:30")).toBe("4 天"); // 週五到週一，週末照算
    expect(est("2026-11-15", "07:30", "2026-11-15", "16:30")).toBe("1 天"); // 週日照上班時間算
    expect(est("2026-11-14", "13:30", "2026-11-15", "10:30")).toBe("6 小時"); // 週六 3 小時 + 週日 3 小時
    expect(est("2026-11-25", "", "2026-11-23", "")).toBeNull();
    expect(est("2026-11-23", "17:00", "2026-11-23", "18:00")).toBeNull();
  });

  it("請假時間：同一天只寫一次日期，同一年結束日不寫年份", () => {
    expect(formatPeriod("2026-11-23", "08:00", "2026-11-23", "12:00")).toBe("2026/11/23 08:00 ~ 12:00");
    expect(formatPeriod("2026-11-23", "07:30", "2026-11-25", "10:30")).toBe("2026/11/23 07:30 ~ 11/25 10:30");
    expect(formatPeriod("2026-12-30", "07:30", "2027-01-02", "16:30")).toBe("2026/12/30 07:30 ~ 2027/01/02 16:30");
    expect(formatPeriod("2026-11-23", "", "", "")).toBe("2026/11/23");
  });

  it("檔名", () => {
    expect(leaveFileName({ name: "王/小明", date: "2026/11/22" })).toBe("請假單-王小明-20261122.docx");
  });

  it("內建的 F-QL-013G 請假單（public/leave/）讀得懂、填得進去", async () => {
    const data = readFileSync("public/leave/F-QL-013G-leave-form.docx");
    const t = await inspectLeaveForm(data);
    expect(t.types).toHaveLength(10);
    expect(t.prefill.name).toBe("");
    const r = await rows(await fillLeaveForm(data, info));
    expect(r[0][1].text).toBe("王小明");
    expect(r[2][1].paras).toEqual(["縂total： 2 天", "2026/11/23 08:00 ~ 2026/11/24 17:00"]);
    expect(r[5][0].paras).toEqual(["王小明", "2026/11/22"]);
  });

  // 其他請假單版型（例如改版前後的 Word 檔）：本機測試時用 LEAVE_DOCX=/path/to/請假單.docx npm test
  it.skipIf(!process.env.LEAVE_DOCX)("真實請假單版型", async () => {
    const data = readFileSync(process.env.LEAVE_DOCX!);
    const t = await inspectLeaveForm(data);
    expect(t.types).toHaveLength(10);
    const out = await fillLeaveForm(data, info);
    const r = await rows(out);
    expect(r[0][1].text).toBe("王小明");
    expect(r[5][0].paras).toEqual(["王小明", "2026/11/22"]);
    // 簽名列：1376 - 240（時間）- 290（事由兩行）
    const heights = (x: string) => [...x.matchAll(/<w:trHeight w:val="(\d+)"/g)].map((m) => Number(m[1]));
    const before = heights(await documentXml(data));
    const after = heights(await documentXml(out));
    // 時間那一列 420 → 720；簽名列扣回 300，再扣事由多一行 290：1376 - 590 = 786 → 最少 800
    expect(after[1]).toBe(720);
    expect(after.at(-1)).toBe(800);
    // 加起來不會比原本高（1376 + 420 → 800 + 720）
    expect(after.reduce((a, b) => a + b)).toBeLessThanOrEqual(before.reduce((a, b) => a + b));
    // 重填不會再調
    expect(heights(await documentXml(await fillLeaveForm(out, info)))).toEqual(after);
    // 總計格兩行，都是原本的字級（放得下，不縮小）
    expect(r[2][1].paras).toEqual(["縂total： 2 天", "2026/11/23 08:00 ~ 2026/11/24 17:00"]);
    expect(await documentXml(out)).not.toMatch(/<w:sz w:val="\d+"\/><w:szCs w:val="\d+"\/>(?:(?!<\/w:r>)[\s\S])*2026\/11\/23/);
    if (process.env.LEAVE_OUT) (await import("node:fs")).writeFileSync(process.env.LEAVE_OUT, out);
  });
});
