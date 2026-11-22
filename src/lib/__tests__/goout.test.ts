import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { fillGoOutForm, formatWordDate, goOutFileName, goOutTimeProblem, GoOutFormError, inspectGoOutForm, type GoOutInfo } from "../goout/goout";

const TEMPLATE = "public/goout/F-QL-034-going-out-form.docx";

const info: GoOutInfo = {
  name: "王小明",
  dept: "教務處",
  position: "老師",
  outDate: "2026-11-23",
  outTime: "14:30",
  backDate: "2026-11-23",
  backTime: "16:05",
  destination: "教育局",
  contact: "0901 234 567",
  reason: "送公文",
  applicant: "王小明",
  date: "2026-11-22",
};

async function documentXml(bytes: Uint8Array | Buffer) {
  return (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
}

const text = (s: string) => [...s.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");

async function rows(bytes: Uint8Array) {
  const xml = await documentXml(bytes);
  return [...xml.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)].map((r) =>
    [...r[0].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((c) => ({
      text: text(c[0]),
      paras: [...c[0].matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map((p) => text(p[0])),
      xml: c[0],
    })),
  );
}

describe("外出單", () => {
  it("Word 日期格式", () => {
    const p = { y: 2026, m: 8, d: 4, h: 14, min: 5 };
    expect(formatWordDate("dd/MM/yyyy", "vi-VN", p)).toBe("04/08/2026");
    expect(formatWordDate("dd/MM/yyyy h:mm am/pm", "vi-VN", p)).toBe("04/08/2026 2:05 CH");
    expect(formatWordDate("dd/MM/yyyy h:mm:ss am/pm", "vi-VN", { ...p, h: 9 })).toBe("04/08/2026 9:05:00 SA");
    expect(formatWordDate("yyyy/M/d HH:mm", "zh-TW", p)).toBe("2026/8/4 14:05");
    expect(formatWordDate("h:mm AM/PM", "en-US", { ...p, h: 0 })).toBe("12:05 AM");
  });

  it("內建版型：填入欄位、日期選擇器、簽名下方押日期", async () => {
    const data = readFileSync(TEMPLATE);
    const t = await inspectGoOutForm(data);
    expect(t.prefill.name).toBe("");

    const out = await fillGoOutForm(data, info);
    const r = await rows(out);
    expect(r[0][1].text).toBe("王小明");
    expect(r[0][3].text).toBe("教務處");
    expect(r[0][5].text).toBe("老師");
    // 外出時間 / 預計返回時間：日/月/年 + 24 小時制時:分，沒有秒、沒有 SA / CH
    expect(r[1][1].text).toBe("23/11/2026 14:30");
    expect(r[1][1].xml).toContain('<w:dateFormat w:val="dd/MM/yyyy HH:mm"/>');
    expect(r[1][1].xml).toContain('w:fullDate="2026-11-23T14:30:00Z"');
    expect(r[1][3].text).toBe("教育局");
    expect(r[2][1].text).toBe("23/11/2026 16:05");
    expect(r[2][1].xml).toContain('<w:dateFormat w:val="dd/MM/yyyy HH:mm"/>');
    expect(r[2][1].xml).toContain('w:fullDate="2026-11-23T16:05:00Z"');
    expect(r[2][3].text).toBe("0901 234 567");
    expect(r[3][1].text).toBe("送公文");
    // 保安確認留白
    expect(r[3][3].text).toBe("");
    expect(r[5][0].paras).toEqual(["王小明", "22/11/2026"]);
    expect(r[5][0].xml).toContain('<w:sz w:val="26"/>');
    expect(r[5][0].xml).toContain('<w:sz w:val="18"/>');

    // 表格上方的填表日期
    const xml = await documentXml(out);
    const top = xml.slice(0, xml.indexOf("<w:tbl>"));
    expect(text(top)).toContain("Date: 22/11/2026");
    expect(top).toContain('w:fullDate="2026-11-22T00:00:00Z"');

    // 作者資訊清掉
    const core = await (await JSZip.loadAsync(out)).file("docProps/core.xml")!.async("string");
    expect(core).not.toMatch(/<(?:\w+:)?creator[^>]*>[^<]+</);
    expect(core).not.toMatch(/<(?:\w+:)?lastModifiedBy[^>]*>[^<]+</);
  });

  it("重填不會累積舊值", async () => {
    const once = await fillGoOutForm(readFileSync(TEMPLATE), info);
    const twice = await fillGoOutForm(once, { ...info, destination: "醫院", outTime: "08:00", backTime: "09:00" });
    const r = await rows(twice);
    expect(r[1][3].text).toBe("醫院");
    expect(r[1][1].text).toBe("23/11/2026 08:00");
    expect(r[2][1].text).toBe("23/11/2026 09:00");
    expect((await inspectGoOutForm(twice)).prefill.destination).toBe("醫院");
  });

  it("不是外出單時給清楚的錯誤", async () => {
    await expect(inspectGoOutForm(readFileSync("public/leave/F-QL-013G-leave-form.docx"))).rejects.toThrow(GoOutFormError);
    await expect(inspectGoOutForm(new Uint8Array([1, 2, 3]))).rejects.toThrow(/無法開啟/);
  });

  it("返回時間檢查與檔名", () => {
    expect(goOutTimeProblem(info)).toBeNull();
    expect(goOutTimeProblem({ ...info, backTime: "14:30" })).toMatch(/晚於/);
    expect(goOutTimeProblem({ ...info, backDate: "2026-11-22" })).toMatch(/晚於/);
    // 上班時間 07:30–16:30
    expect(goOutTimeProblem({ ...info, outTime: "07:30", backTime: "16:30" })).toBeNull();
    expect(goOutTimeProblem({ ...info, outTime: "07:00" })).toMatch(/外出時間要在上班時間 07:30–16:30/);
    expect(goOutTimeProblem({ ...info, backTime: "17:00" })).toMatch(/返回時間要在上班時間/);
    expect(goOutFileName({ name: "王/小明", date: "2026-11-22" })).toBe("外出單-王小明-20261122.docx");
  });
});
