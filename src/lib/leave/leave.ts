import { child, descendants, kids, paragraphText, remove, serialize, W, wAttr, wEl, type Document, type Element } from "../compose/xml";
import {
  FormError,
  cellText,
  paragraphs,
  nextCell,
  gridStart,
  loadDocument,
  makeRun,
  setCellLines,
  setAfterLabel,
  textEms,
  fitRun,
  compact,
  growRow,
  shrinkRow,
  collapseTrailingParagraph,
  scrubAuthors,
  slashDate,
} from "../forms/docx";

export { slashDate, textEms } from "../forms/docx";

/*
 * 請假單（F-QL-013G APPLICATION FOR LEAVE 請假單 ĐƠN XIN PHÉP）填寫。
 * 版型是一個表格：每個欄位用中文標籤找格子，再把值填進旁邊（或下一列）的格子，
 * 版面、字型、頁首頁尾、流程說明全部保持原樣。
 */

export class LeaveFormError extends FormError {}

/** Wingdings 的打勾框（þ = U+00FE；Word 有時存成私用區 U+F0FE） */
const BOX_CHECKED = "þ";
/** Wingdings 的方框：空框 ¨（U+00A8 / U+F0A8），已勾選的框也要認得，這樣填好的請假單可以再拿來改 */
const BOX_ANY = /^[\u00A8\uF0A8\u00FE\uF0FE\uF0FD\uF078]$/;
/** 我們在「其他 Other」後面補的說明，重填時要先拿掉 */
const OTHER_NOTE = /^：/;

/** 假別：中文是穩定的 key，越南文 / 英文只是顯示用 */
export const LEAVE_TYPES = [
  { zh: "事假", vi: "Việc riêng", en: "Personal leave" },
  { zh: "病假", vi: "Phép bệnh", en: "Sick leave" },
  { zh: "生理假", vi: "Phép sinh lý", en: "Menstrual leave" },
  { zh: "補假", vi: "Nghỉ bù", en: "Compensatory leave" },
  { zh: "年假", vi: "Phép năm", en: "Annual leave" },
  { zh: "喪假", vi: "Phép tang", en: "Compassionate leave" },
  { zh: "婚假", vi: "Hôn lễ", en: "Marriage leave" },
  { zh: "產假", vi: "Phép thai sản", en: "Maternity leave" },
  { zh: "公假", vi: "Phép công", en: "Paid leave" },
  { zh: "其他", vi: "Khác", en: "Other" },
] as const;

export type LeaveType = (typeof LEAVE_TYPES)[number]["zh"];

export interface LeaveInfo {
  name: string;
  dept: string;
  position: string;
  /** 勾選的假別（單選） */
  type: LeaveType | null;
  /** 勾「其他」時補充的說明 */
  otherText: string;
  /** 請假時間，例如 2026/11/23 08:00 ~ 2026/11/24 17:00 */
  period: string;
  /** 總計，例如 1.5 天 */
  total: string;
  reason: string;
  /** 申請人簽名欄的名字 */
  applicant: string;
  /** 申請日期 dd/mm/yyyy，押在申請人名字下方 */
  date: string;
}

export interface LeaveTemplateInfo {
  /** 版型裡找得到的假別（依版型順序） */
  types: LeaveType[];
  /** 版型裡原本就有的值（例如事由欄預先打好的字），拿來帶入表單 */
  prefill: Partial<Pick<LeaveInfo, "name" | "dept" | "position" | "period" | "total" | "reason" | "applicant">> & {
    /** 版型裡已經勾選的假別（上傳填過的請假單時） */
    type?: LeaveType;
  };
}

interface Cells {
  name: Element;
  dept: Element;
  position: Element;
  periodLabel: Element;
  totalLabel: Element;
  reason: Element;
  signature: Element;
  /** 每個空框 run 的 w:t，以及它後面標籤對應的假別 */
  boxes: { t: Element; type: LeaveType; labelRun: Element; checked: boolean }[];
}

const isSymbolFont = (r: Element) => /wingdings|symbol/i.test(wAttr(child(child(r, "rPr"), "rFonts"), "ascii") ?? "");
function findCells(doc: Document): Cells {
  const body = child(doc.documentElement, "body");
  if (!body) throw new LeaveFormError("這不是 Word 文件（找不到內文）");
  const tableText = (t: Element) => descendants(t, W, "p").map((p) => paragraphText(p, false)).join("\n");
  const tbl = descendants(body, W, "tbl").find((t) => /姓名/.test(tableText(t)) && /申請人/.test(tableText(t)));
  if (!tbl) throw new LeaveFormError("找不到請假單表格：版型裡要有「姓名」和「申請人」欄位");
  const tcs = descendants(tbl, W, "tc");
  const label = (re: RegExp, what: string) => {
    const tc = tcs.find((c) => re.test(cellText(c)));
    if (!tc) throw new LeaveFormError(`請假單表格裡找不到「${what}」欄位`);
    return tc;
  };
  const valueOf = (re: RegExp, what: string) => {
    const tc = nextCell(label(re, what));
    if (!tc) throw new LeaveFormError(`「${what}」旁邊沒有可以填寫的格子`);
    return tc;
  };

  const applicant = label(/申請人/, "申請人");
  const row = applicant.parentNode as Element;
  let sigRow = row.nextSibling;
  while (sigRow && (sigRow as Element).localName !== "tr") sigRow = sigRow.nextSibling;
  const col = gridStart(applicant);
  const signature = sigRow ? kids(sigRow as Element).find((c) => c.localName === "tc" && gridStart(c) === col) : undefined;
  if (!signature) throw new LeaveFormError("「申請人」下方沒有簽名格");

  // 假別：Wingdings 空框，後面接著的文字就是假別名稱
  const boxes: Cells["boxes"] = [];
  for (const p of descendants(tbl, W, "p")) {
    const runs = kids(p).filter((k) => k.localName === "r");
    runs.forEach((r, i) => {
      const t = child(r, "t");
      if (!t || !BOX_ANY.test(t.textContent ?? "") || !isSymbolFont(r)) return;
      let text = "";
      let labelRun = r;
      for (let j = i + 1; j < runs.length; j++) {
        const tt = child(runs[j], "t")?.textContent ?? "";
        if (BOX_ANY.test(tt) && isSymbolFont(runs[j])) break;
        if (OTHER_NOTE.test(tt)) {
          remove(runs[j]);
          continue;
        }
        text += tt;
        if (tt.trim()) labelRun = runs[j];
      }
      const type = LEAVE_TYPES.find((lt) => text.includes(lt.zh));
      if (type) boxes.push({ t, type: type.zh, labelRun, checked: /^[\u00FE\uF0FE\uF0FD\uF078]$/.test(t.textContent ?? "") });
    });
  }

  return {
    name: valueOf(/姓名/, "姓名"),
    dept: valueOf(/單位/, "單位"),
    position: valueOf(/職務/, "職務"),
    periodLabel: label(/請假時間與日數/, "請假時間與日數"),
    totalLabel: label(/^\s*縂|^\s*總|total/i, "總計 total"),
    reason: valueOf(/請假事由/, "請假事由"),
    signature,
    boxes,
  };
}

/** 標籤格（「…time and days：」）冒號後面的文字 */
const afterColon = (tc: Element) => cellText(tc).split(/[:：]/).slice(1).join("：").trim();

export async function inspectLeaveForm(data: ArrayBuffer | Uint8Array): Promise<LeaveTemplateInfo> {
  const doc = await loadDocument(data, LeaveFormError);
  const c = findCells(doc.doc);
  const types = [...new Set(c.boxes.map((b) => b.type))];
  if (!types.length) throw new LeaveFormError("請假單裡找不到假別的勾選框（□ 事假、□ 病假…）");
  const v = (tc: Element) => cellText(tc).trim();
  return {
    types: LEAVE_TYPES.map((t) => t.zh).filter((t) => types.includes(t)),
    prefill: {
      name: v(c.name),
      dept: v(c.dept),
      position: v(c.position),
      period: afterColon(c.periodLabel),
      total: afterColon(c.totalLabel),
      reason: v(c.reason),
      applicant: v(c.signature).split("\n")[0],
      type: c.boxes.find((b) => b.checked)?.type,
    },
  };
}

/**
 * 總計格：第一行「縂total： 2 天 3 小時」，第二行請假時間「2026/11/23 07:30 ~ 11/25 10:30」。
 * 重填時先拿掉上次加的第二行。
 */
function setTotal(doc: Document, tc: Element, total: string, period: string) {
  const width = Number(wAttr(child(child(tc, "tcPr"), "tcW"), "w") ?? 0) / 20; // twips → pt
  const run = setAfterLabel(doc, tc, total);
  const ps = paragraphs(tc);
  const label = ps.find((x) => /[:：]/.test(paragraphText(x, false))) ?? ps[ps.length - 1];
  ps.slice(ps.indexOf(label) + 1).forEach(remove);
  if (run && width) fitRun(doc, run, ` ${total}`, width - 12 - textEms(cellText(tc).split(/[:：]/)[0] + "：") * 12);
  if (!period) return;
  const p = label.cloneNode(false) as Element;
  const pPr = child(label, "pPr");
  if (pPr) p.appendChild(pPr.cloneNode(true));
  const r = p.appendChild(makeRun(doc, label, period, kids(label).find((k) => k.localName === "r" && child(k, "t")))) as Element;
  tc.insertBefore(p, label.nextSibling);
  compact(doc, p);
  if (width) fitRun(doc, r, period, width - 12);
}

/**
 * 請假時間：標籤「…time and days：」已經幾乎佔滿整格，時間放在下一段（10pt、緊湊行距），
 * 這樣這一列只會多一小行。重填時先拿掉上次加的那一段。
 */
function setPeriod(doc: Document, tc: Element, value: string) {
  setAfterLabel(doc, tc, "");
  const ps = paragraphs(tc);
  const label = ps.find((x) => /[:：]/.test(paragraphText(x, false))) ?? ps[ps.length - 1];
  ps.slice(ps.indexOf(label) + 1).forEach(remove);
  if (!value) return;
  const p = label.cloneNode(false) as Element;
  const pPr = child(label, "pPr");
  if (pPr) p.appendChild(pPr.cloneNode(true));
  p.appendChild(makeRun(doc, label, value, kids(label).find((k) => k.localName === "r" && child(k, "t"))));
  tc.insertBefore(p, label.nextSibling);
  compact(doc, p, 20);
}

export async function fillLeaveForm(data: ArrayBuffer | Uint8Array, info: LeaveInfo): Promise<Uint8Array> {
  const { zip, doc } = await loadDocument(data, LeaveFormError);
  const c = findCells(doc);
  // 簽名格原本是空的 = 第一次填這張版型；重填時高度已經調過，不要再扣
  const firstFill = !cellText(c.signature).trim();

  setCellLines(doc, c.name, [info.name.trim()]);
  setCellLines(doc, c.dept, [info.dept.trim()]);
  setCellLines(doc, c.position, [info.position.trim()]);
  // 請假時間與天數都寫在「縂total：」那一格（兩行）；「請假時間與日數」格只留標籤
  setPeriod(doc, c.periodLabel, "");
  const period = info.period.trim();
  setTotal(doc, c.totalLabel, info.total.trim(), period);
  const reasonLines = info.reason.trim().split(/\r?\n/);
  setCellLines(doc, c.reason, reasonLines);
  if (reasonLines.length > 1) for (const p of paragraphs(c.reason)) compact(doc, p);

  for (const b of c.boxes) {
    const on = info.type === b.type;
    b.t.textContent = on ? BOX_CHECKED : "¨";
    if (b.type === "其他" && on && info.otherText.trim()) {
      const t = child(b.labelRun, "t")!;
      t.textContent = (t.textContent ?? "").trimEnd();
      b.labelRun.parentNode!.insertBefore(
        makeRun(doc, b.labelRun.parentNode as Element, `：${info.otherText.trim()}`, b.labelRun),
        b.labelRun.nextSibling,
      );
    }
  }

  // 簽名格：名字（13pt，比版型的 12pt 大一點），下一行押日期（9pt）；緊湊行距、垂直置中、水平置中
  setCellLines(doc, c.signature, [info.applicant.trim(), info.date.trim()]);
  paragraphs(c.signature).forEach((p, i) => compact(doc, p, i === 1 ? 18 : 26));
  if (firstFill) {
    // 時間那一列加高到 36pt（720 twips）放兩行；事由每多一行約 14.4pt（290 twips）。都從簽名列扣回來
    const grown = period ? growRow(c.totalLabel, 720) : 0;
    shrinkRow(doc, c.signature, grown + (reasonLines.length - 1) * 290);
  }
  for (const p of paragraphs(c.signature)) {
    let pPr = child(p, "pPr");
    if (!pPr) pPr = p.insertBefore(wEl(doc, "pPr"), p.firstChild) as Element;
    if (!child(pPr, "jc")) {
      const jc = wEl(doc, "jc", { val: "center" });
      const after = kids(pPr).find((k) => ["rPr", "sectPr", "pPrChange"].includes(k.localName ?? ""));
      pPr.insertBefore(jc, after ?? null);
    }
  }
  const tcPr = child(c.signature, "tcPr");
  if (tcPr && !child(tcPr, "vAlign")) {
    const after = kids(tcPr).find((k) => ["hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"].includes(k.localName ?? ""));
    tcPr.insertBefore(wEl(doc, "vAlign", { val: "center" }), after ?? null);
  }

  collapseTrailingParagraph(doc);
  zip.file("word/document.xml", serialize(doc));
  await scrubAuthors(zip);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

/**
 * 請假時間的文字：同一天只寫一次日期（2026/11/23 08:00 ~ 17:00）；
 * 同一年結束日不寫年份（2026/11/23 08:00 ~ 11/24 17:00），總計格的字可以大一點。
 */
export function formatPeriod(sd: string, st: string, ed: string, et: string) {
  if (!sd) return "";
  const start = [slashDate(sd), st].filter(Boolean).join(" ");
  if (!ed) return start;
  if (ed === sd) return et ? `${start} ~ ${et}` : start;
  const end = ed.slice(0, 4) === sd.slice(0, 4) ? slashDate(ed.slice(5)) : slashDate(ed);
  return `${start} ~ ${[end, et].filter(Boolean).join(" ")}`;
}

export function leaveFileName(info: Pick<LeaveInfo, "name" | "date">) {
  const safe = (s: string) => s.replace(/[\\/:*?"<>|]+/g, "").trim();
  // 表單上的日期是 dd/mm/yyyy；檔名用 yyyymmdd，檔案依日期排序才會對
  const dmy = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(info.date.trim());
  const ymd = dmy ? dmy[3] + dmy[2].padStart(2, "0") + dmy[1].padStart(2, "0") : info.date.replace(/\D/g, "");
  return ["請假單", safe(info.name), ymd].filter(Boolean).join("-") + ".docx";
}

/** 上班時間 07:30–16:30，一天 = 9 小時 */
export const WORK_START = "07:30";
export const WORK_END = "16:30";
const toMin = (hm: string) => {
  const [h, m] = hm.split(":").map(Number);
  return h * 60 + (m || 0);
};
const DAY_MIN = toMin(WORK_END) - toMin(WORK_START);

export interface Duration {
  days: number;
  hours: number;
}

/**
 * 依起訖時間計算請假時數：只算上班時間（07:30–16:30）。週六、週日也照同樣的上班時間計算。
 * 整天上滿算 1 天，不滿一天的時數加總後以 1 小時為單位無條件進位（最少 1 小時），
 * 滿 9 小時進位成 1 天。只是預設值，使用者可以改。
 */
export function estimateDuration(startDate: string, startTime: string, endDate: string, endTime: string): Duration | null {
  if (!startDate || !endDate) return null;
  const s = new Date(`${startDate}T00:00`);
  const e = new Date(`${endDate}T00:00`);
  if (isNaN(+s) || isNaN(+e) || e < s) return null;
  const ws = toMin(WORK_START);
  const we = toMin(WORK_END);
  let days = 0;
  let partial = 0;
  for (const d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
    const first = +d === +s;
    const last = +d === +e;
    const from = Math.max(ws, first && startTime ? toMin(startTime) : ws);
    const to = Math.min(we, last && endTime ? toMin(endTime) : we);
    const mins = Math.max(0, to - from);
    if (mins >= DAY_MIN) days++;
    else partial += mins;
  }
  if (days === 0 && partial === 0) return null;
  let hours = Math.ceil(partial / 60);
  const perDay = DAY_MIN / 60;
  days += Math.floor(hours / perDay);
  hours %= perDay;
  return { days, hours };
}

/** { days: 2, hours: 3 } → 2 天 3 小時 */
export function formatDuration(d: Duration): string {
  return [d.days ? `${d.days} 天` : "", d.hours ? `${d.hours} 小時` : ""].filter(Boolean).join(" ");
}
