import JSZip from "jszip";
import { child, descendants, kids, paragraphText, parseXml, remove, serialize, setW, W, wAttr, wEl, type Document, type Element } from "../compose/xml";

/*
 * 請假單（F-QL-013G APPLICATION FOR LEAVE 請假單 ĐƠN XIN PHÉP）填寫。
 * 版型是一個表格：每個欄位用中文標籤找格子，再把值填進旁邊（或下一列）的格子，
 * 版面、字型、頁首頁尾、流程說明全部保持原樣。
 */

export class LeaveFormError extends Error {}

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
  /** 申請日期 yyyy/mm/dd，押在申請人名字下方 */
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
const cellText = (tc: Element) => kids(tc).filter((k) => k.localName === "p").map((p) => paragraphText(p, false)).join("\n");
const paragraphs = (tc: Element) => kids(tc).filter((k) => k.namespaceURI === W && k.localName === "p");
const nextCell = (tc: Element) => kids(tc.parentNode as Element).filter((k) => k.localName === "tc")[indexOfCell(tc) + 1];

function indexOfCell(tc: Element) {
  return kids(tc.parentNode as Element)
    .filter((k) => k.localName === "tc")
    .indexOf(tc);
}

/** 格子的起始欄位（考慮 gridSpan），用來對齊上下兩列 */
function gridStart(tc: Element) {
  let col = 0;
  for (const c of kids(tc.parentNode as Element).filter((k) => k.localName === "tc")) {
    if (c === tc) return col;
    col += Number(wAttr(child(child(c, "tcPr"), "gridSpan"), "val") ?? 1);
  }
  return -1;
}

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
  const doc = await loadDocument(data);
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

async function loadDocument(data: ArrayBuffer | Uint8Array) {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    throw new LeaveFormError("無法開啟這個檔案，請確認是 Word .docx 檔");
  }
  const file = zip.file("word/document.xml");
  if (!file) throw new LeaveFormError("這不是 Word .docx 檔（找不到 word/document.xml）");
  return { zip, doc: parseXml(await file.async("string")) };
}

/** 新 run：格式沿用段落標記的格式（版型空白格子裡只有段落標記會帶字型） */
function makeRun(doc: Document, p: Element, text: string, like?: Element): Element {
  const r = wEl(doc, "r");
  const rPr = child(like, "rPr") ?? child(child(p, "pPr"), "rPr");
  if (rPr) r.appendChild(rPr.cloneNode(true));
  const t = wEl(doc, "t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  return r;
}

/** 把格子的內容換成一行或多行文字，第一段的段落格式沿用到其他行 */
function setCellLines(doc: Document, tc: Element, lines: string[]) {
  const ps = paragraphs(tc);
  const first = ps[0] ?? tc.appendChild(wEl(doc, "p"));
  const like = kids(first).find((k) => k.localName === "r" && child(k, "t"));
  ps.slice(1).forEach(remove);
  for (const k of kids(first)) if (k.localName !== "pPr") remove(k);
  lines.forEach((line, i) => {
    const p = i === 0 ? first : (tc.appendChild(first.cloneNode(false)) as Element);
    if (i > 0) {
      const pPr = child(first, "pPr");
      if (pPr) p.appendChild(pPr.cloneNode(true));
    }
    if (line) p.appendChild(makeRun(doc, first, line, like));
  });
}

/** 標籤格：保留「…：」，把冒號後面的舊值換成新值 */
function setAfterLabel(doc: Document, tc: Element, value: string): Element | undefined {
  const ps = paragraphs(tc);
  const p = ps.find((x) => /[:：]/.test(paragraphText(x, false))) ?? ps[ps.length - 1];
  const runs = kids(p).filter((k) => k.localName === "r");
  // 標籤到第一個冒號為止，之後的 run 都是舊值（舊值本身可能有「17:00」這種冒號）
  const k = runs.findIndex((r) => /[:：]/.test(child(r, "t")?.textContent ?? ""));
  if (k >= 0) {
    const t = child(runs[k], "t")!;
    const s = t.textContent ?? "";
    t.textContent = s.slice(0, s.search(/[:：]/) + 1);
    runs.slice(k + 1).forEach(remove);
  }
  if (value) return p.appendChild(makeRun(doc, p, ` ${value}`, runs.find((r) => child(r, "t")))) as Element;
}

/**
 * 估計文字寬度（em）：中文、全形 1；Times New Roman 的數字 0.5、「/ : . ,」約 0.28、空白 0.25、
 * 其他英文字母抓 0.55（大寫較寬，估多一點比較保險）
 */
export function textEms(text: string): number {
  let ems = 0;
  for (const ch of text) {
    if (/[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF\u3000-\u303F]/.test(ch)) ems += 1;
    else if (ch === " ") ems += 0.25;
    else if (/[/:.,;'|]/.test(ch)) ems += 0.28;
    else if (/[0-9]/.test(ch)) ems += 0.5;
    else ems += 0.55;
  }
  return ems;
}

/** 字太長放不下一行時縮小字級（最小 7pt） */
function fitRun(doc: Document, run: Element, text: string, avail: number, base = 12) {
  const size = Math.max(7, Math.min(base, Math.floor((avail / textEms(text)) * 2) / 2));
  if (size >= base) return;
  const rPr = ensure(doc, run, "rPr", ["rPr", "t"]);
  setProp(doc, rPr, "sz", { val: String(Math.round(size * 2)) }, RPR_ORDER);
  setProp(doc, rPr, "szCs", { val: String(Math.round(size * 2)) }, RPR_ORDER);
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

// Word 規定的子元素順序；插入新元素時要放在正確的位置，否則 Word 會說檔案損毀
const PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];
const RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"];

/** 設定（或取代）一個屬性元素，依 Word 的順序放好 */
function setProp(doc: Document, parent: Element, name: string, attrs: Record<string, string>, order: string[]) {
  const old = child(parent, name);
  if (old) remove(old);
  const el = wEl(doc, name, attrs);
  const rank = order.indexOf(name);
  const after = kids(parent).find((k) => order.indexOf(k.localName ?? "") > rank);
  parent.insertBefore(el, after ?? null);
}

function ensure(doc: Document, parent: Element, name: string, order: string[]): Element {
  const found = child(parent, name);
  if (found) return found;
  setProp(doc, parent, name, {}, order);
  return child(parent, name)!;
}

/**
 * 段落改成緊湊行距：版型的文件格線（每行 18pt）會把每一行撐到格線高度，
 * 這張 A5 請假單的表格幾乎剛好佔滿一頁，多一點高度就會掉到第二頁。
 * halfPoints 有給時順便改字級（Word 以半點為單位，18 = 9pt）。
 */
function compact(doc: Document, p: Element, halfPoints?: number) {
  const pPr = child(p, "pPr") ?? (p.insertBefore(wEl(doc, "pPr"), p.firstChild) as Element);
  setProp(doc, pPr, "snapToGrid", { val: "0" }, PPR_ORDER);
  setProp(doc, pPr, "spacing", { before: "0", after: "0", line: "240", lineRule: "auto" }, PPR_ORDER);
  if (!halfPoints) return;
  for (const r of kids(p).filter((k) => k.localName === "r")) {
    const rPr = ensure(doc, r, "rPr", ["rPr", "t"]);
    setProp(doc, rPr, "sz", { val: String(halfPoints) }, RPR_ORDER);
    setProp(doc, rPr, "szCs", { val: String(halfPoints) }, RPR_ORDER);
  }
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

/** 把一列的最小高度加到 target（twips），回傳加了多少 */
function growRow(tc: Element, target: number): number {
  const h = child(child(tc.parentNode as Element, "trPr"), "trHeight");
  const cur = Number(wAttr(h, "val") ?? 0);
  if (!h || !cur || cur >= target) return 0;
  setW(h, "val", String(target));
  return target - cur;
}

/** 簽名列的最小高度（twips）：其他列長高多少，就從這裡扣回來，讓整張表維持一頁 */
function shrinkRow(doc: Document, tc: Element, twips: number) {
  if (twips <= 0) return;
  const tr = tc.parentNode as Element;
  const trPr = child(tr, "trPr");
  const h = child(trPr, "trHeight");
  const cur = Number(wAttr(h, "val") ?? 0);
  if (!h || !cur) return;
  // 至少留 40pt（約 1.4 公分）給主管們簽名
  setW(h, "val", String(Math.max(800, cur - twips)));
}

export async function fillLeaveForm(data: ArrayBuffer | Uint8Array, info: LeaveInfo): Promise<Uint8Array> {
  const { zip, doc } = await loadDocument(data);
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

  // 簽名格：名字，下一行押日期（9pt）；緊湊行距、垂直置中、水平置中
  setCellLines(doc, c.signature, [info.applicant.trim(), info.date.trim()]);
  paragraphs(c.signature).forEach((p, i) => compact(doc, p, i === 1 ? 18 : undefined));
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

  zip.file("word/document.xml", serialize(doc));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

/** 2026-11-23 → 2026/11/23 */
export function slashDate(iso: string) {
  return iso.replace(/-/g, "/");
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
  return ["請假單", safe(info.name), info.date.replace(/\D/g, "")].filter(Boolean).join("-") + ".docx";
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
