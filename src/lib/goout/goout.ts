import { child, descendants, kids, paragraphText, remove, serialize, setW, W, wAttr, type Document, type Element } from "../compose/xml";
import {
  cellText,
  collapseTrailingParagraph,
  compact,
  FormError,
  gridStart,
  loadDocument,
  nextCell,
  paragraphs,
  scrubAuthors,
  setCellLines,
  slashDate,
} from "../forms/docx";
import { WORK_END, WORK_START } from "../leave/leave";

/*
 * 外出單（F-QL-034 A PERMISSION FOR GOING OUT FORM 外出申請單 ĐƠN XIN RA NGOÀI）填寫。
 * 跟請假單一樣：用中文標籤找格子、把值填進旁邊的格子，版面保持原樣。
 * 不同的是填表日期、外出時間、預計返回時間是 Word 的「日期選擇器」控制項，
 * 要同時更新控制項記錄的日期（w:fullDate）和顯示的文字（依控制項的日期格式）。
 */

export class GoOutFormError extends FormError {}

export interface GoOutInfo {
  name: string;
  dept: string;
  position: string;
  /** 外出時間：yyyy-mm-dd、HH:MM */
  outDate: string;
  outTime: string;
  /** 預計返回時間 */
  backDate: string;
  backTime: string;
  destination: string;
  contact: string;
  reason: string;
  /** 申請人簽名欄的名字 */
  applicant: string;
  /** 填表日期 yyyy-mm-dd（上方日期欄，也押在申請人名字下方） */
  date: string;
}

interface Cells {
  name: Element;
  dept: Element;
  position: Element;
  outTime: Element;
  destination: Element;
  backTime: Element;
  contact: Element;
  reason: Element;
  signature: Element;
  /** 「填表日期 Date:」那一段（表格上方） */
  dateParagraph?: Element;
}

function findCells(doc: Document): Cells {
  const body = child(doc.documentElement, "body");
  if (!body) throw new GoOutFormError("這不是 Word 文件（找不到內文）");
  const tableText = (t: Element) => descendants(t, W, "p").map((p) => paragraphText(p, false)).join("\n");
  const tbl = descendants(body, W, "tbl").find((t) => /姓名/.test(tableText(t)) && /申請人/.test(tableText(t)));
  if (!tbl) throw new GoOutFormError("找不到外出單表格：版型裡要有「姓名」和「申請人」欄位");
  const tcs = descendants(tbl, W, "tc");
  const label = (re: RegExp, what: string) => {
    const tc = tcs.find((c) => re.test(cellText(c)));
    if (!tc) throw new GoOutFormError(`外出單表格裡找不到「${what}」欄位`);
    return tc;
  };
  const valueOf = (re: RegExp, what: string) => {
    const tc = nextCell(label(re, what));
    if (!tc) throw new GoOutFormError(`「${what}」旁邊沒有可以填寫的格子`);
    return tc;
  };

  const applicant = label(/申請人/, "申請人");
  let sigRow = applicant.parentNode!.nextSibling;
  while (sigRow && (sigRow as Element).localName !== "tr") sigRow = sigRow.nextSibling;
  const col = gridStart(applicant);
  const signature = sigRow ? kids(sigRow as Element).find((c) => c.localName === "tc" && gridStart(c) === col) : undefined;
  if (!signature) throw new GoOutFormError("「申請人」下方沒有簽名格");

  const dateParagraph = kids(body)
    .filter((k) => k.localName === "p")
    .find((p) => /填表|Date/i.test(paragraphText(p, false)));

  return {
    name: valueOf(/姓名/, "姓名"),
    dept: valueOf(/部門/, "部門"),
    position: valueOf(/職務/, "職務"),
    outTime: valueOf(/外出時間/, "外出時間"),
    destination: valueOf(/外出地點/, "外出地點"),
    backTime: valueOf(/返回時間/, "預計返回時間"),
    contact: valueOf(/聯繫方式|聯絡方式/, "聯繫方式"),
    // 版型寫的是「除外事由」（應為外出事由），兩種都認
    reason: valueOf(/事由/, "事由"),
    signature,
    dateParagraph,
  };
}

/** 日期選擇器控制項（w:sdt 裡有 w:date） */
const dateControl = (el: Element | undefined) =>
  el ? descendants(el, W, "sdt").find((sdt) => child(child(sdt, "sdtPr"), "date")) : undefined;

interface DateParts {
  y: number;
  m: number;
  d: number;
  h: number;
  min: number;
}

function parts(isoDate: string, hm = "00:00"): DateParts | null {
  const dm = /^(\d{4})-(\d{2})-(\d{2})$/.exec(isoDate);
  const tm = /^(\d{1,2}):(\d{2})$/.exec(hm || "00:00");
  if (!dm || !tm) return null;
  return { y: +dm[1], m: +dm[2], d: +dm[3], h: +tm[1], min: +tm[2] };
}

const pad = (n: number, w = 2) => String(n).padStart(w, "0");

/** 上午 / 下午的寫法依控制項的語言（vi-VN：SA / CH） */
function ampm(lid: string, pm: boolean, upper: boolean) {
  if (/^vi/i.test(lid)) return pm ? "CH" : "SA";
  if (/^zh/i.test(lid)) return pm ? "下午" : "上午";
  if (/^ja/i.test(lid)) return pm ? "午後" : "午前";
  const s = pm ? "pm" : "am";
  return upper ? s.toUpperCase() : s;
}

/**
 * 依 Word 日期格式（dd/MM/yyyy h:mm am/pm 這種）排出文字。
 * MM 是月、mm 是分；h 是 12 小時制、H 是 24 小時制；'文字' 原樣輸出。
 */
export function formatWordDate(format: string, lid: string, p: DateParts): string {
  const h12 = p.h % 12 || 12;
  return format.replace(/yyyy|yy|MM|M|dd|d|HH|H|hh|h|mm|m|ss|s|am\/pm|AM\/PM|'[^']*'/g, (t) => {
    switch (t) {
      case "yyyy": return String(p.y);
      case "yy": return pad(p.y % 100);
      case "MM": return pad(p.m);
      case "M": return String(p.m);
      case "dd": return pad(p.d);
      case "d": return String(p.d);
      case "HH": return pad(p.h);
      case "H": return String(p.h);
      case "hh": return pad(h12);
      case "h": return String(h12);
      case "mm": return pad(p.min);
      case "m": return String(p.min);
      case "ss": return "00";
      case "s": return "0";
      case "am/pm": return ampm(lid, p.h >= 12, false);
      case "AM/PM": return ampm(lid, p.h >= 12, true);
      default: return t.slice(1, -1);
    }
  });
}

/** 外出、返回時間的格式：日/月/年 + 24 小時制的時:分（不要秒，也不要 SA / CH） */
export const TIME_FORMAT = "dd/MM/yyyy HH:mm";

/**
 * 更新日期選擇器：記錄的日期與顯示的文字都要改，Word 打開才會一致。
 * 有給 format 時連控制項的格式一起換掉，之後在 Word 裡重選日期也會是同樣的格式。
 */
function setDateControl(sdt: Element, p: DateParts, format?: string) {
  const sdtPr = child(sdt, "sdtPr")!;
  const date = child(sdtPr, "date")!;
  setW(date, "fullDate", `${p.y}-${pad(p.m)}-${pad(p.d)}T${pad(p.h)}:${pad(p.min)}:00Z`);
  const dateFormat = child(date, "dateFormat");
  if (format && dateFormat) setW(dateFormat, "val", format);
  format ??= wAttr(dateFormat, "val") ?? "dd/MM/yyyy";
  const lid = wAttr(child(date, "lid"), "val") ?? "";
  // 原本顯示的是提示文字（「按一下輸入日期」）時，改成一般文字
  const placeholder = child(sdtPr, "showingPlcHdr");
  if (placeholder) remove(placeholder);
  const content = child(sdt, "sdtContent")!;
  const ts = descendants(content, W, "t");
  if (!ts.length) return;
  ts[0].textContent = formatWordDate(format, lid, p);
  for (const t of ts.slice(1)) remove(t.parentNode as Element);
  const rStyle = child(child(ts[0].parentNode as Element, "rPr"), "rStyle");
  if (rStyle && /placeholder/i.test(wAttr(rStyle, "val") ?? "")) remove(rStyle);
}

/** 格子裡有日期選擇器就更新它，沒有就直接寫文字（其他版本的外出單） */
function setDateCell(doc: Document, tc: Element, isoDate: string, hm: string) {
  const p = parts(isoDate, hm);
  const sdt = dateControl(tc);
  if (p && sdt) return setDateControl(sdt, p, TIME_FORMAT);
  const text = p ? formatWordDate(TIME_FORMAT, "", p) : "";
  // 版型的時間格裡第一段是空白、第二段才是日期；沒有控制項時寫在最後一段
  setCellLines(doc, tc, [text]);
}

export interface GoOutTemplateInfo {
  /** 版型裡已經填的內容（上傳填過的外出單時帶入表單） */
  prefill: Partial<Pick<GoOutInfo, "name" | "dept" | "position" | "destination" | "contact" | "reason" | "applicant">>;
}

export async function inspectGoOutForm(data: ArrayBuffer | Uint8Array): Promise<GoOutTemplateInfo> {
  const { doc } = await loadDocument(data, GoOutFormError);
  const c = findCells(doc);
  const v = (tc: Element) => cellText(tc).trim();
  return {
    prefill: {
      name: v(c.name),
      dept: v(c.dept),
      position: v(c.position),
      destination: v(c.destination),
      contact: v(c.contact),
      reason: v(c.reason),
      applicant: v(c.signature).split("\n")[0],
    },
  };
}

export async function fillGoOutForm(data: ArrayBuffer | Uint8Array, info: GoOutInfo): Promise<Uint8Array> {
  const { zip, doc } = await loadDocument(data, GoOutFormError);
  const c = findCells(doc);

  setCellLines(doc, c.name, [info.name.trim()]);
  setCellLines(doc, c.dept, [info.dept.trim()]);
  setCellLines(doc, c.position, [info.position.trim()]);
  setCellLines(doc, c.destination, [info.destination.trim()]);
  setCellLines(doc, c.contact, [info.contact.trim()]);
  const reasonLines = info.reason.trim().split(/\r?\n/);
  setCellLines(doc, c.reason, reasonLines);
  if (reasonLines.length > 1) for (const p of paragraphs(c.reason)) compact(doc, p);

  setDateCell(doc, c.outTime, info.outDate, info.outTime);
  setDateCell(doc, c.backTime, info.backDate, info.backTime);

  // 填表日期（表格上方）
  const fill = parts(info.date);
  const fillSdt = dateControl(c.dateParagraph);
  if (fill && fillSdt) setDateControl(fillSdt, fill);

  // 申請人：名字（13pt），下一行押填表日期 dd/mm/yyyy（9pt），跟請假單一樣
  setCellLines(doc, c.signature, [info.applicant.trim(), slashDate(info.date.trim())]);
  paragraphs(c.signature).forEach((p, i) => compact(doc, p, i === 1 ? 18 : 26));

  collapseTrailingParagraph(doc);
  zip.file("word/document.xml", serialize(doc));
  await scrubAuthors(zip);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

export function goOutFileName(info: Pick<GoOutInfo, "name" | "date">) {
  const safe = (s: string) => s.replace(/[\\/:*?"<>|]+/g, "").trim();
  return ["外出單", safe(info.name), info.date.replace(/\D/g, "")].filter(Boolean).join("-") + ".docx";
}

/** 上班時間 07:30–16:30（跟請假單一樣），外出與返回都要在這段時間內 */
export { WORK_END, WORK_START } from "../leave/leave";

/** 外出、返回時間要在上班時間內，返回要晚於外出 */
export function goOutTimeProblem(info: Pick<GoOutInfo, "outDate" | "outTime" | "backDate" | "backTime">): string | null {
  const inHours = (t: string) => !t || (t >= WORK_START && t <= WORK_END);
  if (!inHours(info.outTime)) return `外出時間要在上班時間 ${WORK_START}–${WORK_END} 內`;
  if (!inHours(info.backTime)) return `返回時間要在上班時間 ${WORK_START}–${WORK_END} 內`;
  if (!info.outDate || !info.backDate) return null;
  const out = `${info.outDate}T${info.outTime || "00:00"}`;
  const back = `${info.backDate}T${info.backTime || "00:00"}`;
  return back <= out ? "預計返回時間要晚於外出時間" : null;
}
