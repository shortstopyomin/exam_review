import JSZip from "jszip";
import { child, kids, paragraphText, parseXml, remove, setW, W, wAttr, wEl, type Document, type Element } from "../compose/xml";

/*
 * 填寫 Word 表單（請假單、外出單…）共用的工具：
 * 找格子、把文字填進格子並沿用版型的字型、調整行距與列高讓表格維持一頁。
 */

export class FormError extends Error {}

export const cellText = (tc: Element) => kids(tc).filter((k) => k.localName === "p").map((p) => paragraphText(p, false)).join("\n");
export const paragraphs = (tc: Element) => kids(tc).filter((k) => k.namespaceURI === W && k.localName === "p");
export const nextCell = (tc: Element) => kids(tc.parentNode as Element).filter((k) => k.localName === "tc")[indexOfCell(tc) + 1];

export function indexOfCell(tc: Element) {
  return kids(tc.parentNode as Element)
    .filter((k) => k.localName === "tc")
    .indexOf(tc);
}

/** 格子的起始欄位（考慮 gridSpan），用來對齊上下兩列 */
export function gridStart(tc: Element) {
  let col = 0;
  for (const c of kids(tc.parentNode as Element).filter((k) => k.localName === "tc")) {
    if (c === tc) return col;
    col += Number(wAttr(child(child(c, "tcPr"), "gridSpan"), "val") ?? 1);
  }
  return -1;
}

export async function loadDocument(data: ArrayBuffer | Uint8Array, Err: new (message: string) => Error = FormError) {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    throw new Err("無法開啟這個檔案，請確認是 Word .docx 檔");
  }
  const file = zip.file("word/document.xml");
  if (!file) throw new Err("這不是 Word .docx 檔（找不到 word/document.xml）");
  return { zip, doc: parseXml(await file.async("string")) };
}

/** 新 run：格式沿用段落標記的格式（版型空白格子裡只有段落標記會帶字型） */
export function makeRun(doc: Document, p: Element, text: string, like?: Element): Element {
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
export function setCellLines(doc: Document, tc: Element, lines: string[]) {
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
export function setAfterLabel(doc: Document, tc: Element, value: string): Element | undefined {
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
export function fitRun(doc: Document, run: Element, text: string, avail: number, base = 12) {
  const size = Math.max(7, Math.min(base, Math.floor((avail / textEms(text)) * 2) / 2));
  if (size >= base) return;
  const rPr = ensure(doc, run, "rPr", ["rPr", "t"]);
  setProp(doc, rPr, "sz", { val: String(Math.round(size * 2)) }, RPR_ORDER);
  setProp(doc, rPr, "szCs", { val: String(Math.round(size * 2)) }, RPR_ORDER);
}

// Word 規定的子元素順序；插入新元素時要放在正確的位置，否則 Word 會說檔案損毀
export const PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];
export const RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"];

/** 設定（或取代）一個屬性元素，依 Word 的順序放好 */
export function setProp(doc: Document, parent: Element, name: string, attrs: Record<string, string>, order: string[]) {
  const old = child(parent, name);
  if (old) remove(old);
  const el = wEl(doc, name, attrs);
  const rank = order.indexOf(name);
  const after = kids(parent).find((k) => order.indexOf(k.localName ?? "") > rank);
  parent.insertBefore(el, after ?? null);
}

export function ensure(doc: Document, parent: Element, name: string, order: string[]): Element {
  const found = child(parent, name);
  if (found) return found;
  setProp(doc, parent, name, {}, order);
  return child(parent, name)!;
}

/**
 * 段落改成緊湊行距：版型的文件格線（每行 18pt）會把每一行撐到格線高度，
 * 請假單、外出單這種 A5 表格幾乎剛好佔滿一頁，多一點高度就會掉到第二頁。
 * halfPoints 有給時順便改字級（Word 以半點為單位，18 = 9pt）。
 */
export function compact(doc: Document, p: Element, halfPoints?: number) {
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

/** 把一列的最小高度加到 target（twips），回傳加了多少 */
export function growRow(tc: Element, target: number): number {
  const h = child(child(tc.parentNode as Element, "trPr"), "trHeight");
  const cur = Number(wAttr(h, "val") ?? 0);
  if (!h || !cur || cur >= target) return 0;
  setW(h, "val", String(target));
  return target - cur;
}

/** 簽名列的最小高度（twips）：其他列長高多少，就從這裡扣回來，讓整張表維持一頁 */
export function shrinkRow(doc: Document, tc: Element, twips: number) {
  if (twips <= 0) return;
  const tr = tc.parentNode as Element;
  const trPr = child(tr, "trPr");
  const h = child(trPr, "trHeight");
  const cur = Number(wAttr(h, "val") ?? 0);
  if (!h || !cur) return;
  // 至少留 40pt（約 1.4 公分）給主管們簽名
  setW(h, "val", String(Math.max(800, cur - twips)));
}

/**
 * Word 規定表格後面一定要有一個段落。表單的表格剛好佔滿一頁時，
 * 後面那個空段落（一行格線 18pt）會被擠到第二頁，變成一張空白頁。
 * 把它縮成 1pt 高、不對齊格線，就能留在第一頁。
 */
export function collapseTrailingParagraph(doc: Document) {
  const body = child(doc.documentElement, "body");
  if (!body) return;
  const items = kids(body).filter((k) => k.localName !== "sectPr");
  const last = items[items.length - 1];
  const before = items[items.length - 2];
  if (!last || last.localName !== "p" || before?.localName !== "tbl") return;
  if (paragraphText(last).trim() || kids(last).some((k) => k.localName !== "pPr")) return;
  const pPr = child(last, "pPr") ?? (last.insertBefore(wEl(doc, "pPr"), last.firstChild) as Element);
  setProp(doc, pPr, "snapToGrid", { val: "0" }, PPR_ORDER);
  setProp(doc, pPr, "spacing", { before: "0", after: "0", line: "20", lineRule: "exact" }, PPR_ORDER);
  const rPr = ensure(doc, pPr, "rPr", PPR_ORDER);
  setProp(doc, rPr, "sz", { val: "2" }, RPR_ORDER);
  setProp(doc, rPr, "szCs", { val: "2" }, RPR_ORDER);
}

/**
 * 表單上的日期一律寫成 日/月/年（跟外出單版型的日期欄一樣）：
 * 2026-11-23 → 23/11/2026；只給月-日時 11-25 → 25/11
 */
export function slashDate(iso: string) {
  return iso.split("-").reverse().join("/");
}

/**
 * 清掉 Word 檔裡記錄的作者（建立者、最後修改者）。版型是從同事電腦存出來的，
 * 這些名字不該跟著版型公開，也不該出現在別人填好的表單裡。
 */
export async function scrubAuthors(zip: JSZip) {
  const file = zip.file("docProps/core.xml");
  if (!file) return;
  const xml = await file.async("string");
  const cleaned = xml.replace(/<((?:\w+:)?(?:creator|lastModifiedBy))(\s[^>]*)?>[^<]*<\/\1>/g, "<$1$2></$1>");
  if (cleaned !== xml) zip.file("docProps/core.xml", cleaned);
}
