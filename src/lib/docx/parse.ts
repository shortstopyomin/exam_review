import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";
import type { Paragraph, ParsedDocument } from "../types";
import { NumberingTracker, type LevelDef } from "./numbering";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

// 只用到 DOM 的一小部分；自訂最小介面，避免 xmldom 與 lib.dom 型別衝突。
interface XNode {
  nodeType: number;
  localName?: string | null;
  namespaceURI?: string | null;
  childNodes: ArrayLike<XNode>;
  textContent?: string | null;
  getAttributeNS?(ns: string, name: string): string | null;
  getAttribute?(name: string): string | null;
}

function elements(node: XNode): XNode[] {
  const out: XNode[] = [];
  for (let i = 0; i < node.childNodes.length; i++) {
    const c = node.childNodes[i];
    if (c.nodeType === 1) out.push(c);
  }
  return out;
}

function is(node: XNode, ns: string, name: string): boolean {
  return node.localName === name && node.namespaceURI === ns;
}

function child(node: XNode | undefined, name: string, ns = W): XNode | undefined {
  if (!node) return undefined;
  return elements(node).find((c) => is(c, ns, name));
}

function wAttr(node: XNode | undefined, name: string): string | null {
  if (!node) return null;
  return node.getAttributeNS?.(W, name) || node.getAttribute?.(`w:${name}`) || null;
}

function parseXml(xml: string): XNode {
  return new DOMParser().parseFromString(xml, "text/xml") as unknown as XNode;
}

function root(doc: XNode): XNode | undefined {
  return elements(doc)[0];
}

async function readXml(zip: JSZip, path: string): Promise<XNode | undefined> {
  const file = zip.file(path);
  if (!file) return undefined;
  return root(parseXml(await file.async("string")));
}

function loadNumbering(numberingRoot: XNode | undefined): NumberingTracker {
  const abstracts = new Map<string, Map<number, LevelDef>>();
  const nums = new Map<string, { abstractId: string; overrides: Map<number, number> }>();
  if (numberingRoot) {
    for (const el of elements(numberingRoot)) {
      if (is(el, W, "abstractNum")) {
        const levels = new Map<number, LevelDef>();
        for (const lvl of elements(el).filter((c) => is(c, W, "lvl"))) {
          levels.set(Number(wAttr(lvl, "ilvl") ?? 0), {
            start: Number(wAttr(child(lvl, "start"), "val") ?? 1),
            fmt: wAttr(child(lvl, "numFmt"), "val") ?? "decimal",
            text: wAttr(child(lvl, "lvlText"), "val") ?? "",
          });
        }
        abstracts.set(wAttr(el, "abstractNumId") ?? "", levels);
      } else if (is(el, W, "num")) {
        const overrides = new Map<number, number>();
        for (const ov of elements(el).filter((c) => is(c, W, "lvlOverride"))) {
          const start = wAttr(child(ov, "startOverride"), "val");
          if (start !== null) overrides.set(Number(wAttr(ov, "ilvl") ?? 0), Number(start));
        }
        nums.set(wAttr(el, "numId") ?? "", {
          abstractId: wAttr(child(el, "abstractNumId"), "val") ?? "",
          overrides,
        });
      }
    }
  }
  return new NumberingTracker(abstracts, nums);
}

interface NumPr {
  numId: string;
  ilvl: number;
}

/** 段落樣式也可能帶自動編號（例如「清單段落」樣式） */
function loadStyleNumbering(stylesRoot: XNode | undefined): Map<string, NumPr> {
  const direct = new Map<string, Partial<NumPr>>();
  const basedOn = new Map<string, string>();
  if (stylesRoot) {
    for (const st of elements(stylesRoot).filter((c) => is(c, W, "style"))) {
      const id = wAttr(st, "styleId");
      if (!id) continue;
      const based = wAttr(child(st, "basedOn"), "val");
      if (based) basedOn.set(id, based);
      const numPr = child(child(st, "pPr"), "numPr");
      if (numPr) {
        direct.set(id, {
          numId: wAttr(child(numPr, "numId"), "val") ?? undefined,
          ilvl: Number(wAttr(child(numPr, "ilvl"), "val") ?? 0),
        });
      }
    }
  }
  const resolved = new Map<string, NumPr>();
  for (const id of new Set([...direct.keys(), ...basedOn.keys()])) {
    let cur: string | undefined = id;
    for (let depth = 0; cur && depth < 10; depth++) {
      const d = direct.get(cur);
      if (d?.numId) {
        resolved.set(id, { numId: d.numId, ilvl: d.ilvl ?? 0 });
        break;
      }
      cur = basedOn.get(cur);
    }
  }
  return resolved;
}

interface Ctx {
  numbering: NumberingTracker;
  styleNumbering: Map<string, NumPr>;
  paragraphs: Paragraph[];
}

function readParagraph(p: XNode, inTable: boolean, ctx: Ctx) {
  let text = "";
  let hasMath = false;
  let imageCount = 0;

  const walk = (node: XNode) => {
    for (const el of elements(node)) {
      if (el.namespaceURI === W) {
        switch (el.localName) {
          case "t":
            text += el.textContent ?? "";
            continue;
          case "tab":
            text += "\t";
            continue;
          case "br":
          case "cr":
            text += "\n";
            continue;
          case "drawing":
          case "pict":
            imageCount++;
            continue;
          case "pPr":
          case "rPr":
          case "delText": // 追蹤修訂中被刪除的文字
          case "instrText": // 功能變數代碼
            continue;
        }
      } else if (el.namespaceURI === M) {
        hasMath = true;
        if (el.localName === "t") {
          text += el.textContent ?? "";
          continue;
        }
      }
      walk(el);
    }
  };
  walk(p);

  const pPr = child(p, "pPr");
  const numPrEl = child(pPr, "numPr");
  let numPr: NumPr | undefined;
  if (numPrEl) {
    const numId = wAttr(child(numPrEl, "numId"), "val");
    if (numId) numPr = { numId, ilvl: Number(wAttr(child(numPrEl, "ilvl"), "val") ?? 0) };
  } else {
    const styleId = wAttr(child(pPr, "pStyle"), "val");
    if (styleId) numPr = ctx.styleNumbering.get(styleId);
  }

  let numberingLabel: string | undefined;
  // numId "0" 代表明確取消編號
  if (numPr && numPr.numId !== "0") {
    numberingLabel = ctx.numbering.next(numPr.numId, numPr.ilvl);
  }

  ctx.paragraphs.push({
    index: ctx.paragraphs.length,
    text: numberingLabel ? `${numberingLabel} ${text}` : text,
    numberingLabel,
    inTable,
    hasMath,
    imageCount,
  });
}

function walkBlocks(node: XNode, inTable: boolean, ctx: Ctx) {
  for (const el of elements(node)) {
    if (is(el, W, "p")) readParagraph(el, inTable, ctx);
    else if (is(el, W, "tbl")) walkBlocks(el, true, ctx);
    else if (is(el, W, "tr") || is(el, W, "tc") || is(el, W, "sdt") || is(el, W, "sdtContent"))
      walkBlocks(el, inTable, ctx);
  }
}

export class DocxParseError extends Error {}

export async function parseDocx(data: ArrayBuffer | Uint8Array, fileName = "試卷.docx"): Promise<ParsedDocument> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    throw new DocxParseError("無法讀取檔案：這不是有效的 .docx 檔。若是舊版 .doc，請先在 Word 中「另存新檔」為 .docx。");
  }
  const docRoot = await readXml(zip, "word/document.xml");
  const body = child(docRoot, "body");
  if (!body) throw new DocxParseError("檔案中找不到 Word 文件內容（word/document.xml）。");

  const ctx: Ctx = {
    numbering: loadNumbering(await readXml(zip, "word/numbering.xml")),
    styleNumbering: loadStyleNumbering(await readXml(zip, "word/styles.xml")),
    paragraphs: [],
  };
  walkBlocks(body, false, ctx);
  return { fileName, paragraphs: ctx.paragraphs };
}
