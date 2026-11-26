import { DOMParser, XMLSerializer, type Document, type Element, type Node } from "@xmldom/xmldom";

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
export const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
export const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
export const V = "urn:schemas-microsoft-com:vml";
export const O = "urn:schemas-microsoft-com:office:office";
export const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

export type { Document, Element, Node };

export function parseXml(xml: string): Document {
  // 有些工具存檔時會在 XML 前面加 UTF-8 BOM（Word 可以開），xmldom 會因此報錯
  return new DOMParser().parseFromString(xml.replace(/^﻿/, ""), "text/xml");
}

export function serialize(doc: Document): string {
  return new XMLSerializer().serializeToString(doc);
}

export function kids(node: Node): Element[] {
  const out: Element[] = [];
  for (let c = node.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) out.push(c as Element);
  return out;
}

export function is(node: Node | null | undefined, ns: string, name: string): boolean {
  return !!node && node.nodeType === 1 && (node as Element).localName === name && (node as Element).namespaceURI === ns;
}

export function child(node: Node | null | undefined, name: string, ns = W): Element | undefined {
  return node ? kids(node).find((c) => is(c, ns, name)) : undefined;
}

/** 所有子孫元素（含自己），文件順序 */
export function descendants(node: Element, ns: string, name: string): Element[] {
  const out: Element[] = [];
  const walk = (n: Element) => {
    if (is(n, ns, name)) out.push(n);
    for (const c of kids(n)) walk(c);
  };
  walk(node);
  return out;
}

export function allElements(node: Element): Element[] {
  const out: Element[] = [];
  const walk = (n: Element) => {
    out.push(n);
    for (const c of kids(n)) walk(c);
  };
  walk(node);
  return out;
}

export function wAttr(node: Element | undefined, name: string): string | null {
  if (!node) return null;
  return node.getAttributeNS(W, name) || node.getAttribute(`w:${name}`) || null;
}

export function setW(node: Element, name: string, value: string) {
  node.setAttributeNS(W, `w:${name}`, value);
}

export function wEl(doc: Document, name: string, attrs: Record<string, string> = {}): Element {
  const el = doc.createElementNS(W, `w:${name}`);
  for (const [k, v] of Object.entries(attrs)) setW(el, k, v);
  return el;
}

export function remove(node: Node) {
  node.parentNode?.removeChild(node);
}

/** 段落裡的文字節點（w:t 與 m:t），和 docx/parse.ts 讀文字的順序一致 */
function textNodes(p: Element, includeMath: boolean): Element[] {
  const out: Element[] = [];
  const walk = (n: Element) => {
    for (const c of kids(n)) {
      if (c.namespaceURI === W) {
        if (c.localName === "t") {
          out.push(c);
          continue;
        }
        if (c.localName === "delText" || c.localName === "instrText" || c.localName === "pPr" || c.localName === "rPr")
          continue;
      } else if (c.namespaceURI === M && c.localName === "t") {
        if (includeMath) out.push(c);
        continue;
      }
      walk(c);
    }
  };
  walk(p);
  return out;
}

export function paragraphText(p: Element, includeMath = true): string {
  return textNodes(p, includeMath)
    .map((t) => t.textContent ?? "")
    .join("");
}

function setText(t: Element, text: string) {
  while (t.firstChild) t.removeChild(t.firstChild);
  t.appendChild(t.ownerDocument!.createTextNode(text));
  if (t.namespaceURI === W) t.setAttribute("xml:space", "preserve");
}

/**
 * 在段落中取代文字，可以跨越多個 run（例如「Mid」「term」分在兩個 run）。
 * 取代後的文字放在比對開頭所在的節點，沿用那個 run 的格式。
 * 只處理 w:t（不改算式）。
 */
export function replaceInParagraph(p: Element, re: RegExp, replacement: (m: RegExpExecArray) => string): boolean {
  const nodes = textNodes(p, false);
  if (!nodes.length) return false;
  const texts = nodes.map((n) => n.textContent ?? "");
  const full = texts.join("");
  const flags = re.flags.includes("g") ? re.flags : re.flags + "g";
  const g = new RegExp(re.source, flags);
  const matches: { start: number; end: number; text: string }[] = [];
  for (let m = g.exec(full); m; m = g.exec(full)) {
    if (m[0].length === 0) {
      g.lastIndex++;
      continue;
    }
    matches.push({ start: m.index, end: m.index + m[0].length, text: replacement(m) });
  }
  if (!matches.length) return false;

  // 由後往前改，前面的位置才不會跑掉
  const starts: number[] = [];
  let pos = 0;
  for (const t of texts) {
    starts.push(pos);
    pos += t.length;
  }
  for (const mt of matches.reverse()) {
    for (let i = nodes.length - 1; i >= 0; i--) {
      const s = starts[i];
      const e = s + texts[i].length;
      if (e <= mt.start || s >= mt.end) continue;
      const from = Math.max(mt.start, s) - s;
      const to = Math.min(mt.end, e) - s;
      const insert = mt.start >= s && mt.start < e ? mt.text : "";
      texts[i] = texts[i].slice(0, from) + insert + texts[i].slice(to);
    }
  }
  nodes.forEach((n, i) => {
    if ((n.textContent ?? "") !== texts[i]) setText(n, texts[i]);
  });
  return true;
}

/** 把 run 在某個 w:t 的第 offset 個字切成兩個 run（格式相同），回傳後半段的 run */
function splitRun(r: Element, t: Element, offset: number): Element {
  const parts = kids(r).filter((c) => !is(c, W, "rPr"));
  const idx = parts.indexOf(t);
  const r2 = r.cloneNode(true) as Element;
  const parts2 = kids(r2).filter((c) => !is(c, W, "rPr"));
  const text = t.textContent ?? "";
  setText(t, text.slice(0, offset));
  setText(parts2[idx], text.slice(offset));
  // 前半段留 t 之前的內容，後半段留 t 之後的內容
  parts.slice(idx + 1).forEach(remove);
  parts2.slice(0, idx).forEach(remove);
  r.parentNode!.insertBefore(r2, r.nextSibling);
  return r2;
}

/**
 * 段落中 [start, end) 這段文字所在的 run（位置和 replaceInParagraph 一樣，不含算式）。
 * run 只有一部分在範圍內時先切開，回傳的 run 整個都在範圍內，可以直接改格式。
 */
export function runsInRange(p: Element, start: number, end: number): Element[] {
  const out: Element[] = [];
  let pos = 0;
  for (let t of textNodes(p, false)) {
    const len = (t.textContent ?? "").length;
    const s = pos;
    pos += len;
    const r = t.parentNode as Element;
    if (!len || pos <= start || s >= end || !is(r, W, "r")) continue;
    let run = r;
    if (start > s) {
      // 切開後，這個 w:t 的後半段是新 run 的第一個內容
      run = splitRun(run, t, start - s);
      t = kids(run).find((c) => is(c, W, "t"))!;
    }
    if (end < pos) splitRun(run, t, end - Math.max(start, s));
    if (!out.includes(run)) out.push(run);
  }
  return out;
}

/** 刪掉段落最前面 n 個字元（含算式文字，與 paragraphText 相同的計算方式） */
export function removeLeadingChars(p: Element, n: number) {
  for (const t of textNodes(p, true)) {
    if (n <= 0) break;
    const s = t.textContent ?? "";
    const cut = Math.min(n, s.length);
    setText(t, s.slice(cut));
    n -= cut;
  }
}

// ─── 預覽文字：把算式轉成看得懂的線性文字 ─────────────────────────

const SUP: Record<string, string> = {
  "0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹",
  "+": "⁺", "-": "⁻", "－": "⁻", "＋": "⁺", "(": "⁽", ")": "⁾", "n": "ⁿ",
};
const SUB: Record<string, string> = {
  "0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄", "5": "₅", "6": "₆", "7": "₇", "8": "₈", "9": "₉",
};

function toScript(s: string, map: Record<string, string>, fallback: string): string {
  const t = s.normalize("NFKC");
  return [...t].every((c) => map[c] || /\s/.test(c)) ? [...t].map((c) => map[c] ?? "").join("") : `${fallback}(${s})`;
}

const wrap = (s: string) => (s.length > 1 && !/^\(.*\)$/.test(s) ? `(${s})` : s);

/** Word 的 EQ 功能變數：\R(,x) 根號、\F(a,b) 分數，可以巢狀 */
export function formatEqField(instr: string): string {
  const src = instr.replace(/^\s*EQ\s*/i, "");
  let i = 0;
  // 讀到 stop 裡的字元（括號外）為止
  const seq = (stop: string): string => {
    let out = "";
    while (i < src.length && !stop.includes(src[i])) {
      if (src[i] === "\\") {
        const m = /^\\([A-Za-z]+)((?:\\[a-z]{2}-?\d*)*)\s*/.exec(src.slice(i));
        if (!m) {
          out += src[i + 1] ?? "";
          i += 2;
          continue;
        }
        i += m[0].length;
        const cmd = m[1].toUpperCase();
        if (src[i] !== "(") continue;
        i++;
        const args: string[] = [seq(",)")];
        while (src[i] === ",") {
          i++;
          args.push(seq(",)"));
        }
        i++; // ")"
        const a = args.map((x) => x.trim());
        if (cmd === "R") out += a.length > 1 ? `${a[0] ? toScript(a[0], SUP, "^") : ""}√${wrap(a[1])}` : `√${wrap(a[0])}`;
        else if (cmd === "F") out += `${wrap(a[0])}/${wrap(a[1] ?? "")}`;
        else if (cmd === "S" && a.length === 1) out += a[0];
        else out += a.join(cmd === "A" ? " " : "");
      } else if (src[i] === "(") {
        i++;
        out += `(${seq(")")})`;
        i++;
      } else out += src[i++];
    }
    return out;
  };
  return seq("").trim();
}

function mathText(el: Element): string {
  const part = (name: string) => {
    const c = child(el, name, M);
    return c ? mathText(c) : "";
  };
  if (el.namespaceURI === M) {
    switch (el.localName) {
      case "t":
        return el.textContent ?? "";
      case "sSup":
        return part("e") + toScript(part("sup"), SUP, "^");
      case "sSub":
        return part("e") + toScript(part("sub"), SUB, "_");
      case "sSubSup":
        return part("e") + toScript(part("sub"), SUB, "_") + toScript(part("sup"), SUP, "^");
      case "rad": {
        const deg = part("deg");
        return `${deg ? toScript(deg, SUP, "^") : ""}√${wrap(part("e"))}`;
      }
      case "f":
        return `${wrap(part("num"))}/${wrap(part("den"))}`;
      case "rPr":
      case "ctrlPr":
      case "fPr":
      case "radPr":
      case "sSupPr":
      case "sSubPr":
      case "dPr":
        return "";
      case "d": {
        const pr = child(el, "dPr", M);
        const beg = pr ? (child(pr, "begChr", M)?.getAttributeNS(M, "val") ?? "(") : "(";
        const end = pr ? (child(pr, "endChr", M)?.getAttributeNS(M, "val") ?? ")") : ")";
        return beg + kids(el).filter((c) => is(c, M, "e")).map(mathText).join(",") + end;
      }
    }
  }
  return kids(el).map(mathText).join("");
}

/** 一個段落的預覽文字（含 EQ 功能變數與 OMML 算式） */
export function previewParagraph(p: Element): string {
  let out = "";
  // 功能變數可以巢狀（EQ 裡面再放 EQ）：內層的代碼併回外層一起解析
  const stack: { instr: string; showResult: boolean }[] = [];
  const isEq = (s: string) => /^\s*EQ\b/i.test(s);

  const run = (r: Element) => {
    const va = wAttr(child(child(r, "rPr"), "vertAlign"), "val");
    for (const c of kids(r)) {
      if (c.namespaceURI !== W) continue;
      const top = stack[stack.length - 1];
      switch (c.localName) {
        case "fldChar": {
          const type = wAttr(c, "fldCharType");
          if (type === "begin") stack.push({ instr: "", showResult: false });
          else if (type === "separate" && top && !isEq(top.instr)) top.showResult = true;
          else if (type === "end" && top) {
            stack.pop();
            if (isEq(top.instr)) {
              const parent = stack[stack.length - 1];
              if (parent && !parent.showResult) parent.instr += " " + top.instr.replace(/^\s*EQ\s*/i, "");
              else out += formatEqField(top.instr);
            }
          }
          break;
        }
        case "instrText":
          if (top && !top.showResult) {
            const t = c.textContent ?? "";
            top.instr += va === "superscript" && t.trim() ? toScript(t, SUP, "^") : t;
          }
          break;
        case "t":
          if (!top || top.showResult) {
            const t = c.textContent ?? "";
            out += va === "superscript" ? toScript(t, SUP, "^") : va === "subscript" ? toScript(t, SUB, "_") : t;
          }
          break;
        case "tab":
          out += " ";
          break;
        case "br":
        case "cr":
          out += "\n";
          break;
      }
    }
  };

  const walk = (n: Element) => {
    for (const c of kids(n)) {
      if (is(c, W, "r")) run(c);
      else if (is(c, W, "fldSimple")) {
        const instr = c.getAttributeNS(W, "instr") ?? "";
        if (/^\s*EQ\b/i.test(instr)) out += formatEqField(instr);
        else walk(c);
      } else if (c.namespaceURI === M && (c.localName === "oMath" || c.localName === "oMathPara")) out += mathText(c);
      else if (is(c, W, "pPr") || is(c, W, "del")) continue;
      else walk(c);
    }
  };
  walk(p);
  return out;
}
