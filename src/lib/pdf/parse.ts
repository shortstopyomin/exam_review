import type { ParsedDocument, Paragraph } from "../types";

export class PdfParseError extends Error {}

/** pdf.js 裡用到的部分：瀏覽器用 pdfjs-dist，測試（Node）用 pdfjs-dist/legacy */
export interface PdfLib {
  getDocument(src: { data: Uint8Array; verbosity?: number }): { promise: Promise<PdfDoc>; destroy(): Promise<void> };
}
interface PdfDoc {
  numPages: number;
  getPage(n: number): Promise<PdfPage>;
}
interface PdfPage {
  view: number[];
  getTextContent(): Promise<{ items: unknown[] }>;
}
interface TextItem {
  str: string;
  transform: number[];
  width: number;
  height: number;
  hasEOL: boolean;
}

export function isPdf(data: ArrayBuffer | Uint8Array): boolean {
  const b = data instanceof Uint8Array ? data : new Uint8Array(data);
  // 「%PDF-」可能前面有幾個垃圾位元組
  const head = String.fromCharCode(...b.subarray(0, 1024));
  return head.includes("%PDF-");
}

async function loadBrowserPdfjs(): Promise<PdfLib> {
  const pdfjs = await import("pdfjs-dist");
  if (!pdfjs.GlobalWorkerOptions.workerPort) {
    pdfjs.GlobalWorkerOptions.workerPort = new Worker(new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url), {
      type: "module",
    });
  }
  return pdfjs as unknown as PdfLib;
}

/** 頁碼行：「1」「- 2 -」「第 3 頁」「3 / 5」「Page 2 of 5」 */
const PAGE_NUMBER_RE = /^[-–—\s]*(?:第\s*)?\d{1,3}(?:\s*(?:\/|of)\s*\d{1,3})?\s*頁?[-–—\s]*$|^page\s*\d{1,3}(?:\s*of\s*\d{1,3})?$/i;

/** 換行後的配分單位：「分」「分）」「points」 */
const SCORE_UNIT_START_RE = /^\s*(?:分|points?\b|pts?\b|marks?\b)/i;
/** 句尾標點：上一行以這些結尾，下一行就是新的一句 */
const SENTENCE_END_RE = /[。？?！!：:；;）)」』\]】.．…]\s*$/;

interface Line {
  y: number;
  height: number;
  xEnd: number;
  text: string;
}

/**
 * 把一頁的文字項目組成一行一行的文字。
 * pdf.js 的項目依內容順序排列，每行開頭常有一個空字串、行尾帶 hasEOL；
 * 同一行的項目 y 座標相同（上下標只差一點點），換到另一個 y 就是新的一行。
 */
function pageLines(items: unknown[]): Line[] {
  const lines: Line[] = [];
  let line: Line | undefined;
  const close = () => {
    if (line && line.text.trim()) lines.push(line);
    line = undefined;
  };

  for (const raw of items) {
    const it = raw as Partial<TextItem>;
    if (typeof it.str !== "string" || !it.transform) continue;
    const x = it.transform[4];
    const y = it.transform[5];
    const h = Math.abs(it.height ?? 0) || Math.abs(it.transform[3]) || 10;

    if (it.str) {
      if (line && Math.abs(y - line.y) > Math.max(2, 0.5 * Math.min(h, line.height))) close();
      if (!line) {
        line = { y, height: h, xEnd: x + (it.width ?? 0), text: it.str };
      } else {
        // 兩個項目之間有明顯空隙但沒有空白字元：補一個空白
        const gap = x - line.xEnd;
        if (gap > 0.25 * h && !/\s$/.test(line.text) && !/^\s/.test(it.str)) line.text += " ";
        line.text += it.str;
        line.xEnd = Math.max(line.xEnd, x + (it.width ?? 0));
        line.height = Math.max(line.height, h);
      }
    }
    if (it.hasEOL) close();
  }
  close();
  return lines;
}

/**
 * 解析 PDF 試卷。PDF 沒有段落，每一行文字當成一個段落；
 * 題目換行的部分會在分析結構時併入同一題。
 * 表格、圖片、方程式的資訊 PDF 裡拿不到，一律當成沒有。
 */
export async function parsePdf(
  data: ArrayBuffer | Uint8Array,
  fileName = "試卷.pdf",
  lib?: PdfLib,
): Promise<ParsedDocument> {
  const pdfjs = lib ?? (await loadBrowserPdfjs());
  // pdf.js 會把傳入的 buffer 轉移給 worker，先複製一份
  const bytes = new Uint8Array(data instanceof Uint8Array ? data : new Uint8Array(data)).slice();

  const task = pdfjs.getDocument({ data: bytes, verbosity: 0 });
  let doc: PdfDoc;
  try {
    doc = await task.promise;
  } catch (e) {
    const name = (e as Error).name;
    if (name === "PasswordException") throw new PdfParseError("這個 PDF 有密碼保護，請先移除密碼再上傳。");
    throw new PdfParseError(`無法讀取 PDF：${(e as Error).message}`);
  }

  const paragraphs: Paragraph[] = [];
  try {
    for (let n = 1; n <= doc.numPages; n++) {
      const page = await doc.getPage(n);
      const [, y0, , y1] = page.view;
      const margin = (y1 - y0) * 0.08;
      let prev: Line | undefined;
      for (const line of pageLines((await page.getTextContent()).items)) {
        const text = line.text.replace(/\s+$/, "");
        // 頁首、頁尾的頁碼
        if ((line.y < y0 + margin || line.y > y1 - margin) && PAGE_NUMBER_RE.test(text.trim())) continue;

        const last = paragraphs.at(-1);
        // 緊接在上一行正下方
        const below = prev !== undefined && prev.y > line.y && prev.y - line.y <= 2 * Math.max(prev.height, line.height);
        // 配分被換行切開：「每題 5 分，共 20」⏎「分 Word Problems…」→ 接回同一行
        if (last && below && /\d\s*$/.test(last.text) && SCORE_UNIT_START_RE.test(text)) {
          last.text += text;
        } else {
          paragraphs.push({
            index: paragraphs.length,
            text,
            inTable: false,
            hasMath: false,
            imageCount: 0,
            // 上一行沒有句尾標點：這行可能是換行接下去的文字（「…number is⏎34. Her classmate…」）
            ...(last && below && !SENTENCE_END_RE.test(last.text) ? { wrapped: true } : {}),
          });
        }
        prev = line;
      }
    }
  } finally {
    await task.destroy();
  }

  if (!paragraphs.length) {
    throw new PdfParseError("這個 PDF 抽不出文字，可能是掃描的圖片檔。請上傳 Word 檔，或先用 OCR 轉成可選取文字的 PDF。");
  }
  return { fileName, paragraphs };
}
