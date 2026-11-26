import JSZip from "jszip";
import { parseDocx } from "../docx/parse";
import { analyzeStructure, parseQuestionStart } from "../structure";
import type { ExamStructure, ParsedDocument, SectionKind } from "../types";
import {
  CT,
  M,
  O,
  R,
  REL,
  V,
  W,
  W14,
  WP,
  allElements,
  child,
  descendants,
  is,
  kids,
  paragraphText,
  parseXml,
  previewParagraph,
  remove,
  removeLeadingChars,
  replaceInParagraph,
  serialize,
  setW,
  wAttr,
  wEl,
  type Document,
  type Element,
} from "./xml";

export type ExamType = "midterm" | "final";

export class ComposeError extends Error {}

// ─── 共用：把 body 的段落編號對應到最上層元素 ─────────────────────────

/**
 * docx/parse.ts 會把表格裡的段落也算進段落編號。
 * 這裡算出每個段落編號屬於 body 底下的第幾個元素（段落或整個表格）。
 */
function mapParagraphsToBlocks(body: Element): { blocks: Element[]; blockOf: number[] } {
  const blocks = kids(body);
  const blockOf: number[] = [];
  const count = (el: Element, top: number) => {
    for (const c of kids(el)) {
      if (is(c, W, "p")) blockOf.push(top);
      else if (is(c, W, "tbl") || is(c, W, "tr") || is(c, W, "tc") || is(c, W, "sdt") || is(c, W, "sdtContent"))
        count(c, top);
    }
  };
  blocks.forEach((b, i) => {
    if (is(b, W, "p")) blockOf.push(i);
    else if (is(b, W, "tbl") || is(b, W, "sdt")) count(b, i);
  });
  return { blocks, blockOf };
}

/** 空白段落：沒有文字、算式、圖片（題目之間的作答空間） */
function isBlankParagraph(el: Element): boolean {
  if (!is(el, W, "p")) return false;
  if (paragraphText(el).trim()) return false;
  return !allElements(el).some(
    (e) =>
      (e.namespaceURI === W && (e.localName === "drawing" || e.localName === "pict" || e.localName === "object")) ||
      (e.namespaceURI === M && e.localName === "oMath") ||
      (e.namespaceURI === W && e.localName === "sectPr"),
  );
}

/** 題目前面手打的題號：「1.」「( )1.」「（ B ）12．」 */
function leadingNumberLength(text: string): number {
  const q = parseQuestionStart(text);
  if (!q) return 0;
  const rest = q.rest.replace(/^[\s　]+/, "");
  return text.length - rest.length;
}

/** 刪掉段落開頭手打的題號，以及題號後面的 tab */
/**
 * 題目開頭要拿掉的部分：手打的題號，加上作答括號（空的或填了答案的），例如
 * 「( )1.」「1.（　）」「（ B ）」。題號用 Word 自動編號、括號手打在後面時，只剩括號。
 * 版型的編號會重新產生「(　　)1.」，新考卷也不該留答案。
 */
const ANSWER_BRACKET_RE = /^[\s\u3000]*[（(][\s\u3000]*[A-Za-zＡ-Ｅａ-ｅ○ＯＸX╳✕✗]?[\s\u3000]*[)）][\s\u3000]*/;
const EMPTY_BRACKET_RE = /^[\s\u3000]*[（(][\s\u3000]*[)）][\s\u3000]*/;
export function questionPrefixLength(text: string): number {
  const n = leadingNumberLength(text);
  // 前面沒有題號時只拿掉空括號：「(A) 21 (B) 27」是選項，不是填了答案的括號
  const m = (n ? ANSWER_BRACKET_RE : EMPTY_BRACKET_RE).exec(text.slice(n));
  // 括號後面要接題目文字，避免把「(A)」選項開頭的段落整個吃掉
  return m && text.length > n + m[0].length ? n + m[0].length : n;
}

function stripTypedNumber(p: Element) {
  const n = questionPrefixLength(paragraphText(p));
  if (!n) return;
  removeLeadingChars(p, n);
  for (const r of kids(p)) {
    if (!is(r, W, "r")) continue;
    if (paragraphText(r)) break;
    for (const t of kids(r)) if (is(t, W, "tab")) remove(t);
  }
}

// ─── 版型 ────────────────────────────────────────────────

export interface TemplateSection {
  title: string;
  kind: SectionKind;
  /** 版型標題寫的「每題 X 分」，當作預設配分 */
  perQuestionScore?: number;
}

/** 考卷標題裡可以改的資訊 */
export interface ExamInfo {
  examType: ExamType;
  /** 學年度開始的年份：2026 代表 2026-2027 */
  startYear?: number;
  semester?: number;
  grade?: number;
}

export interface TemplateInfo extends Partial<ExamInfo> {
  sections: TemplateSection[];
}

export async function inspectTemplate(data: ArrayBuffer | Uint8Array, fileName: string): Promise<TemplateInfo> {
  const doc = await parseDocx(data, fileName);
  const st = analyzeStructure(doc.paragraphs);
  const sections = st.sections.filter((s) => s.paragraphIndex >= 0);
  if (!sections.length)
    throw new ComposeError("版型裡找不到大題標題（例如「Part 1選擇題 / Multiple choice」或「一、選擇題」），無法判斷題目要放在哪裡。");
  const before = doc.paragraphs
    .filter((p) => p.index < sections[0].paragraphIndex)
    .map((p) => p.text)
    .join("\n");
  return {
    sections: sections.map((s) => ({ title: s.title, kind: s.kind, perQuestionScore: s.perQuestionScore })),
    ...detectExamInfo(before),
  };
}

// 「Final」只在後面接 Exam/Test… 時才算考試名稱，避免改到「final answer」
const MIDTERM_RE = /mid[-\s]?term/gi;
const FINAL_RE = /(?<![A-Za-z])final(?=[\s_-]*(?:exam|test|assessment|examination|$|[^A-Za-z]))/gi;
const FINAL_TEXT_RE = /(?<![A-Za-z])final(?=\s*(?:exam|test|assessment|examination)s?(?![A-Za-z]))/gi;

export function detectExamType(text: string): ExamType | undefined {
  if (/期中/.test(text) || new RegExp(MIDTERM_RE.source, "i").test(text)) return "midterm";
  if (/期末/.test(text) || new RegExp(FINAL_TEXT_RE.source, "i").test(text)) return "final";
  return undefined;
}

/*
 * 標題欄位。只比對數字本身（用前後文判斷），取代時才不會動到旁邊的文字格式，
 * 例如版型裡紅字的「8」後面接黑字的「年級」。
 */
const YEAR_SEP = String.raw`\s*[-–—~～/_]\s*`;
const TEXT_FIELDS = {
  /** 「2026-2027 學年度」「Year 2026-2027」的第一個年份 */
  startYear: [new RegExp(String.raw`(?<!\d)\d{4}(?=${YEAR_SEP}\d{4}(?!\d))`, "g")],
  endYear: [new RegExp(String.raw`(?<=(?<!\d)\d{4}${YEAR_SEP})\d{4}(?!\d)`, "g")],
  /** 「第 1學期」「Semester 1」 */
  semester: [/(?<=第\s*)\d(?=\s*學期)/g, /(?<=semester\s*)\d(?!\d)/gi],
  /** 「8年級」「Grade 8」 */
  grade: [/(?<!\d)\d{1,2}(?=\s*年級)/g, /(?<=grade\s*)\d{1,2}(?!\d)/gi],
};
/** 檔名：「2026_2027_S1_Midterm_G8_Math」 */
const NAME_FIELDS = {
  startYear: TEXT_FIELDS.startYear,
  endYear: TEXT_FIELDS.endYear,
  semester: [/(?<=(?<![A-Za-z0-9])S)\d(?!\d)/g, ...TEXT_FIELDS.semester],
  grade: [/(?<=(?<![A-Za-z0-9])G)\d{1,2}(?!\d)/g, ...TEXT_FIELDS.grade],
};

function firstNumber(text: string, res: RegExp[]): number | undefined {
  for (const re of res) {
    const m = new RegExp(re.source, re.flags.replace("g", "")).exec(text);
    if (m) return Number(m[0]);
  }
  return undefined;
}

export function detectExamInfo(text: string): Partial<ExamInfo> {
  return {
    examType: detectExamType(text),
    startYear: firstNumber(text, TEXT_FIELDS.startYear),
    semester: firstNumber(text, TEXT_FIELDS.semester),
    grade: firstNumber(text, TEXT_FIELDS.grade),
  };
}

type Replacer = (re: RegExp, to: (m: string) => string) => void;

function applyInfo(info: ExamInfo, fields: typeof TEXT_FIELDS, replace: Replacer, examTypeText: boolean) {
  if (info.examType === "final") {
    replace(/期中/g, () => "期末");
    replace(MIDTERM_RE, (m) => caseLike(m, "Final"));
  } else {
    replace(/期末/g, () => "期中");
    replace(examTypeText ? FINAL_TEXT_RE : FINAL_RE, (m) => caseLike(m, "Midterm"));
  }
  const set = (res: RegExp[], value: number | undefined) => {
    if (value === undefined || !Number.isFinite(value)) return;
    for (const re of res) replace(re, () => String(value));
  };
  // 先換第二個年份：它用前一個年份判斷位置，第一個先換掉也還是四位數，順序其實無妨
  set(fields.endYear, info.startYear !== undefined ? info.startYear + 1 : undefined);
  set(fields.startYear, info.startYear);
  set(fields.semester, info.semester);
  set(fields.grade, info.grade);
}

/** 依考試資訊改檔名：年份、S1、G8、Midterm/Final */
export function outputFileName(templateName: string, info: ExamInfo): string {
  const base = templateName.replace(/\.docx$/i, "");
  let name = base;
  applyInfo(info, NAME_FIELDS, (re, to) => (name = name.replace(re, (m) => to(m))), false);
  return `${name === base ? `${base}-組卷` : name}.docx`;
}

/** 標題段落：期中/期末、學年度、學期、年級 */
function setExamInfo(p: Element, info: ExamInfo) {
  applyInfo(info, TEXT_FIELDS, (re, to) => replaceInParagraph(p, re, (m) => to(m[0])), true);
}

const caseLike = (sample: string, word: string) =>
  sample === sample.toUpperCase() ? word.toUpperCase() : sample[0] === sample[0].toLowerCase() ? word.toLowerCase() : word;

// ─── 來源試卷 ──────────────────────────────────────────────

export interface SourceQuestion {
  id: string;
  sourceIndex: number;
  /** 原試卷的大題標題 */
  sectionTitle: string;
  kind: SectionKind;
  /** 原題號 */
  number: number;
  /** 預覽文字（算式只有線性文字） */
  text: string;
  hasMath: boolean;
  imageCount: number;
  /** 比對重複題用 */
  fingerprint: string;
  /** 與哪一題重複（較早出現的那題） */
  duplicateOf?: string;
  /** docx：題目在 body 中的元素範圍；pdf：題目文字（每段一行） */
  blocks?: Element[];
  lines?: string[];
}

export interface SourceExam {
  /** 上傳順序編號（移除檔案後不重新編號，題目 id 才不會變） */
  index: number;
  name: string;
  format: "docx" | "pdf";
  questions: SourceQuestion[];
  /** docx 才有：原檔，用來複製圖片與編號定義 */
  zip?: JSZip;
  numbering?: Document;
  styles?: Document;
  rels?: Map<string, { type: string; target: string; external: boolean }>;
}

function fnv(bytes: Uint8Array): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < bytes.length; i++) {
    h ^= bytes[i];
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16);
}

function normalizeForCompare(text: string): string {
  return text
    .normalize("NFKC")
    .toLowerCase()
    // 算式裡的 + − = 不能拿掉，否則「x+3」和「x−3」會被當成同一題
    .replace(/[\s\u3000]+/g, "")
    .replace(/[−－]/g, "-");
}

/** 大題標題看不出題型時，有 (A)(B) 選項就當選擇題 */
function guessKind(kind: SectionKind, body: string): SectionKind {
  if (kind !== "other") return kind;
  return /[（(]\s*A\s*[)）][\s\S]*[（(]\s*B\s*[)）]/.test(body) ? "choice" : "short";
}

async function readRels(zip: JSZip, path: string) {
  const map = new Map<string, { type: string; target: string; external: boolean }>();
  const file = zip.file(path);
  if (!file) return map;
  const doc = parseXml(await file.async("string"));
  for (const r of kids(doc.documentElement!)) {
    map.set(r.getAttribute("Id") ?? "", {
      type: r.getAttribute("Type") ?? "",
      target: r.getAttribute("Target") ?? "",
      external: r.getAttribute("TargetMode") === "External",
    });
  }
  return map;
}

function relAttrs(el: Element): { el: Element; name: string; value: string }[] {
  const out: { el: Element; name: string; value: string }[] = [];
  for (const e of allElements(el)) {
    for (let i = 0; i < e.attributes.length; i++) {
      const a = e.attributes[i];
      if (a.namespaceURI === R) out.push({ el: e, name: a.name, value: a.value });
    }
  }
  return out;
}

function resolvePart(target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = ["word", ...target.split("/")];
  const out: string[] = [];
  for (const p of parts) {
    if (p === "..") out.pop();
    else if (p !== "." && p !== "") out.push(p);
  }
  return out.join("/");
}

export async function loadDocxSource(data: ArrayBuffer | Uint8Array, name: string, sourceIndex: number): Promise<SourceExam> {
  const parsed = await parseDocx(data, name);
  const zip = await JSZip.loadAsync(data);
  const docXml = parseXml(await zip.file("word/document.xml")!.async("string"));
  const body = child(docXml.documentElement, "body");
  if (!body) throw new ComposeError(`${name}：找不到 Word 文件內容。`);
  const { blocks, blockOf } = mapParagraphsToBlocks(body);
  if (blockOf.length !== parsed.paragraphs.length) throw new ComposeError(`${name}：段落對應失敗，無法複製題目。`);

  const rels = await readRels(zip, "word/_rels/document.xml.rels");
  const readOpt = async (p: string) => {
    const f = zip.file(p);
    return f ? parseXml(await f.async("string")) : undefined;
  };
  const st = analyzeStructure(parsed.paragraphs);
  const questions: SourceQuestion[] = [];

  for (const sec of st.sections) {
    sec.questions.forEach((q, qi) => {
      const first = blockOf[q.startParagraph];
      let last = blockOf[q.endParagraph];
      // 下一題或下一個大題開始之前的空白段落是作答空間，一起帶走
      const next = sec.questions[qi + 1];
      const limit = next ? blockOf[next.startParagraph] : blocks.length;
      while (last + 1 < limit && isBlankParagraph(blocks[last + 1])) last++;
      const range = blocks.slice(first, last + 1).filter((b) => !is(b, W, "sectPr"));
      const paras = parsed.paragraphs.slice(q.startParagraph, q.endParagraph + 1);
      const lines = range.flatMap((b) => (is(b, W, "p") ? [b] : descendants(b, W, "p"))).map(previewParagraph);
      if (lines.length) lines[0] = lines[0].slice(questionPrefixLength(lines[0]));
      const text = lines.join("\n").replace(/^\s+|\s+$/g, "").replace(/\n{3,}/g, "\n\n");
      questions.push({
        id: `${sourceIndex}:${questions.length}`,
        sourceIndex,
        sectionTitle: sec.title,
        kind: guessKind(sec.kind, q.body),
        number: q.number,
        text,
        hasMath: paras.some((p) => p.hasMath) || range.some((b) => allElements(b).some((e) => is(e, W, "instrText") && /^\s*EQ\b/i.test(e.textContent ?? ""))) || /√/.test(text),
        imageCount: range.reduce(
          (n, b) =>
            n +
            allElements(b).filter((e) => (is(e, W, "drawing") || is(e, W, "pict") || is(e, W, "object"))).length,
          0,
        ),
        fingerprint: "",
        blocks: range,
      });
    });
  }

  // 重複題比對：文字 + 圖片內容（只有圖片不同的題目不會被當成重複）
  for (const q of questions) {
    const media: string[] = [];
    for (const b of q.blocks!) {
      for (const a of relAttrs(b)) {
        const rel = rels.get(a.value);
        const f = rel && !rel.external ? zip.file(resolvePart(rel.target)) : null;
        if (f) media.push(fnv(await f.async("uint8array")));
      }
    }
    q.fingerprint = `${normalizeForCompare(q.text)}|${media.join(",")}`;
  }

  return {
    index: sourceIndex,
    name,
    format: "docx",
    questions,
    zip,
    numbering: await readOpt("word/numbering.xml"),
    styles: await readOpt("word/styles.xml"),
    rels,
  };
}

/** PDF 只有文字：換行接續的行併回同一段 */
export function loadPdfSource(parsed: ParsedDocument, sourceIndex: number): SourceExam {
  const st: ExamStructure = analyzeStructure(parsed.paragraphs);
  const questions: SourceQuestion[] = [];
  for (const sec of st.sections) {
    for (const q of sec.questions) {
      const lines: string[] = [];
      for (const p of parsed.paragraphs.slice(q.startParagraph, q.endParagraph + 1)) {
        const t = p.text.trim();
        if (!t) continue;
        if (p.wrapped && lines.length) lines[lines.length - 1] += t;
        else lines.push(t);
      }
      if (lines.length) lines[0] = lines[0].slice(questionPrefixLength(lines[0]));
      const text = q.body.trim();
      questions.push({
        id: `${sourceIndex}:${questions.length}`,
        sourceIndex,
        sectionTitle: sec.title,
        kind: guessKind(sec.kind, q.body),
        number: q.number,
        text,
        hasMath: false,
        imageCount: 0,
        fingerprint: `${normalizeForCompare(text)}|`,
        lines,
      });
    }
  }
  return { index: sourceIndex, name: parsed.fileName, format: "pdf", questions };
}

/** 標出重複題：同樣內容只有第一次出現的不算重複 */
export function markDuplicates(sources: SourceExam[]) {
  const seen = new Map<string, string>();
  for (const s of sources) {
    for (const q of s.questions) {
      q.duplicateOf = undefined;
      const key = q.fingerprint;
      if (key === "|") continue; // 沒有文字也沒有圖
      const first = seen.get(key);
      if (first) q.duplicateOf = first;
      else seen.set(key, q.id);
    }
  }
}

/** 題目要放進版型的哪一個大題 */
export function defaultSectionFor(kind: SectionKind, sections: TemplateSection[]): number {
  const exact = sections.findIndex((s) => s.kind === kind);
  if (exact >= 0) return exact;
  if (kind !== "choice") {
    const nonChoice = sections.findIndex((s) => s.kind !== "choice");
    if (nonChoice >= 0) return nonChoice;
  }
  return 0;
}

// ─── 組卷 ────────────────────────────────────────────────

export interface ComposeRequest {
  template: ArrayBuffer | Uint8Array;
  templateName: string;
  info: ExamInfo;
  /** 依大題排好的題目（陣列索引對應版型的大題） */
  sections: SourceQuestion[][];
  sources: SourceExam[];
  /** 各大題的每題配分（陣列索引對應版型的大題）；沒給的大題配分留空格 */
  scores?: (number | undefined)[];
}

/** 沒指定配分時，大題標題裡的配分數字換成空格，讓老師自己填 */
const SCORE_BLANK = "____";
const NUM = String.raw`\d+(?:\.\d+)?`;
const UNIT = String.raw`(?:points?|pts?\.?|marks?)`;
// 先換成記號再填數字，總分填進去後才不會又被當成每題配分換掉
const PER_MARK = "";
const TOTAL_MARK = "";
/** 大題總分：「共 40 分」「Total: 40」「40 points total」「(40 marks)」「40%」 */
const TOTAL_PATTERNS: RegExp[] = [
  new RegExp(String.raw`(共\s*)${NUM}(\s*分)`, "g"),
  new RegExp(String.raw`(total(?:\s+of)?\s*[:：]?\s*)${NUM}`, "gi"),
  new RegExp(String.raw`(^|[^\w.])${NUM}(\s*${UNIT}\s*(?:in\s+)?total)`, "gi"),
  new RegExp(String.raw`([(（]\s*)${NUM}(\s*${UNIT}\s*[)）])`, "gi"),
  new RegExp(String.raw`(^|[^\w.])${NUM}(\s*[%％])`, "g"),
];
/** 每題配分：「每題 4 分」「每格 2 分」「4 points each」 */
const PER_ZH = new RegExp(String.raw`(每小?題|每(?:一|個)?(?:格|空格?|小格))(\s*)${NUM}(\s*分)`, "g");
const PER_EN = new RegExp(String.raw`(^|[^\w.])${NUM}(\s*${UNIT}\b)`, "gi");
const PER_BLANK_EN = /\b(per|each|for\s+each)(\s+)(?:blank|box)(?:es)?\b/gi;
const QUESTION_SCORE = new RegExp(String.raw`([（(\[]\s*)${NUM}(\s*(?:分|${UNIT})\s*[)）\]])`, "gi");

/** 分數顯示：最多兩位小數，避免 0.1 + 0.2 的浮點誤差 */
export const formatScore = (n: number) => String(Math.round(n * 100) / 100);

/** 大題總分 = 每題配分 × 題數（題數是這次選進來的題目數，跟原試卷題號無關） */
export function sectionTotal(perQuestion: number | undefined, count: number): number | undefined {
  return perQuestion === undefined ? undefined : Math.round(perQuestion * count * 100) / 100;
}

/**
 * 大題標題的配分：有指定每題配分就填入每題分數和總分，「每格」改成「每題」；
 * 沒指定的留空格。
 */
export function fillScores(p: Element, perQuestion?: number, total?: number) {
  for (const re of TOTAL_PATTERNS) replaceInParagraph(p, re, (m) => `${m[1]}${TOTAL_MARK}${m[2] ?? ""}`);
  const per = perQuestion !== undefined;
  replaceInParagraph(p, PER_ZH, (m) => `${per ? "每題" : m[1]}${m[2]}${PER_MARK}${m[3]}`);
  replaceInParagraph(p, PER_EN, (m) => `${m[1]}${PER_MARK}${m[2]}`);
  if (per) replaceInParagraph(p, PER_BLANK_EN, (m) => `${m[1]}${m[2]}question`);
  replaceInParagraph(p, new RegExp(PER_MARK, "g"), () => (per ? formatScore(perQuestion) : SCORE_BLANK));
  replaceInParagraph(p, new RegExp(TOTAL_MARK, "g"), () => (total !== undefined ? formatScore(total) : SCORE_BLANK));
}


interface SectionNumbering {
  /** 版型題目第一段的段落格式（不含編號），套用到每題第一段 */
  pPr?: Element;
  /** 第一段文字的格式（PDF 題目用） */
  rPr?: Element;
  numId?: string;
  ilvl: string;
  /** 版型沒用自動編號時，手打題號的樣式，例如 "(　　)#." */
  manualPrefix?: string;
  /** 版型每題後面的空白行數（選擇題通常 1 行，計算題是作答空間） */
  blankAfter: number;
  /** 版型的空白段落（沿用字級、行距），補空白行時複製它 */
  blankPara?: Element;
}

class Composer {
  private numbering?: Document;
  private styleIds = new Set<string>();
  private rels!: Document;
  private contentTypes!: Document;
  private nextRel = 1;
  private nextNum = 1;
  private nextAbstract = 0;
  private copiedParts = new Map<string, string>();
  private numMap = new Map<string, string>();
  private shapeSeq = 0;
  private counter = 0;

  constructor(
    private zip: JSZip,
    private doc: Document,
  ) {}

  async init() {
    const numFile = this.zip.file("word/numbering.xml");
    if (numFile) this.numbering = parseXml(await numFile.async("string"));
    const stylesFile = this.zip.file("word/styles.xml");
    if (stylesFile) {
      const styles = parseXml(await stylesFile.async("string"));
      for (const s of kids(styles.documentElement!)) {
        const id = wAttr(s, "styleId");
        if (id) this.styleIds.add(id);
      }
    }
    this.rels = parseXml(await this.zip.file("word/_rels/document.xml.rels")!.async("string"));
    for (const r of kids(this.rels.documentElement!)) {
      const n = Number((r.getAttribute("Id") ?? "").replace(/^rId/, ""));
      if (n >= this.nextRel) this.nextRel = n + 1;
    }
    this.contentTypes = parseXml(await this.zip.file("[Content_Types].xml")!.async("string"));
    if (this.numbering) {
      for (const el of kids(this.numbering.documentElement!)) {
        if (is(el, W, "num")) this.nextNum = Math.max(this.nextNum, Number(wAttr(el, "numId")) + 1);
        if (is(el, W, "abstractNum")) this.nextAbstract = Math.max(this.nextAbstract, Number(wAttr(el, "abstractNumId")) + 1);
      }
    }
  }

  private ensureNumbering(): Document {
    if (!this.numbering) {
      this.numbering = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"/>`);
      this.addRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml");
      this.addOverride("/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
    }
    return this.numbering;
  }

  private addRel(type: string, target: string, external = false): string {
    const id = `rId${this.nextRel++}`;
    const r = this.rels.createElementNS(REL, "Relationship");
    r.setAttribute("Id", id);
    r.setAttribute("Type", type);
    r.setAttribute("Target", target);
    if (external) r.setAttribute("TargetMode", "External");
    this.rels.documentElement!.appendChild(r);
    return id;
  }

  private addOverride(partName: string, contentType: string) {
    const o = this.contentTypes.createElementNS(CT, "Override");
    o.setAttribute("PartName", partName);
    o.setAttribute("ContentType", contentType);
    this.contentTypes.documentElement!.appendChild(o);
  }

  private hasDefault(ext: string) {
    return kids(this.contentTypes.documentElement!).some(
      (e) => e.localName === "Default" && (e.getAttribute("Extension") ?? "").toLowerCase() === ext.toLowerCase(),
    );
  }

  /** 新的 w:num，指向某個 abstractNum，並從 1 重新開始編號 */
  newNum(abstractId: string, restartLevel?: string): string {
    const numbering = this.ensureNumbering();
    const id = String(this.nextNum++);
    const num = wEl(numbering, "num", { numId: id });
    num.appendChild(wEl(numbering, "abstractNumId", { val: abstractId }));
    if (restartLevel !== undefined) {
      const ov = wEl(numbering, "lvlOverride", { ilvl: restartLevel });
      ov.appendChild(wEl(numbering, "startOverride", { val: "1" }));
      num.appendChild(ov);
    }
    const root = numbering.documentElement!;
    const tail = kids(root).find((e) => is(e, W, "numIdMacAtCleanup"));
    root.insertBefore(num, tail ?? null);
    return id;
  }

  abstractOf(numId: string): string | undefined {
    const num = this.numbering && kids(this.numbering.documentElement!).find((e) => is(e, W, "num") && wAttr(e, "numId") === numId);
    return wAttr(child(num, "abstractNumId"), "val") ?? undefined;
  }

  /** 把來源試卷的編號定義搬過來（題目裡的選項、小題用到自動編號時） */
  private importNum(src: SourceExam, numId: string): string | undefined {
    const key = `${src.name}#${numId}`;
    if (this.numMap.has(key)) return this.numMap.get(key);
    const srcRoot = src.numbering?.documentElement;
    const srcNum = srcRoot && kids(srcRoot).find((e) => is(e, W, "num") && wAttr(e, "numId") === numId);
    const absId = wAttr(child(srcNum, "abstractNumId"), "val");
    const srcAbs = srcRoot && kids(srcRoot).find((e) => is(e, W, "abstractNum") && wAttr(e, "abstractNumId") === absId);
    if (!srcNum || !srcAbs || child(srcAbs, "numStyleLink")) {
      this.numMap.set(key, "");
      return undefined;
    }
    const numbering = this.ensureNumbering();
    const root = numbering.documentElement!;
    const abs = numbering.importNode(srcAbs, true) as Element;
    const newAbs = String(this.nextAbstract++);
    setW(abs, "abstractNumId", newAbs);
    const nsid = child(abs, "nsid");
    if (nsid) setW(nsid, "val", (0x10000000 + Math.floor(Math.random() * 0xefffffff)).toString(16).toUpperCase());
    const styleLink = child(abs, "styleLink");
    if (styleLink) remove(styleLink);
    for (const ps of descendants(abs, W, "pStyle")) remove(ps);
    const firstNum = kids(root).find((e) => is(e, W, "num"));
    root.insertBefore(abs, firstNum ?? null);

    const num = numbering.importNode(srcNum, true) as Element;
    for (const a of [...Array.from({ length: num.attributes.length }, (_, i) => num.attributes[i])])
      if (a.localName === "durableId") num.removeAttributeNode(a);
    const newId = String(this.nextNum++);
    setW(num, "numId", newId);
    setW(child(num, "abstractNumId")!, "val", newAbs);
    const tail = kids(root).find((e) => is(e, W, "numIdMacAtCleanup"));
    root.insertBefore(num, tail ?? null);
    this.numMap.set(key, newId);
    return newId;
  }

  private async copyPart(src: SourceExam, relId: string): Promise<string | undefined> {
    const rel = src.rels?.get(relId);
    if (!rel) return undefined;
    if (rel.external) return this.addRel(rel.type, rel.target, true);
    const key = `${src.name}#${rel.target}`;
    const done = this.copiedParts.get(key);
    if (done) return done;
    const srcPath = resolvePart(rel.target);
    const file = src.zip!.file(srcPath);
    if (!file) return undefined;
    const base = srcPath.split("/").pop()!;
    const dir = srcPath.split("/").slice(1, -1).join("/") || "media";
    const newName = `${dir}/compose${++this.counter}_${base}`;
    this.zip.file(`word/${newName}`, await file.async("uint8array"));
    const ext = base.includes(".") ? base.split(".").pop()! : "";
    if (!this.hasDefault(ext)) {
      const srcCT = parseXml(await src.zip!.file("[Content_Types].xml")!.async("string"));
      const entries = kids(srcCT.documentElement!);
      const override = entries.find((e) => e.localName === "Override" && e.getAttribute("PartName") === `/${srcPath}`);
      const def = entries.find(
        (e) => e.localName === "Default" && (e.getAttribute("Extension") ?? "").toLowerCase() === ext.toLowerCase(),
      );
      const type = override?.getAttribute("ContentType") ?? def?.getAttribute("ContentType");
      if (type) this.addOverride(`/word/${newName}`, type);
    }
    const id = this.addRel(rel.type, newName);
    this.copiedParts.set(key, id);
    return id;
  }

  /** 複製來源試卷一題的所有元素到版型文件 */
  async importQuestion(
    q: SourceQuestion,
    src: SourceExam,
    numbering: SectionNumbering,
    index: number,
    perQuestion?: number,
  ): Promise<Element[]> {
    if (q.lines) return this.pdfQuestion(q.lines, numbering, index);
    const out: Element[] = [];
    for (const b of q.blocks!) out.push(this.doc.importNode(b, true) as Element);
    const first = out.find((e) => is(e, W, "p"));
    // 題號由版型的自動編號重新產生：拿掉原本的編號（自動編號或手打的題號）
    if (first) {
      const num = child(child(first, "pPr"), "numPr");
      if (num) remove(num);
      stripTypedNumber(first);
    }

    for (const el of out) {
      // 圖片、OLE 物件、超連結：搬檔案並換新的 relationship id
      for (const a of relAttrs(el)) {
        const id = await this.copyPart(src, a.value);
        if (id) a.el.setAttributeNS(R, a.name, id);
        else a.el.removeAttributeNS(R, a.name.split(":").pop()!);
      }
      this.cleanup(el, src);
    }
    if (first) this.applyQuestionNumber(first, numbering, index);
    for (const el of out) for (const p of [...(is(el, W, "p") ? [el] : []), ...descendants(el, W, "p")]) {
      // 題目後面原本的「（8 分）」：換成這次的每題配分，沒指定就留空格
      replaceInParagraph(p, QUESTION_SCORE, (m) => `${m[1]}${perQuestion !== undefined ? formatScore(perQuestion) : SCORE_BLANK}${m[2]}`);
    }
    // 空白行比版型少時補足（原卷題目之間沒空行、計算題沒留作答空間）
    let blanks = 0;
    for (let i = out.length - 1; i >= 0 && isBlankParagraph(out[i]); i--) blanks++;
    out.push(...this.blankLines(numbering, numbering.blankAfter - blanks));
    return out;
  }

  private applyQuestionNumber(p: Element, numbering: SectionNumbering, index: number) {
    const old = child(p, "pPr");
    let pPr: Element;
    if (numbering.pPr) {
      pPr = numbering.pPr.cloneNode(true) as Element;
      // 段落內的文字格式標記（rPr）沿用原題
      const oldRPr = old && child(old, "rPr");
      const tplRPr = child(pPr, "rPr");
      if (tplRPr) remove(tplRPr);
      if (oldRPr) pPr.appendChild(oldRPr.cloneNode(true));
      if (old) p.replaceChild(pPr, old);
      else p.insertBefore(pPr, p.firstChild);
    } else {
      pPr = old ?? wEl(this.doc, "pPr");
      if (!old) p.insertBefore(pPr, p.firstChild);
    }
    const existing = child(pPr, "numPr");
    if (existing) remove(existing);
    if (numbering.numId) {
      const numPr = wEl(this.doc, "numPr");
      numPr.appendChild(wEl(this.doc, "ilvl", { val: numbering.ilvl }));
      numPr.appendChild(wEl(this.doc, "numId", { val: numbering.numId }));
      // numPr 要放在 pStyle / keepNext 等之後；放在 pStyle 後面最保險
      const style = child(pPr, "pStyle");
      const order = ["keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl"];
      const after = [style, ...order.map((n) => child(pPr, n))].filter(Boolean).pop() as Element | undefined;
      pPr.insertBefore(numPr, after ? after.nextSibling : pPr.firstChild);
    } else if (numbering.manualPrefix) {
      const r = wEl(this.doc, "r");
      const t = wEl(this.doc, "t");
      t.setAttribute("xml:space", "preserve");
      t.appendChild(this.doc.createTextNode(numbering.manualPrefix.replace("#", String(index + 1))));
      r.appendChild(t);
      p.insertBefore(r, pPr.nextSibling);
    }
  }

  private pdfQuestion(lines: string[], numbering: SectionNumbering, index: number): Element[] {
    const out: Element[] = [];
    lines.forEach((line, i) => {
      const p = wEl(this.doc, "p");
      const r = wEl(this.doc, "r");
      if (numbering.rPr) r.appendChild(numbering.rPr.cloneNode(true));
      const t = wEl(this.doc, "t");
      t.setAttribute("xml:space", "preserve");
      t.appendChild(this.doc.createTextNode(line));
      r.appendChild(t);
      p.appendChild(r);
      if (i === 0) this.applyQuestionNumber(p, numbering, index);
      else if (numbering.pPr) {
        const ind = child(numbering.pPr, "ind");
        if (ind) {
          const pPr = wEl(this.doc, "pPr");
          pPr.appendChild(wEl(this.doc, "ind", { left: wAttr(ind, "left") ?? wAttr(ind, "start") ?? "0" }));
          p.insertBefore(pPr, p.firstChild);
        }
      }
      out.push(p);
    });
    out.push(...this.blankLines(numbering, numbering.blankAfter));
    return out;
  }

  private blankLines(numbering: SectionNumbering, n: number): Element[] {
    return Array.from({ length: Math.max(0, n) }, () =>
      numbering.blankPara ? (numbering.blankPara.cloneNode(true) as Element) : wEl(this.doc, "p"),
    );
  }

  /** 移除會在新文件裡失效或衝突的東西 */
  private cleanup(el: Element, src: SourceExam) {
    for (const e of allElements(el)) {
      if (e.namespaceURI === W) {
        switch (e.localName) {
          case "bookmarkStart":
          case "bookmarkEnd":
          case "commentRangeStart":
          case "commentRangeEnd":
          case "sectPr":
          case "pageBreakBefore":
            remove(e);
            continue;
          case "commentReference":
          case "footnoteReference":
          case "endnoteReference":
            remove(e.parentNode && is(e.parentNode, W, "r") ? e.parentNode : e);
            continue;
          case "br":
            if (wAttr(e, "type") === "page") remove(e);
            continue;
          case "pStyle":
          case "rStyle":
          case "tblStyle": {
            const id = wAttr(e, "val");
            if (!id || !this.styleIds.has(id)) remove(e);
            continue;
          }
          case "numPr": {
            const numId = wAttr(child(e, "numId"), "val");
            const mapped = numId && numId !== "0" ? this.importNum(src, numId) : undefined;
            if (mapped) setW(child(e, "numId")!, "val", mapped);
            else if (numId !== "0") remove(e);
            continue;
          }
        }
      }
      // Word 用來追蹤段落的 id，重複時可能要修復文件
      e.removeAttributeNS(W14, "paraId");
      e.removeAttributeNS(W14, "textId");
      // VML 圖形 id 不能重複（MathType、舊版圖片）
      if (is(e, V, "shape") && e.getAttribute("id")) {
        const oldId = e.getAttribute("id")!;
        const newId = `_x0000_i${2000 + ++this.shapeSeq}`;
        e.setAttribute("id", newId);
        for (const ole of descendants(el, O, "OLEObject")) if (ole.getAttribute("ShapeID") === oldId) ole.setAttribute("ShapeID", newId);
      }
    }
  }

  async finish(): Promise<void> {
    // 圖片的 docPr id 在整份文件裡不能重複
    let id = 1;
    for (const d of descendants(this.doc.documentElement!, WP, "docPr")) d.setAttribute("id", String(id++));
    this.zip.file("word/document.xml", serialize(this.doc));
    this.zip.file("word/_rels/document.xml.rels", serialize(this.rels));
    this.zip.file("[Content_Types].xml", serialize(this.contentTypes));
    if (this.numbering) this.zip.file("word/numbering.xml", serialize(this.numbering));
  }
}

export async function composeExam(req: ComposeRequest): Promise<Uint8Array> {
  const parsed = await parseDocx(req.template, req.templateName);
  const st = analyzeStructure(parsed.paragraphs);
  const sections = st.sections.filter((s) => s.paragraphIndex >= 0);
  if (!sections.length) throw new ComposeError("版型裡找不到大題標題。");

  const zip = await JSZip.loadAsync(req.template);
  const doc = parseXml(await zip.file("word/document.xml")!.async("string"));
  const body = child(doc.documentElement, "body")!;
  const { blocks, blockOf } = mapParagraphsToBlocks(body);
  if (blockOf.length !== parsed.paragraphs.length) throw new ComposeError("版型段落對應失敗。");

  const composer = new Composer(zip, doc);
  await composer.init();

  const sectPr = blocks.find((b) => is(b, W, "sectPr"));
  const firstHeader = blockOf[sections[0].paragraphIndex];

  // 1. 大題標題前：學校標題、班級座號、作答說明 —— 只改期中／期末
  const prefix = blocks.slice(0, firstHeader);
  for (const b of prefix) for (const p of is(b, W, "p") ? [b] : descendants(b, W, "p")) setExamInfo(p, req.info);

  const output: Element[] = [...prefix];
  for (let si = 0; si < sections.length; si++) {
    const sec = sections[si];
    const headerStart = blockOf[sec.paragraphIndex];
    const nextSection = sections[si + 1];
    const headerEnd = sec.questions.length
      ? blockOf[sec.questions[0].startParagraph]
      : nextSection
        ? blockOf[nextSection.paragraphIndex]
        : blocks.length;
    // 2. 大題標題（和下一行的配分說明）：填入每題配分和總分，沒指定就留空格
    const picked = req.sections[si] ?? [];
    const perQuestion = req.scores?.[si];
    const total = sectionTotal(perQuestion, picked.length);
    const header = blocks.slice(headerStart, headerEnd).filter((b) => !is(b, W, "sectPr"));
    for (const b of header) if (is(b, W, "p")) fillScores(b, perQuestion, total);
    output.push(...header);

    // 3. 這個大題的編號方式：沿用版型第一題
    const numbering: SectionNumbering = { ilvl: "0", blankAfter: 1 };
    const tq = sec.questions[0];
    if (tq) {
      // 版型第一題後面有幾行空白，新題目就至少留幾行
      const stop = sec.questions[1]
        ? blockOf[sec.questions[1].startParagraph]
        : nextSection
          ? blockOf[nextSection.paragraphIndex]
          : blocks.length;
      let blankAfter = 0;
      for (let i = blockOf[tq.endParagraph] + 1; i < stop && isBlankParagraph(blocks[i]); i++) {
        if (!blankAfter) numbering.blankPara = blocks[i].cloneNode(true) as Element;
        blankAfter++;
      }
      numbering.blankAfter = blankAfter;
      const qp = blocks[blockOf[tq.startParagraph]];
      const pPr = is(qp, W, "p") ? child(qp, "pPr") : undefined;
      if (pPr) {
        numbering.pPr = pPr.cloneNode(true) as Element;
        const np = child(numbering.pPr, "numPr");
        if (np) remove(np);
      }
      numbering.rPr = (child(kids(qp).find((e) => is(e, W, "r") && paragraphText(e)), "rPr")?.cloneNode(true) ??
        undefined) as Element | undefined;
      const numPr = pPr && child(pPr, "numPr");
      const numId = wAttr(child(numPr, "numId"), "val");
      const label = parsed.paragraphs[tq.startParagraph].numberingLabel;
      if (numId && numId !== "0" && label) {
        const abs = composer.abstractOf(numId);
        numbering.ilvl = wAttr(child(numPr, "ilvl"), "val") ?? "0";
        if (abs !== undefined) numbering.numId = composer.newNum(abs, numbering.ilvl);
      } else {
        const text = parsed.paragraphs[tq.startParagraph].text;
        const n = leadingNumberLength(text);
        numbering.manualPrefix = n ? text.slice(0, n).replace(/\d+/, "#") : "#. ";
      }
    } else {
      numbering.manualPrefix = "#. ";
    }

    for (let qi = 0; qi < picked.length; qi++) {
      const q = picked[qi];
      const src = req.sources.find((s) => s.index === q.sourceIndex);
      if (!src) throw new ComposeError("找不到題目的來源試卷。");
      output.push(...(await composer.importQuestion(q, src, numbering, qi, perQuestion)));
    }
  }

  while (body.firstChild) body.removeChild(body.firstChild);
  for (const el of output) body.appendChild(el);
  if (sectPr) body.appendChild(sectPr);

  await composer.finish();
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
