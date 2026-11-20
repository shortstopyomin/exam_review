import type { ExamStructure, Paragraph, Question, Section, SectionKind } from "./types";

/**
 * 大題標題，例如：
 * 「一、選擇題（每題2分，共40分）」「貳、非選擇題」「第二部分 計算題」
 * 「Part 2計算題和應用題 / Calculation and word problems (每題5分，共60分)」
 * 「Part II: Word Problems」「Section B - Multiple Choice」
 * 第 1 個群組之後的文字用來判斷題型。
 */
const SECTION_RES: RegExp[] = [
  /^\s*[一二三四五六七八九十]{1,3}\s*[、.．]\s*(.+)$/,
  /^\s*[壹貳參肆伍陸柒捌玖拾]{1,3}\s*[、.．]\s*(.+)$/,
  /^\s*第\s*[一二三四五六七八九十\d]{1,3}\s*(?:部分|部份|大題)\s*[:：、.．\-–—]?\s*(.*)$/,
  /^\s*(?:part|section)\s*(?:\d{1,2}|[IVX]{1,4}|[A-H])(?![A-Za-z\d])\s*[:：.．、\-–—]?\s*(.*)$/i,
];

/**
 * 阿拉伯數字編號的大題：「1、 計算題：每題3分，共15分」「2、 畫畫看、做做看: 每題5分，共15分」
 * 和題號長得一樣，所以題型名稱後面要緊接著冒號或括號與配分說明（每…分／共…分）才算大題。
 */
const ARABIC_SECTION_RE = /^\s*\d{1,2}\s*[、.．](?!\d)\s*([^：:（(]{1,20}?)\s*[：:（(]\s*((?:每|共).*)$/;

export function matchSection(text: string): string | null {
  // 標題段落裡可能有手動換行（Shift+Enter），例如中文標題換行接英文說明：
  // 「一、在正確的□ 中打勾。每題4分，共20分⏎Choose the correct answer. Put a ✓ in the correct box.」
  text = text.replace(/\s*\n\s*/g, " ");
  for (const re of SECTION_RES) {
    const m = text.match(re);
    if (m) return m[1] ?? "";
  }
  const m = text.match(ARABIC_SECTION_RE);
  if (m && hasScore(parseScoreText(m[2]))) return m[1];
  return null;
}

/**
 * 題目開頭：「1.」「1、」「( )1.」「（ B ）12．」
 * (?!\d) 避免把小數「3.5」當成題號。
 */
const QUESTION_RE =
  /^\s*(?:[（(]\s*([A-Za-zＡ-Ｅ○ＯOＸX╳✕✗]?)\s*[)）]\s*)?(\d{1,3})\s*[.．、](?!\d)/;

const N = "(\\d+(?:\\.\\d+)?)";
const PTS = "(?:points?|pts?\\.?|marks?)";

export interface ScoreText {
  perQuestion?: number;
  /** 「每格 X 分」：只記錄，不換算；大題分數採用「共 Y 分」 */
  perBlank?: number;
  total?: number;
  conflicts: string[];
}

const ZH_PER_QUESTION = new RegExp(`每小?題\\s*${N}\\s*分`);
/** 「每格 2 分」「每空 2 分」「每個空格 2 分」 */
const ZH_PER_BLANK = new RegExp(`每(?:一|個)?(?:格|空格?|小格)\\s*${N}\\s*分`);
const ZH_TOTAL = new RegExp(`共\\s*${N}\\s*分`);
/** 「2 points per blank」「each blank is worth 2 points」 */
const EN_PER_BLANK = [
  new RegExp(`${N}\\s*${PTS}\\s*(?:per|for\\s+each|each|/)\\s*(?:blank|box)`, "i"),
  new RegExp(`each\\s+(?:blank|box)(?:\\s+is)?(?:\\s+worth)?\\s*[:：]?\\s*${N}\\s*${PTS}`, "i"),
];
/** 「5 points per question」「each question is worth 5 points」 */
const EN_PER_QUESTION = [
  new RegExp(`${N}\\s*${PTS}\\s*(?:per\\s+question|/\\s*question|for\\s+each\\s+question|each\\s+question)`, "i"),
  new RegExp(`each\\s+question(?:\\s+is)?(?:\\s+worth)?\\s*[:：]?\\s*${N}\\s*${PTS}`, "i"),
];
/** 「2 points each」：沒說是每題還是每格，跟著中文描述的單位 */
const EN_PER_EACH = [
  new RegExp(`${N}\\s*${PTS}\\s*(?:each|apiece)(?!\\s+(?:question|blank|box))`, "i"),
  new RegExp(`each(?:\\s+is)?(?:\\s+worth)?\\s*[:：]?\\s*${N}\\s*${PTS}`, "i"),
];
// 英文：「60 points total」「total 60 points」「Total: 60」「(60 marks)」
const EN_TOTAL = [
  new RegExp(`${N}\\s*${PTS}\\s*(?:in\\s+)?total`, "i"),
  new RegExp(`total(?:\\s+of)?\\s*[:：]?\\s*${N}(?:\\s*${PTS})?`, "i"),
  new RegExp(`[(（]\\s*${N}\\s*${PTS}\\s*[)）]`, "i"),
];

function firstMatch(text: string, res: RegExp | RegExp[]): { value: number; text: string } | undefined {
  for (const re of ([] as RegExp[]).concat(res)) {
    const m = text.match(re);
    if (m) return { value: Number(m[1]), text: m[0].trim() };
  }
  return undefined;
}

/**
 * 解析配分描述，例如：
 * 「每題5分，共60分」「每格 2 分、共 30 分｜2 points each, 30 points total」
 * 「5 points each, 60 points total」
 * 中英文都有寫時，比對兩邊的分數、單位是否一致。
 */
export function parseScoreText(text: string): ScoreText {
  const zhQ = firstMatch(text, ZH_PER_QUESTION);
  const zhB = firstMatch(text, ZH_PER_BLANK);
  const zhT = firstMatch(text, ZH_TOTAL);
  const enB = firstMatch(text, EN_PER_BLANK);
  const enQ = firstMatch(text, EN_PER_QUESTION);
  const enE = enB || enQ ? undefined : firstMatch(text, EN_PER_EACH);
  const enT = firstMatch(text, EN_TOTAL);
  const conflicts: string[] = [];

  // 英文「each」的單位跟著中文：每格 → 每格；每題或沒寫 → 每題
  const zhUnit = zhB ? "blank" : zhQ ? "question" : undefined;
  const enPerBlank = enB ?? (enE && zhUnit === "blank" ? enE : undefined);
  const enPerQuestion = enQ ?? (enE && zhUnit !== "blank" ? enE : undefined);

  if (zhB && enQ) conflicts.push(`中文寫「${zhB.text}」（每格），英文寫「${enQ.text}」（每題），單位不一致`);
  if (zhQ && enB) conflicts.push(`中文寫「${zhQ.text}」（每題），英文寫「${enB.text}」（每格），單位不一致`);
  const zhPer = zhB ?? zhQ;
  const enPer = enPerBlank ?? enPerQuestion;
  if (zhPer && enPer && zhPer.value !== enPer.value) {
    conflicts.push(`中文寫「${zhPer.text}」，英文寫「${enPer.text}」，分數不一致`);
  }
  if (zhT && enT && zhT.value !== enT.value) {
    conflicts.push(`中文寫「${zhT.text}」，英文寫「${enT.text}」，總分不一致`);
  }

  return {
    perQuestion: zhQ?.value ?? enPerQuestion?.value,
    perBlank: zhB?.value ?? enPerBlank?.value,
    total: zhT?.value ?? enT?.value,
    conflicts,
  };
}

function hasScore(s: ScoreText) {
  return s.perQuestion !== undefined || s.perBlank !== undefined || s.total !== undefined;
}

/** 單題配分：「（8 分）」「(5 points)」「[5 marks]」 */
const QUESTION_SCORE_RE = new RegExp(`[（(\\[]\\s*${N}\\s*(?:分|${PTS})\\s*[)）\\]]`, "i");

export function detectKind(title: string): SectionKind {
  // 「非選擇題」要先判斷，否則會被當成選擇題
  if (/非選/.test(title)) return "short";
  if (/選擇|multiple[-\s]?choice|\bMCQ/i.test(title)) return "choice";
  if (/是非|true\s*(?:or|\/|and)?\s*false|\bT\s*\/\s*F\b/i.test(title)) return "truefalse";
  if (/填充|填空|fill[-\s]?in|blanks?/i.test(title)) return "fill";
  if (/問答|簡答|計算|應用|申論|題組|作文|閱讀|calculation|word\s+problems?|short[-\s]answer|essay|open[-\s]ended|problem\s+solving/i.test(title))
    return "short";
  return "other";
}

function toNum(m: RegExpMatchArray | null): number | undefined {
  return m ? Number(m[1]) : undefined;
}

/** 題號後面直接接數字：「8.3盆仙人掌…」「9.1塊香草蛋糕…」，看起來像小數 */
const QUESTION_BEFORE_DIGIT_RE = new RegExp(QUESTION_RE.source.replace("(?!\\d)", "(?=\\d)"));

/**
 * 解析題號。題號後面直接接數字時（「8.3盆…」）可能是小數，
 * 只有在給了 expectedNext 且題號剛好等於它（接續上一題）時才當成題目。
 */
export function parseQuestionStart(
  text: string,
  expectedNext?: number,
): { number: number; answerMark: string; rest: string } | null {
  let m = text.match(QUESTION_RE);
  if (!m && expectedNext !== undefined) {
    m = text.match(QUESTION_BEFORE_DIGIT_RE);
    if (m && Number(m[2]) !== expectedNext) m = null;
  }
  if (!m) return null;
  return { number: Number(m[2]), answerMark: m[1] ?? "", rest: text.slice(m[0].length) };
}

/** 作答說明的標題：「作答說明 / Instructions」「注意事項：」「Directions」 */
const INSTRUCTION_HEADER_RE =
  /^\s*(?:作答說明|作答須知|注意事項|考生須知|考試說明|instructions?|directions?|general\s+instructions?|notes?\s+to\s+(?:students|candidates))\s*(?:[:：/／]|$)/i;

/**
 * 常見的作答說明句子。出現在任何地方都不算題目，也不會被併進前一題。
 * 刻意寫得很具體：「Complete the table below.」是題目，不能被誤判。
 */
const INSTRUCTION_SENTENCE_RE = new RegExp(
  [
    String.raw`^(?:please\s+)?(?:complete|answer|attempt)\s+all\s+(?:the\s+)?questions`,
    String.raw`^(?:please\s+)?(?:read\s+)?(?:all\s+)?(?:the\s+)?questions?\s+carefully`,
    String.raw`^(?:please\s+)?write\s+(?:all\s+)?(?:of\s+)?your\s+answers?`,
    String.raw`^(?:please\s+)?show\s+(?:all\s+)?(?:of\s+)?your\s+(?:work|working|steps)`,
    String.raw`answer\s+sheet|answer\s+booklet`,
    String.raw`calculators?\s+(?:is|are|may)\s+(?:not\s+)?(?:allowed|permitted|be\s+used)`,
    String.raw`^time\s+allowed|^duration\s*[:：]`,
    String.raw`^本試卷|試卷共\s*\d+\s*頁|共\s*\d+\s*頁`,
    String.raw`答案卡|答案卷|答案紙|答案欄|答案寫在|作答於|請將答案`,
    String.raw`^請(?:用|以|使用)[一-鿿]{0,6}筆`,
    String.raw`^考試時間|^作答時間`,
    String.raw`^(?:請)?(?:仔細)?閱讀題目|^請詳讀`,
  ].join("|"),
  "i",
);

export function isInstruction(text: string): boolean {
  const t = text.trim();
  if (INSTRUCTION_HEADER_RE.test(t)) return true;
  const q = parseQuestionStart(t);
  return INSTRUCTION_SENTENCE_RE.test(q ? q.rest.trim() : t);
}

export function analyzeStructure(paragraphs: Paragraph[]): ExamStructure {
  const sections: Section[] = [];
  const instructions: number[] = [];
  let current: Section | undefined;
  let question: Question | undefined;

  const isSectionHeader = (p: Paragraph) => {
    const text = p.text.trim();
    // 中英雙語標題比較長，放寬到 150 字
    return !p.inTable && text.length > 0 && text.length <= 150 ? matchSection(text) : null;
  };
  // 有大題標題時，第一個大題之前的編號段落都是前言／作答說明，不算題目
  const firstSection = paragraphs.find((p) => isSectionHeader(p) !== null)?.index ?? Infinity;
  let inInstructionBlock = false;

  const finishQuestion = () => {
    if (question && current) {
      question.score = toNum(question.body.match(QUESTION_SCORE_RE));
      current.questions.push(question);
    }
    question = undefined;
  };

  for (const p of paragraphs) {
    const text = p.text.trim();
    if (!text) continue;

    const sec = isSectionHeader(p);
    if (sec !== null) {
      finishQuestion();
      inInstructionBlock = false;
      current = {
        title: text,
        paragraphIndex: p.index,
        kind: detectKind(sec),
        ...scoreFields(parseScoreText(text)),
        questions: [],
      };
      sections.push(current);
      continue;
    }

    // 下一題的題號：用來分辨「8.3盆仙人掌…」是第 8 題還是小數
    const expectedNext = (question?.number ?? current?.questions.at(-1)?.number ?? 0) + 1;
    const q = parseQuestionStart(text, expectedNext);

    // 配分寫在大題標題的下一行（第一題之前），例如：
    // 「Part 3 填充題 / Fill in the blanks」
    // 「每格 2 分、共 30 分｜2 points each, 30 points total」
    if (current && !current.questions.length && !question && !q) {
      const score = parseScoreText(text);
      if (hasScore(score)) {
        current.perQuestionScore ??= score.perQuestion;
        current.perBlankScore ??= score.perBlank;
        current.totalScore ??= score.total;
        if (score.conflicts.length) current.scoreConflicts = [...(current.scoreConflicts ?? []), ...score.conflicts];
        current.scoreLines = [...(current.scoreLines ?? []), p.index];
        continue;
      }
    }

    const beforeFirstSection = p.index < firstSection && firstSection !== Infinity;
    if (
      isInstruction(text) ||
      (inInstructionBlock && !current) ||
      (q && beforeFirstSection)
    ) {
      finishQuestion();
      if (INSTRUCTION_HEADER_RE.test(text)) inInstructionBlock = true;
      instructions.push(p.index);
      continue;
    }

    if (q) {
      finishQuestion();
      if (!current) {
        current = { title: "（未分大題）", paragraphIndex: -1, kind: "other", questions: [] };
        sections.push(current);
      }
      question = {
        number: q.number,
        answerMark: q.answerMark,
        startParagraph: p.index,
        endParagraph: p.index,
        body: q.rest,
      };
      continue;
    }

    if (question) {
      question.endParagraph = p.index;
      question.body += "\n" + text;
    }
  }
  finishQuestion();
  return { sections, instructions };
}

function scoreFields(score: ScoreText): Partial<Section> {
  return {
    perQuestionScore: score.perQuestion,
    perBlankScore: score.perBlank,
    totalScore: score.total,
    ...(score.conflicts.length ? { scoreConflicts: score.conflicts } : {}),
  };
}

/**
 * 大題總分：優先採用「共 Y 分」；沒寫時才用「每題 X 分 × 題數」或各題「（X 分）」加總推算。
 * 只用來加總全卷分數，不檢查大題內的題數是否符合配分。
 * 「每格 X 分」不換算（不數格子）。
 */
export function sectionScore(s: Section): { total?: number } {
  if (s.totalScore !== undefined) return { total: s.totalScore };
  if (s.perQuestionScore !== undefined) return { total: s.perQuestionScore * s.questions.length };
  if (s.perBlankScore === undefined && s.questions.length && s.questions.every((q) => q.score !== undefined))
    return { total: s.questions.reduce((a, q) => a + (q.score ?? 0), 0) };
  return {};
}
