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

export function matchSection(text: string): string | null {
  for (const re of SECTION_RES) {
    const m = text.match(re);
    if (m) return m[1] ?? "";
  }
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

// 只認「每題／每小題」；「每格」無法換算成題數，改看「共 Y 分」
// 英文：「5 points each」「each question 5 points」「5 pts/question」
const PER_SCORE_RES = [
  new RegExp(`每小?題\\s*${N}\\s*分`),
  new RegExp(`${N}\\s*${PTS}\\s*(?:each|per\\s+question|/\\s*question|for\\s+each)`, "i"),
  new RegExp(`each(?:\\s+question)?(?:\\s+is)?(?:\\s+worth)?\\s*[:：]?\\s*${N}\\s*${PTS}`, "i"),
];
// 英文：「60 points total」「total 60 points」「Total: 60」「(60 marks)」
const TOTAL_SCORE_RES = [
  new RegExp(`共\\s*${N}\\s*分`),
  new RegExp(`${N}\\s*${PTS}\\s*(?:in\\s+)?total`, "i"),
  new RegExp(`total(?:\\s+of)?\\s*[:：]?\\s*${N}(?:\\s*${PTS})?`, "i"),
  new RegExp(`[(（]\\s*${N}\\s*${PTS}\\s*[)）]`, "i"),
];

function firstNum(text: string, res: RegExp[]): number | undefined {
  for (const re of res) {
    const m = text.match(re);
    if (m) return Number(m[1]);
  }
  return undefined;
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

export function parseQuestionStart(text: string): { number: number; answerMark: string; rest: string } | null {
  const m = text.match(QUESTION_RE);
  if (!m) return null;
  return { number: Number(m[2]), answerMark: m[1] ?? "", rest: text.slice(m[0].length) };
}

export function analyzeStructure(paragraphs: Paragraph[]): ExamStructure {
  const sections: Section[] = [];
  let current: Section | undefined;
  let question: Question | undefined;

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

    // 中英雙語標題比較長，放寬到 150 字
    const sec = !p.inTable && text.length <= 150 ? matchSection(text) : null;
    if (sec !== null) {
      finishQuestion();
      current = {
        title: text,
        paragraphIndex: p.index,
        kind: detectKind(sec),
        perQuestionScore: firstNum(text, PER_SCORE_RES),
        totalScore: firstNum(text, TOTAL_SCORE_RES),
        questions: [],
      };
      sections.push(current);
      continue;
    }

    const q = parseQuestionStart(text);
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
  return { sections };
}

/** 依大題配分推算該大題總分；無法判斷時回傳 undefined */
export function sectionScore(s: Section): { total?: number; computed?: number } {
  let computed: number | undefined;
  if (s.perQuestionScore !== undefined) computed = s.perQuestionScore * s.questions.length;
  else if (s.questions.length && s.questions.every((q) => q.score !== undefined))
    computed = s.questions.reduce((a, q) => a + (q.score ?? 0), 0);
  return { total: s.totalScore ?? computed, computed };
}
