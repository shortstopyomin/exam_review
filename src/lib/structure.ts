import type { ExamStructure, Paragraph, Question, Section, SectionKind } from "./types";

/** 大題標題：「一、選擇題（每題2分，共40分）」 */
const SECTION_RE = /^\s*([一二三四五六七八九十]{1,3})\s*[、.．]\s*(.+)$/;

/**
 * 題目開頭：「1.」「1、」「( )1.」「（ B ）12．」
 * (?!\d) 避免把小數「3.5」當成題號。
 */
const QUESTION_RE =
  /^\s*(?:[（(]\s*([A-Za-zＡ-Ｅ○ＯOＸX╳✕✗]?)\s*[)）]\s*)?(\d{1,3})\s*[.．、](?!\d)/;

// 只認「每題／每小題」；「每格」無法換算成題數，改看「共 Y 分」
const PER_SCORE_RE = /每小?題\s*(\d+(?:\.\d+)?)\s*分/;
const TOTAL_SCORE_RE = /共\s*(\d+(?:\.\d+)?)\s*分/;
const QUESTION_SCORE_RE = /[（(]\s*(\d+(?:\.\d+)?)\s*分\s*[)）]/;

export function detectKind(title: string): SectionKind {
  if (/選擇/.test(title)) return "choice";
  if (/是非/.test(title)) return "truefalse";
  if (/填充|填空/.test(title)) return "fill";
  if (/問答|簡答|計算|申論|題組|作文|閱讀|非選/.test(title)) return "short";
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

    const sec = !p.inTable && text.length <= 80 ? text.match(SECTION_RE) : null;
    if (sec) {
      finishQuestion();
      current = {
        title: text,
        paragraphIndex: p.index,
        kind: detectKind(sec[2]),
        perQuestionScore: toNum(text.match(PER_SCORE_RE)),
        totalScore: toNum(text.match(TOTAL_SCORE_RE)),
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
