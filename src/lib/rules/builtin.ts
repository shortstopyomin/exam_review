import type { RuleDefinition } from "./engine";
import { sectionScore } from "../structure";
import type { RuleConfig } from "../types";
import { EXAM_TITLE } from "./title";
import { CHINESE_SUBJECTS, ENGLISH_SUBJECTS, checkChineseSubjects, checkEnglishSubjects } from "./subjects";

function excerptAround(text: string, start: number, end: number, pad = 12): string {
  const s = Math.max(0, start - pad);
  const e = Math.min(text.length, end + pad);
  return (s > 0 ? "…" : "") + text.slice(s, e) + (e < text.length ? "…" : "");
}

function short(text: string, max = 40): string {
  const t = text.replace(/\s+/g, " ").trim();
  return t.length > max ? t.slice(0, max) + "…" : t;
}

function compileRegex(pattern: unknown, flags: unknown): RegExp {
  const f = typeof flags === "string" ? flags : "";
  return new RegExp(String(pattern), f.includes("g") ? f : f + "g");
}

function validateRegex(c: RuleConfig, key: string): string | undefined {
  try {
    compileRegex(c[key], c.flags);
  } catch (e) {
    return `${key} 不是有效的正規表示式：${(e as Error).message}`;
  }
}

const SCORE_SUM: RuleDefinition = {
  description: "各大題配分加總需等於指定分數；並檢查「每題 X 分」×題數是否與「共 Y 分」相符",
  defaultSeverity: "error",
  validate: (c) => (c.expect !== undefined && typeof c.expect !== "number" ? "expect 必須是數字" : undefined),
  run({ structure, config, report }) {
    const expect = (config.expect as number | undefined) ?? 100;
    let sum = 0;
    let unknown = false;
    for (const s of structure.sections) {
      if (!s.questions.length && s.paragraphIndex === -1) continue;
      const { total, computed } = sectionScore(s);
      const at = { paragraphIndex: s.paragraphIndex >= 0 ? s.paragraphIndex : s.questions[0]?.startParagraph, excerpt: short(s.title) };
      for (const c of s.scoreConflicts ?? []) report(`「${short(s.title, 20)}」配分說明中英文不一致：${c}`, at);
      if (s.totalScore !== undefined && computed !== undefined && Math.abs(s.totalScore - computed) > 1e-9) {
        const how =
          s.perQuestionScore !== undefined
            ? `每題 ${s.perQuestionScore} 分 × ${s.questions.length} 題 = ${computed} 分`
            : `各題配分加總為 ${computed} 分`;
        report(`「${short(s.title, 20)}」標示共 ${s.totalScore} 分，但${how}`, at);
      }
      if (total === undefined) {
        unknown = true;
        report(`無法判斷「${short(s.title, 20)}」的配分，請在大題標題標示「每題 X 分」或「共 Y 分」`, at);
      } else {
        sum += total;
      }
    }
    if (!structure.sections.length) {
      report("找不到任何大題或題目，無法計算總分");
    } else if (!unknown && Math.abs(sum - expect) > 1e-9) {
      report(`全卷配分加總為 ${sum} 分，應為 ${expect} 分`);
    }
  },
};

const NUMBERING: RuleDefinition = {
  description: "題號需連續、不可重複或跳號；每個大題從 1 開始或接續上一大題",
  defaultSeverity: "error",
  validate: (c) =>
    c.restart !== undefined && !["auto", "always", "never"].includes(String(c.restart))
      ? "restart 必須是 auto、always 或 never"
      : undefined,
  run({ structure, config, report }) {
    const mode = (config.restart as string | undefined) ?? "auto";
    let prevLast = 0;
    for (const s of structure.sections) {
      const qs = s.questions;
      if (!qs.length) continue;
      const first = qs[0];
      const okStarts =
        mode === "always" ? [1] : mode === "never" ? [prevLast + 1] : [1, prevLast + 1];
      if (!okStarts.includes(first.number)) {
        report(
          `「${short(s.title, 20)}」第一題題號為 ${first.number}，應為 ${[...new Set(okStarts)].join(" 或 ")}`,
          { paragraphIndex: first.startParagraph, excerpt: short(first.number + ". " + first.body) },
        );
      }
      for (let i = 1; i < qs.length; i++) {
        const prev = qs[i - 1].number;
        const q = qs[i];
        const at = { paragraphIndex: q.startParagraph, excerpt: short(q.number + ". " + q.body) };
        if (q.number === prev) report(`題號 ${q.number} 重複`, at);
        else if (q.number < prev) report(`題號由 ${prev} 倒退為 ${q.number}`, at);
        else if (q.number !== prev + 1)
          report(`題號由 ${prev} 跳到 ${q.number}，缺少第 ${prev + 1}${q.number - prev > 2 ? `～${q.number - 1}` : ""} 題`, at);
      }
      prevLast = qs[qs.length - 1].number;
    }
  },
};

const FULLWIDTH_LETTER = /[Ａ-Ｅａ-ｅ]/g;
const normLetter = (s: string) =>
  s.replace(FULLWIDTH_LETTER, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).toUpperCase();

/** 找出題目中的選項標記，回傳字母與該選項內容 */
export function extractOptions(body: string): { letter: string; text: string }[] {
  const marks: { letter: string; start: number; end: number }[] = [];
  const paren = /[（(]\s*([A-Ea-eＡ-Ｅａ-ｅ])\s*[)）]/g;
  const lineStart = /(^|\n)\s*([A-E])\s*[.．、]/g;
  for (const m of body.matchAll(paren)) marks.push({ letter: normLetter(m[1]), start: m.index!, end: m.index! + m[0].length });
  for (const m of body.matchAll(lineStart)) {
    const start = m.index! + m[1].length;
    if (!marks.some((k) => k.start <= start && start < k.end))
      marks.push({ letter: m[2], start, end: m.index! + m[0].length });
  }
  marks.sort((a, b) => a.start - b.start);
  return marks.map((k, i) => ({
    letter: k.letter,
    text: body.slice(k.end, marks[i + 1]?.start ?? body.length).trim(),
  }));
}

const CHOICE_OPTIONS: RuleDefinition = {
  description: "選擇題需有正確數量的選項（A、B、C、D…依序排列），且選項內容不可重複",
  defaultSeverity: "error",
  validate: (c) => {
    const n = c.count;
    const ok = n === undefined || typeof n === "number" || (Array.isArray(n) && n.every((x) => typeof x === "number"));
    return ok ? undefined : "count 必須是數字或數字清單，例如 4 或 [4, 5]";
  },
  run({ structure, config, report }) {
    const counts = config.count === undefined ? [4] : ([] as number[]).concat(config.count as number | number[]);
    for (const s of structure.sections) {
      if (s.kind !== "choice") continue;
      for (const q of s.questions) {
        const at = { paragraphIndex: q.startParagraph, excerpt: short(q.number + ". " + q.body) };
        const opts = extractOptions(q.body);
        const letters = opts.map((o) => o.letter);
        if (!opts.length) {
          report(`第 ${q.number} 題找不到選項`, at);
          continue;
        }
        if (!counts.includes(opts.length)) {
          report(`第 ${q.number} 題有 ${opts.length} 個選項（${letters.join("")}），應為 ${counts.join(" 或 ")} 個`, at);
        }
        const expected = letters.map((_, i) => String.fromCharCode(65 + i)).join("");
        if (letters.join("") !== expected) {
          report(`第 ${q.number} 題選項順序為 ${letters.join("")}，應為 ${expected}`, at);
        }
        const seen = new Map<string, string>();
        for (const o of opts) {
          const key = o.text.replace(/\s+/g, "");
          if (!key) {
            report(`第 ${q.number} 題選項 (${o.letter}) 沒有內容`, at);
            continue;
          }
          const dup = seen.get(key);
          if (dup) report(`第 ${q.number} 題選項 (${dup}) 與 (${o.letter}) 內容相同：「${short(o.text, 15)}」`, at);
          else seen.set(key, o.letter);
        }
      }
    }
  },
};

const ANSWER_LEFTOVER: RuleDefinition = {
  description: "題號前的作答括號必須是空的（避免把答案留在試卷上）",
  defaultSeverity: "error",
  run({ structure, report }) {
    for (const s of structure.sections) {
      for (const q of s.questions) {
        if (q.answerMark) {
          report(`第 ${q.number} 題的作答括號內有「${q.answerMark}」，疑似忘記刪除答案`, {
            paragraphIndex: q.startParagraph,
            excerpt: short(`(${q.answerMark})${q.number}. ${q.body}`),
          });
        }
      }
    }
  },
};

const PAIRS: [string, string][] = [
  ["(", ")"], ["（", "）"], ["[", "]"], ["【", "】"], ["「", "」"], ["『", "』"], ["《", "》"], ["〈", "〉"],
];

const BRACKET_BALANCE: RuleDefinition = {
  description: "每個段落的括號、引號需成對出現",
  defaultSeverity: "warning",
  run({ doc, report }) {
    for (const p of doc.paragraphs) {
      for (const [open, close] of PAIRS) {
        const a = p.text.split(open).length - 1;
        const b = p.text.split(close).length - 1;
        if (a !== b) {
          report(`「${open}」出現 ${a} 次，「${close}」出現 ${b} 次`, { paragraphIndex: p.index, excerpt: short(p.text, 50) });
        }
      }
    }
  },
};

const REGEX: RuleDefinition = {
  description: "自訂正規表示式：mode: forbid 找到就報告；mode: require 找不到才報告",
  defaultSeverity: "warning",
  validate: (c) => {
    const mode = c.mode ?? "forbid";
    if (mode !== "forbid" && mode !== "require") return "mode 必須是 forbid 或 require";
    if (c.pattern === undefined && c.patterns === undefined) return "需要 pattern 或 patterns";
    if (c.patterns !== undefined && !Array.isArray(c.patterns)) return "patterns 必須是清單";
    for (const p of (c.patterns as unknown[] | undefined) ?? [c.pattern]) {
      const err = validateRegex({ ...c, pattern: p }, "pattern");
      if (err) return err;
    }
  },
  run({ doc, config, report }) {
    const mode = (config.mode as string | undefined) ?? "forbid";
    const patterns = (config.patterns as string[] | undefined) ?? [config.pattern as string];
    const within = config.within as number | undefined;
    const maxReports = (config.maxReports as number | undefined) ?? 20;

    let paragraphs = doc.paragraphs.filter((p) => p.text.trim());
    if (within !== undefined) paragraphs = paragraphs.slice(0, within);

    if (mode === "require") {
      for (const pattern of patterns) {
        const re = compileRegex(pattern, config.flags);
        if (!paragraphs.some((p) => (re.lastIndex = 0, re.test(p.text)))) {
          report(`${within ? `前 ${within} 段` : "全文"}找不到「${pattern}」`);
        }
      }
      return;
    }

    let n = 0;
    for (const p of paragraphs) {
      for (const pattern of patterns) {
        for (const m of p.text.matchAll(compileRegex(pattern, config.flags))) {
          if (n++ >= maxReports) {
            report(`符合的地方超過 ${maxReports} 處，其餘省略`);
            return;
          }
          report(`找到「${m[0]}」`, {
            paragraphIndex: p.index,
            excerpt: excerptAround(p.text, m.index!, m.index! + m[0].length),
          });
          if (!m[0].length) break; // 避免空字串比對造成無窮迴圈
        }
      }
    }
  },
};


const isStringList = (v: unknown) => v === undefined || (Array.isArray(v) && v.every((x) => typeof x === "string"));

const SUBJECT_NAME: RuleDefinition = {
  description: "檢查試卷開頭的中英文科目名稱是否寫錯或拼錯（G1–G12 常見科目）",
  defaultSeverity: "warning",
  validate: (c) => {
    for (const k of ["extraSubjects", "ignore"]) if (!isStringList(c[k])) return `${k} 必須是文字清單`;
    if (c.within !== undefined && typeof c.within !== "number") return "within 必須是數字";
  },
  run({ doc, config, report }) {
    const within = (config.within as number | undefined) ?? 5;
    const extra = (config.extraSubjects as string[] | undefined) ?? [];
    const ignore = new Set(((config.ignore as string[] | undefined) ?? []).map((s) => s.toLowerCase()));
    const zh = [...CHINESE_SUBJECTS, ...extra.filter((s) => /[\u4e00-\u9fff]/.test(s))];
    const en = [...ENGLISH_SUBJECTS, ...extra.filter((s) => !/[\u4e00-\u9fff]/.test(s))];

    let foundAny = false;
    const header = doc.paragraphs.filter((p) => p.text.trim()).slice(0, within);
    for (const p of header) {
      for (const f of [...checkChineseSubjects(p.text, zh), ...checkEnglishSubjects(p.text, en)]) {
        if (ignore.has(f.found.toLowerCase())) continue;
        foundAny = true;
        if (f.kind === "typo") {
          report(`科目名稱「${f.found}」疑似${/[A-Za-z]/.test(f.found) ? "拼錯" : "寫錯"}，應為「${f.suggestion}」`, {
            paragraphIndex: p.index,
            excerpt: short(p.text, 50),
          });
        }
      }
    }
    if (!foundAny && config.requireSubject !== false) {
      report(`前 ${within} 段找不到科目名稱（例如「數學科」「Grade 7 Mathematics」）`);
    }
  },
};

export const builtinRules: Record<string, RuleDefinition> = {
  "score-sum": SCORE_SUM,
  numbering: NUMBERING,
  "choice-options": CHOICE_OPTIONS,
  "answer-leftover": ANSWER_LEFTOVER,
  "bracket-balance": BRACKET_BALANCE,
  regex: REGEX,
  "subject-name": SUBJECT_NAME,
  "exam-title": EXAM_TITLE,
};
