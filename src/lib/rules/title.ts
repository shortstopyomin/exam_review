import type { RuleDefinition, RuleContext } from "./engine";
import type { Paragraph } from "../types";
import { editDistance } from "./subjects";

/**
 * exam-title：檢查試卷標題（前幾段）的
 * 學校名稱、學年度、學期、年級、考試類型，
 * 並檢查中英文標題彼此是否一致。
 */

export type ExamType = "midterm" | "final" | "quiz" | "mock";

const EXAM_LABEL: Record<ExamType, { zh: string; en: string }> = {
  midterm: { zh: "期中", en: "Midterm" },
  final: { zh: "期末", en: "Final" },
  quiz: { zh: "小考", en: "Quiz" },
  mock: { zh: "模擬考", en: "Mock" },
};

const ZH_EXAM: [RegExp, ExamType][] = [
  [/期中/g, "midterm"],
  [/期末/g, "final"],
  [/模擬考?/g, "mock"],
  [/小考|隨堂測驗/g, "quiz"],
];
const EN_EXAM: [RegExp, ExamType][] = [
  [/\bmid-?\s?terms?\b|\bmid[-\s]?semester\b/gi, "midterm"],
  [/\bfinals?\b|\bend[-\s]of[-\s](?:the[-\s])?(?:semester|term)\b/gi, "final"],
  [/\bmock\b/gi, "mock"],
  [/\bquiz(?:zes)?\b/gi, "quiz"],
];

const ZH_DIGITS: Record<string, number> = {
  一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10,
  "１": 1, "２": 2,
};

/** 「十二」「10」「２」→ 數字 */
function toNumber(s: string): number | undefined {
  if (/^\d+$/.test(s)) return Number(s);
  if (s in ZH_DIGITS) return ZH_DIGITS[s];
  const m = s.match(/^([一二三四五六七八九]?)十([一二三四五六七八九]?)$/);
  if (m) return (m[1] ? ZH_DIGITS[m[1]] : 1) * 10 + (m[2] ? ZH_DIGITS[m[2]] : 0);
  return undefined;
}

interface Found<T> {
  value: T;
  text: string;
  lang: "zh" | "en";
  paragraph: Paragraph;
}

function findAll<T>(paragraphs: Paragraph[], lang: "zh" | "en", re: RegExp, map: (m: RegExpMatchArray) => T | undefined) {
  const out: Found<T>[] = [];
  for (const p of paragraphs) {
    for (const m of p.text.matchAll(re)) {
      const value = map(m);
      if (value !== undefined) out.push({ value, text: m[0].trim(), lang, paragraph: p });
    }
  }
  return out;
}

// ── 學年度 ────────────────────────────────────────

/** 依日期推算學年度與學期：8 月～隔年 1 月為第 1 學期，2～7 月為第 2 學期 */
export function currentTerm(today: Date, secondSemesterStartMonth = 2, firstSemesterStartMonth = 8) {
  const y = today.getFullYear();
  const m = today.getMonth() + 1;
  const startYear = m >= firstSemesterStartMonth ? y : y - 1;
  const semester = m >= firstSemesterStartMonth || m < secondSemesterStartMonth ? 1 : 2;
  return { startYear, semester };
}

/** 「2026-2027」「2026–27」「2026/2027」；中英文標題通用 */
const YEAR_RANGE_RE = /(20\d{2})\s*[-–—~～/]\s*(20\d{2}|\d{2})(?!\d)/g;
/** 民國學年度「115 學年度」 */
const ROC_YEAR_RE = /(?<!\d)(1\d{2})\s*學年度?/g;

// ── 學期 ──────────────────────────────────────────

const ZH_SEMESTER_RE = /第\s*([一二12１２])\s*學期|([上下])學期/g;
/**
 * Semester 1 / Semester I / Semester One / 1st Semester / First Semester / Fall Semester
 * Sem 1 / Sem. 1 / Term 1 / 1st Term / First Term / S1
 * （「Midterm」「Mid Term 2」的 term 不會被當成學期）
 */
const EN_SEMESTER_RE =
  /\b(?:semester|sem\.?|(?<!mid[-\s]?)term)\s*(1|2|I{1,2}|one|two)\b|\b(first|second|1st|2nd)\s+(?:semester|term)\b|\b(fall|spring)\s+semester\b|\bS([12])\b/gi;

function enSemester(m: RegExpMatchArray): number | undefined {
  const v = (m[1] ?? m[2] ?? m[3] ?? m[4] ?? "").toLowerCase();
  if (["1", "i", "one", "first", "1st", "fall"].includes(v)) return 1;
  if (["2", "ii", "two", "second", "2nd", "spring"].includes(v)) return 2;
}

// ── 年級 ──────────────────────────────────────────

const ZH_GRADE_RE = /(\d{1,2}|[一二三四五六七八九十]{1,3})\s*年級/g;
const EN_GRADE_RE = /\bgrade\s*(\d{1,2})\b|\bG(\d{1,2})\b|\b(\d{1,2})(?:st|nd|rd|th)\s+grade\b/gi;

// ── 學校名稱 ──────────────────────────────────────

/**
 * 在標題中找最接近 target 的片段。
 * 回傳 distance = 0 代表完全正確。
 */
export function closestSubstring(text: string, target: string, ignoreCase = false) {
  const t = ignoreCase ? target.toLowerCase() : target;
  const src = ignoreCase ? text.toLowerCase() : text;
  if (src.includes(t)) {
    const i = src.indexOf(t);
    return { found: text.slice(i, i + target.length), distance: 0 };
  }
  let best: { found: string; distance: number } | undefined;
  const chars = [...src];
  const orig = [...text];
  const L = [...t].length;
  const slack = Math.max(2, Math.ceil(L * 0.2));
  for (let len = Math.max(1, L - slack); len <= L + slack; len++) {
    for (let i = 0; i + len <= chars.length; i++) {
      const d = editDistance(chars.slice(i, i + len).join(""), t);
      if (!best || d < best.distance || (d === best.distance && Math.abs(len - L) < Math.abs([...best.found].length - L))) {
        best = { found: orig.slice(i, i + len).join("").trim(), distance: d };
      }
    }
  }
  return best ?? { found: "", distance: Infinity };
}

function checkSchool(ctx: RuleContext, paragraphs: Paragraph[], name: string, lang: "zh" | "en") {
  const ignoreCase = lang === "en";
  let best: { found: string; distance: number; p: Paragraph } | undefined;
  for (const p of paragraphs) {
    const r = closestSubstring(p.text, name, ignoreCase);
    if (!best || r.distance < best.distance) best = { ...r, p };
  }
  const label = lang === "zh" ? "中文學校名稱" : "英文學校名稱";
  const tolerance = Math.max(1, Math.floor([...name].length * 0.25));
  if (!best || best.distance > tolerance) {
    ctx.report(`標題中找不到${label}「${name}」`);
    return;
  }
  const at = { paragraphIndex: best.p.index, excerpt: best.found };
  if (best.distance > 0) {
    ctx.report(`${label}「${best.found}」疑似寫錯，應為「${name}」`, at);
  } else if (lang === "en" && best.found !== name) {
    ctx.report(`${label}大小寫不正確：「${best.found}」應為「${name}」`, at);
  }
}

// ── 規則本體 ──────────────────────────────────────

function parseToday(v: unknown): Date {
  if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}$/.test(v)) {
    const [y, m, d] = v.split("-").map(Number);
    return new Date(y, m - 1, d);
  }
  return new Date();
}

function parseYearSetting(v: unknown, today: Date): number | undefined {
  if (v === undefined || v === "auto") return currentTerm(today).startYear;
  if (v === false || v === "off") return undefined;
  const m = String(v).match(/^(20\d{2})\s*[-–/]\s*(20\d{2})$/);
  return m ? Number(m[1]) : undefined;
}

const EXAM_TYPES = Object.keys(EXAM_LABEL);

export const EXAM_TITLE: RuleDefinition = {
  description: "檢查試卷標題的學校名稱、學年度、學期、年級、考試類型，以及中英文標題是否一致",
  defaultSeverity: "error",
  validate(c) {
    for (const k of ["schoolZh", "schoolEn"]) {
      if (c[k] !== undefined && typeof c[k] !== "string") return `${k} 必須是文字`;
    }
    if (c.examType !== undefined && !EXAM_TYPES.includes(String(c.examType))) {
      return `examType 必須是 ${EXAM_TYPES.join("、")} 其中之一`;
    }
    if (c.semester !== undefined && !["auto", 1, 2, false, "off"].includes(c.semester as never)) {
      return "semester 必須是 1、2、\"auto\" 或 \"off\"";
    }
    if (
      c.academicYear !== undefined &&
      !["auto", "off", false].includes(c.academicYear as never) &&
      !/^20\d{2}\s*[-–/]\s*20\d{2}$/.test(String(c.academicYear))
    ) {
      return "academicYear 必須是 \"auto\"、\"off\" 或像 \"2026-2027\" 的格式";
    }
    if (c.grade !== undefined && typeof c.grade !== "number") return "grade 必須是數字";
    if (c.today !== undefined && !/^\d{4}-\d{2}-\d{2}$/.test(String(c.today))) return "today 必須是 YYYY-MM-DD";
  },
  run(ctx) {
    const { doc, config, report } = ctx;
    const within = (config.within as number | undefined) ?? 5;
    const title = doc.paragraphs.filter((p) => p.text.trim()).slice(0, within);
    const today = parseToday(config.today);
    const at = (f: Found<unknown>) => ({ paragraphIndex: f.paragraph.index, excerpt: f.paragraph.text.trim() });
    // 是否有中文標題／英文標題（用來判斷「某一種語言漏寫」）
    const hasZh = title.some((p) => /[一-鿿]{4,}/.test(p.text) && /(學期|學年|年級|考|評量|測驗)/.test(p.text));
    const hasEn = title.some((p) => /\b(exam|examination|test|quiz|assessment|semester|grade)\b/i.test(p.text));

    // 學校名稱
    if (config.schoolZh) checkSchool(ctx, title, String(config.schoolZh), "zh");
    if (config.schoolEn) checkSchool(ctx, title, String(config.schoolEn), "en");

    // 學年度
    const expectedYear = parseYearSetting(config.academicYear, today);
    const years = findAll(title, "en", YEAR_RANGE_RE, (m) => {
      const start = Number(m[1]);
      const end = m[2].length === 2 ? Math.floor(start / 100) * 100 + Number(m[2]) : Number(m[2]);
      return { start, end };
    });
    for (const y of years) {
      if (y.value.end !== y.value.start + 1) {
        report(`學年度「${y.text}」前後年份不連續`, at(y));
      } else if (expectedYear !== undefined && y.value.start !== expectedYear) {
        report(`學年度「${y.text}」應為「${expectedYear}-${expectedYear + 1}」`, at(y));
      }
    }
    const rocYears = findAll(title, "zh", ROC_YEAR_RE, (m) => Number(m[1]));
    for (const y of rocYears) {
      if (expectedYear !== undefined && y.value + 1911 !== expectedYear) {
        report(`「${y.text}」應為「${expectedYear - 1911} 學年度」`, at(y));
      }
    }
    if (expectedYear !== undefined && !years.length && !rocYears.length) {
      report(`標題中找不到學年度（例如「${expectedYear}-${expectedYear + 1}」）`);
    }

    // 學期
    const semSetting = config.semester ?? "auto";
    const expectedSem =
      semSetting === "auto" ? currentTerm(today).semester : typeof semSetting === "number" ? semSetting : undefined;
    const sems = [
      ...findAll(title, "zh", ZH_SEMESTER_RE, (m) => (m[2] ? (m[2] === "上" ? 1 : 2) : toNumber(m[1]))),
      ...findAll(title, "en", EN_SEMESTER_RE, enSemester),
    ];
    for (const s of sems) {
      if (expectedSem !== undefined && s.value !== expectedSem) {
        report(`「${s.text}」應為第 ${expectedSem} 學期`, at(s));
      }
    }
    reportInconsistent(ctx, sems, "學期", (v) => `第 ${v} 學期`, expectedSem);
    if (expectedSem !== undefined) {
      if (!sems.length) report(`標題中找不到學期（例如「第 ${expectedSem} 學期」）`);
      else if (hasZh && !sems.some((x) => x.lang === "zh")) report(`中文標題找不到學期（例如「第 ${expectedSem} 學期」）`);
      else if (hasEn && !sems.some((x) => x.lang === "en")) report(`英文標題找不到學期（例如「Semester ${expectedSem}」）`);
    }

    // 年級
    const grades = [
      ...findAll(title, "zh", ZH_GRADE_RE, (m) => toNumber(m[1])),
      ...findAll(title, "en", EN_GRADE_RE, (m) => Number(m[1] ?? m[2] ?? m[3])),
    ];
    const expectedGrade = config.grade as number | undefined;
    for (const g of grades) {
      if (expectedGrade !== undefined && g.value !== expectedGrade) {
        report(`「${g.text}」應為 ${expectedGrade} 年級`, at(g));
      }
    }
    reportInconsistent(ctx, grades, "年級", (v) => `${v} 年級`, expectedGrade);

    // 考試類型
    const exams = [
      ...ZH_EXAM.flatMap(([re, type]) => findAll(title, "zh", re, () => type)),
      ...EN_EXAM.flatMap(([re, type]) => findAll(title, "en", re, () => type)),
    ];
    const expectedExam = config.examType as ExamType | undefined;
    if (expectedExam) {
      for (const e of exams) {
        if (e.value !== expectedExam) {
          const want = e.lang === "zh" ? EXAM_LABEL[expectedExam].zh : EXAM_LABEL[expectedExam].en;
          report(`考試類型「${e.text}」應為「${want}」`, at(e));
        }
      }
      if (!exams.some((e) => e.lang === "zh") && hasZh) {
        report(`中文標題找不到考試類型「${EXAM_LABEL[expectedExam].zh}」`);
      }
      if (!exams.some((e) => e.lang === "en") && hasEn) {
        report(`英文標題找不到考試類型「${EXAM_LABEL[expectedExam].en}」`);
      }
    }
    reportInconsistent(ctx, exams, "考試類型", (v) => EXAM_LABEL[v].zh, expectedExam);
  },
};

/**
 * 中英文標題的值不一致時回報。
 * 若已經有「應為」設定，個別錯誤已經回報過，就不再重複回報不一致。
 */
function reportInconsistent<T>(
  ctx: RuleContext,
  found: Found<T>[],
  label: string,
  show: (v: T) => string,
  expected: T | undefined,
) {
  if (expected !== undefined) return;
  const zh = found.filter((f) => f.lang === "zh");
  const en = found.filter((f) => f.lang === "en");
  const values = new Set(found.map((f) => f.value));
  if (values.size <= 1) return;
  const describe = (list: Found<T>[]) => [...new Set(list.map((f) => `「${f.text}」`))].join("、") || "（無）";
  const first = found.find((f) => f.value !== found[0].value)!;
  ctx.report(
    `中英文標題的${label}不一致：中文 ${describe(zh)}，英文 ${describe(en)}（${[...values].map(show).join(" vs ")}）`,
    { paragraphIndex: first.paragraph.index, excerpt: first.paragraph.text.trim() },
  );
}
