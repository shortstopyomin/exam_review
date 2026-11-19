/**
 * G1–G12 常見科目名稱（臺灣 108 課綱 + 常見英文名稱）。
 * 規則檔可以用 extraSubjects 補充學校自訂的科目名稱。
 */
export const CHINESE_SUBJECTS = [
  // 語文
  "國語", "國語文", "國文", "華語", "語文", "作文", "書法", "閱讀",
  "英語", "英語文", "英文", "英語聽力",
  "本土語", "本土語文", "閩南語", "臺灣台語", "台灣台語", "臺語", "台語", "客語", "客家語",
  "原住民族語", "原住民族語文", "新住民語", "新住民語文", "臺灣手語",
  "第二外國語文", "日語", "日文", "法語", "法文", "德語", "德文", "韓語", "韓文", "西班牙語", "西班牙文",
  // 數學
  "數學", "數學甲", "數學乙", "數學A", "數學B",
  // 自然
  "自然", "自然科學", "自然與生活科技", "理化", "物理", "化學", "生物", "地球科學",
  // 社會
  "社會", "歷史", "地理", "公民", "公民與社會",
  // 科技
  "科技", "資訊", "資訊科技", "生活科技", "電腦",
  // 藝術
  "藝術", "藝術與人文", "音樂", "美術", "視覺藝術", "表演藝術", "藝術生活",
  // 綜合活動
  "綜合", "綜合活動", "家政", "童軍", "輔導", "生涯規劃", "生命教育",
  // 健康與體育
  "健康", "體育", "健康與體育", "健康與護理", "護理",
  // 其他
  "生活", "全民國防教育", "國防", "法律與生活",
];

/**
 * 常見簡稱：只用來認定「寫對了」，不會拿來當更正建議。
 * 否則它們和正式名稱只差一個字，「地裡」可能被建議改成「地科」。
 */
export const CHINESE_ABBREVIATIONS = [
  "國寫", "英聽", "數甲", "數乙", "數A", "數B", "地科", "地球", "公社",
  "資科", "生科", "視藝", "表藝", "健體", "健教",
];

export const ENGLISH_SUBJECTS = [
  "Chinese", "Mandarin", "Chinese Language", "Chinese Literature",
  "English", "English Language", "English Literature", "Language Arts", "English Language Arts", "Literature",
  "Japanese", "French", "German", "Spanish", "Korean",
  "Mathematics", "Math", "Algebra", "Geometry", "Calculus", "Statistics", "Precalculus", "Trigonometry",
  "Science", "Natural Science", "Physics", "Chemistry", "Biology", "Earth Science", "Environmental Science",
  "Physical Science", "Life Science", "General Science",
  "Social Studies", "Social Science", "History", "World History", "Geography", "Civics",
  "Civics and Society", "Economics", "Government", "Psychology", "Philosophy",
  "Technology", "Information Technology", "Computer Science", "Living Technology", "Computing",
  "Music", "Art", "Arts", "Visual Arts", "Performing Arts", "Fine Arts", "Drama",
  "Physical Education", "Health", "Health Education", "Health and Physical Education",
  "Home Economics", "Integrative Activities", "Life Curriculum", "Life Education",
];

/** Damerau–Levenshtein（OSA）距離：相鄰字元對調也算一次錯誤，例如 Histroy */
export function editDistance(a: string, b: string): number {
  const A = [...a];
  const B = [...b];
  const d: number[][] = Array.from({ length: A.length + 1 }, (_, i) =>
    Array.from({ length: B.length + 1 }, (_, j) => (i === 0 ? j : j === 0 ? i : 0)),
  );
  for (let i = 1; i <= A.length; i++) {
    for (let j = 1; j <= B.length; j++) {
      const cost = A[i - 1] === B[j - 1] ? 0 : 1;
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost);
      if (i > 1 && j > 1 && A[i - 1] === B[j - 2] && A[i - 2] === B[j - 1]) {
        d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
      }
    }
  }
  return d[A.length][B.length];
}

export interface SubjectFinding {
  kind: "ok" | "typo";
  found: string;
  /** 建議的正確寫法（typo 時） */
  suggestion?: string;
}

// ── 中文 ─────────────────────────────────────────

const CJK = "\\u4e00-\\u9fff";

/** 學校名稱、學年度、年級等不是科目的字樣，先移除以免誤判 */
const NOISE_RE = new RegExp(
  [
    `[${CJK}○◯〇Ｏ]*(?:國民中學|國民小學|高級中學|高級中等學校|完全中學|實驗學校|高中|國中|國小|中學|小學|學校)`,
    `\\d*\\s*學年度?`,
    `第[一二三四五六七八九十\\d]+學期|[上下]學期`,
    `[一二三四五六七八九十]+年級|[一二三四五六七八九十\\d]+年[甲乙丙丁戊己一二三四五六七八九十\\d]*班|[高國][一二三](?![${CJK}])`,
  ].join("|"),
  "g",
);

/** 「科目：數學」 */
const LABELLED_RE = new RegExp(`科目\\s*[:：]\\s*([${CJK}A-Za-z]{1,10})`, "g");

/** 「數學科 第一次段考」「自然科學期末考」「英文試卷」 */
const BEFORE_EXAM_RE = new RegExp(
  `([${CJK}]{2,8}?[AB甲乙]?)(?:科|領域)?(?=\\s*(?:第[一二三四五六七八九十\\d]+次|期中|期末|段考|定期|評量|試卷|試題|測驗|考試|小考|複習|練習|習作|學習單|作業|講義|$))`,
  "g",
);

function zhThreshold(len: number) {
  return len >= 6 ? 2 : 1;
}

export function checkChineseSubjects(
  text: string,
  subjects: string[],
  abbreviations: string[] = CHINESE_ABBREVIATIONS,
): SubjectFinding[] {
  const known = [...subjects, ...abbreviations];
  const cleaned = text.replace(NOISE_RE, " ");
  const candidates = new Set<string>();
  for (const m of cleaned.matchAll(LABELLED_RE)) candidates.add(m[1]);
  for (const m of cleaned.matchAll(BEFORE_EXAM_RE)) candidates.add(m[1]);
  // 標題中獨立的詞剛好是科目名稱（例如「四年級 數學 習作」）也算
  for (const token of cleaned.split(/[\s，,、。：:（）()【】「」\-—]+/)) {
    if (known.includes(token)) candidates.add(token);
  }

  const out: SubjectFinding[] = [];
  for (const c of candidates) {
    // 候選字串以正確科目名稱結尾（例如「國中數學」）→ 正確。
    // 用「結尾」而不是「包含」：「自然科厂」含有「自然」，但其實是「自然科學」寫錯。
    const exact = known.filter((s) => c.endsWith(s)).sort((a, b) => b.length - a.length)[0];
    if (exact) {
      out.push({ kind: "ok", found: exact });
      continue;
    }
    // 比對候選字串的結尾（科目名稱通常緊接在「科」「段考」之前）；只用正式名稱當建議
    let best: { s: string; suffix: string; d: number } | undefined;
    for (const s of subjects) {
      for (const len of [s.length - 1, s.length, s.length + 1]) {
        if (len < 2 || len > c.length) continue;
        const suffix = c.slice(-len);
        const d = editDistance(suffix, s);
        if (d <= zhThreshold(s.length) && d > 0 && (!best || d < best.d || (d === best.d && s.length > best.s.length))) {
          best = { s, suffix, d };
        }
      }
    }
    if (best) out.push({ kind: "typo", found: best.suffix, suggestion: best.s });
  }
  return out;
}

// ── 英文 ─────────────────────────────────────────

function enThreshold(len: number) {
  if (len <= 4) return 0; // 太短的字（Art、Math）不做模糊比對，避免誤報
  return len <= 8 ? 1 : 2;
}

export function checkEnglishSubjects(text: string, subjects: string[]): SubjectFinding[] {
  const words = [...text.matchAll(/[A-Za-z]+/g)].map((m) => m[0]);
  const lower = subjects.map((s) => s.toLowerCase());
  const covered = new Set<number>(); // 已經判定過的字（避免同一個錯字被重複回報）
  const out: SubjectFinding[] = [];
  const maxN = Math.max(...subjects.map((s) => s.split(" ").length));

  const phraseAt = (i: number, n: number) => words.slice(i, i + n).join(" ");
  const mark = (i: number, n: number) => {
    for (let k = i; k < i + n; k++) covered.add(k);
  };
  const isCovered = (i: number, n: number) => {
    for (let k = i; k < i + n; k++) if (covered.has(k)) return true;
    return false;
  };

  const exactPass = (n: number) => {
    for (let i = 0; i + n <= words.length; i++) {
      if (isCovered(i, n)) continue;
      const pl = phraseAt(i, n).toLowerCase();
      // 允許複數 s，例如 Sciences
      const idx = lower.findIndex((s) => s === pl || s + "s" === pl);
      if (idx >= 0) {
        out.push({ kind: "ok", found: subjects[idx] });
        mark(i, n);
      }
    }
  };
  const fuzzyPass = (n: number) => {
    for (let i = 0; i + n <= words.length; i++) {
      if (isCovered(i, n)) continue;
      const p = phraseAt(i, n);
      const pl = p.toLowerCase();
      let best: { s: string; d: number } | undefined;
      lower.forEach((s, k) => {
        // 首字母必須相同以降低誤報
        if (s.split(" ").length !== n || s[0] !== pl[0]) return;
        const d = editDistance(pl, s);
        if (d > 0 && d <= enThreshold(s.length) && (!best || d < best.d)) best = { s: subjects[k], d };
      });
      if (best) {
        out.push({ kind: "typo", found: p, suggestion: best.s });
        mark(i, n);
      }
    }
  };

  // 多字片語（Earth Science）要先於單字比對，
  // 否則「Earht Science」中的 Science 會先被判為正確，錯字 Earht 就漏掉了。
  for (let n = maxN; n >= 2; n--) exactPass(n);
  for (let n = maxN; n >= 2; n--) fuzzyPass(n);
  exactPass(1);
  fuzzyPass(1);
  return out;
}
