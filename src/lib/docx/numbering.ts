/** 將數字依 Word 的 numFmt 格式化 */
export function formatNumber(n: number, fmt: string): string {
  switch (fmt) {
    case "decimal":
      return String(n);
    case "decimalZero":
      return String(n).padStart(2, "0");
    case "decimalFullWidth":
    case "decimalFullWidth2":
      return String(n).replace(/\d/g, (d) => String.fromCharCode(0xff10 + Number(d)));
    case "decimalEnclosedCircle":
    case "decimalEnclosedCircleChinese":
      return n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n);
    case "upperLetter":
      return letters(n).toUpperCase();
    case "lowerLetter":
      return letters(n);
    case "upperRoman":
      return roman(n);
    case "lowerRoman":
      return roman(n).toLowerCase();
    case "taiwaneseCounting":
    case "taiwaneseCountingThousand":
    case "chineseCounting":
    case "chineseCountingThousand":
    case "ideographDigital":
      return chinese(n);
    case "ideographLegalTraditional":
      return chinese(n, "零壹貳參肆伍陸柒捌玖", "拾");
    case "ideographTraditional":
      return "甲乙丙丁戊己庚辛壬癸"[(n - 1) % 10];
    case "ideographZodiac":
      return "子丑寅卯辰巳午未申酉戌亥"[(n - 1) % 12];
    case "bullet":
    case "none":
      return "";
    default:
      return String(n);
  }
}

function letters(n: number): string {
  // Word: a..z, aa..zz, aaa...
  const ch = String.fromCharCode(97 + ((n - 1) % 26));
  return ch.repeat(Math.floor((n - 1) / 26) + 1);
}

function roman(n: number): string {
  const table: [number, string][] = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
    [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
  ];
  let out = "";
  for (const [v, s] of table) {
    while (n >= v) {
      out += s;
      n -= v;
    }
  }
  return out;
}

function chinese(n: number, digits = "〇一二三四五六七八九", ten = "十"): string {
  if (n < 10) return digits[n];
  if (n < 100) {
    const t = Math.floor(n / 10);
    const o = n % 10;
    return (t === 1 && ten === "十" ? "" : digits[t]) + ten + (o ? digits[o] : "");
  }
  return String(n);
}

export interface LevelDef {
  start: number;
  fmt: string;
  text: string;
}

/**
 * 追蹤 Word 自動編號的計數狀態。
 * Word 的計數器是以 abstractNum 為單位共用的；若 w:num 帶有 startOverride，
 * 第一次遇到該 numId 時會重設計數。
 */
export class NumberingTracker {
  private counters = new Map<string, (number | undefined)[]>();
  private seenNumIds = new Set<string>();

  constructor(
    private abstracts: Map<string, Map<number, LevelDef>>,
    private nums: Map<string, { abstractId: string; overrides: Map<number, number> }>,
  ) {}

  next(numId: string, ilvl: number): string | undefined {
    const num = this.nums.get(numId);
    if (!num) return undefined;
    const levels = this.abstracts.get(num.abstractId);
    const lvl = levels?.get(ilvl);
    if (!levels || !lvl) return undefined;

    let counters = this.counters.get(num.abstractId);
    if (!counters) {
      counters = [];
      this.counters.set(num.abstractId, counters);
    }
    if (!this.seenNumIds.has(numId)) {
      this.seenNumIds.add(numId);
      for (const [level, start] of num.overrides) {
        counters[level] = start - 1;
        counters.length = Math.max(counters.length, level + 1);
        for (let i = level + 1; i < counters.length; i++) counters[i] = undefined;
      }
    }

    const cur = counters[ilvl];
    counters[ilvl] = cur === undefined ? lvl.start : cur + 1;
    for (let i = ilvl + 1; i < counters.length; i++) counters[i] = undefined;

    if (lvl.fmt === "bullet" || lvl.fmt === "none") return undefined;
    return lvl.text.replace(/%(\d)/g, (_, d: string) => {
      const level = Number(d) - 1;
      const def = levels.get(level);
      const value = counters![level] ?? def?.start ?? 1;
      return formatNumber(value, def?.fmt ?? "decimal");
    });
  }
}
