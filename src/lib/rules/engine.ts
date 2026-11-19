import { parse as parseToml } from "smol-toml";
import type { ExamStructure, Issue, ParsedDocument, RuleConfig, RuleSet, Severity } from "../types";
import { analyzeStructure } from "../structure";
import { builtinRules } from "./builtin";

export interface RuleContext {
  doc: ParsedDocument;
  structure: ExamStructure;
  config: RuleConfig;
  severity: Severity;
  /** 建立一個 Issue；message 可被規則檔的 message 覆寫 */
  report(message: string, at?: { paragraphIndex?: number; excerpt?: string }): void;
}

export interface RuleDefinition {
  description: string;
  defaultSeverity: Severity;
  /** 驗證設定，回傳錯誤訊息；沒問題回傳 undefined */
  validate?(config: RuleConfig): string | undefined;
  run(ctx: RuleContext): void;
}

export class RuleSetError extends Error {}

const SEVERITIES: Severity[] = ["error", "warning", "info"];

export function parseRuleSet(source: string): RuleSet {
  let data: unknown;
  try {
    data = parseToml(source);
  } catch (e) {
    throw new RuleSetError(`規則檔 TOML 格式錯誤：${(e as Error).message}`);
  }
  if (!data || typeof data !== "object" || !Array.isArray((data as RuleSet).rules)) {
    throw new RuleSetError("規則檔必須包含至少一個 [[rules]] 區塊。");
  }
  const set = data as RuleSet;
  const seen = new Set<string>();
  set.rules.forEach((r, i) => {
    const where = `第 ${i + 1} 條規則`;
    if (!r || typeof r !== "object") throw new RuleSetError(`${where}格式錯誤。`);
    if (!r.id) throw new RuleSetError(`${where}缺少 id。`);
    if (seen.has(r.id)) throw new RuleSetError(`規則 id「${r.id}」重複。`);
    seen.add(r.id);
    const def = builtinRules[r.type];
    if (!def) {
      throw new RuleSetError(
        `規則「${r.id}」的 type「${r.type}」不存在。可用的 type：${Object.keys(builtinRules).join("、")}`,
      );
    }
    if (r.severity && !SEVERITIES.includes(r.severity)) {
      throw new RuleSetError(`規則「${r.id}」的 severity 必須是 error、warning 或 info。`);
    }
    const err = def.validate?.(r);
    if (err) throw new RuleSetError(`規則「${r.id}」設定錯誤：${err}`);
  });
  return { name: set.name || "未命名規則集", rules: set.rules };
}

export interface LintResult {
  structure: ExamStructure;
  issues: Issue[];
  /** 各規則的問題數，0 代表通過 */
  ruleSummary: { id: string; type: string; severity: Severity; count: number; description: string }[];
  passed: boolean;
}

export function lint(doc: ParsedDocument, ruleSet: RuleSet): LintResult {
  const structure = analyzeStructure(doc.paragraphs);
  const issues: Issue[] = [];
  const ruleSummary: LintResult["ruleSummary"] = [];

  for (const config of ruleSet.rules) {
    if (config.enabled === false) continue;
    const def = builtinRules[config.type];
    const severity = config.severity ?? def.defaultSeverity;
    let count = 0;
    const ctx: RuleContext = {
      doc,
      structure,
      config,
      severity,
      report(message, at) {
        count++;
        issues.push({
          ruleId: config.id,
          severity,
          message: config.message ? `${config.message}（${message}）` : message,
          ...at,
        });
      },
    };
    try {
      def.run(ctx);
    } catch (e) {
      ctx.report(`規則執行失敗：${(e as Error).message}`);
    }
    ruleSummary.push({ id: config.id, type: config.type, severity, count, description: def.description });
  }

  const order: Record<Severity, number> = { error: 0, warning: 1, info: 2 };
  issues.sort(
    (a, b) =>
      (a.paragraphIndex ?? -1) - (b.paragraphIndex ?? -1) || order[a.severity] - order[b.severity],
  );
  return { structure, issues, ruleSummary, passed: !issues.some((i) => i.severity === "error") };
}
