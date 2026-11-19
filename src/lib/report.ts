import type { LintResult } from "./rules/engine";
import type { ParsedDocument } from "./types";
import { sectionScore } from "./structure";

const LABEL = { error: "錯誤", warning: "警告", info: "提醒" } as const;

/** 產生 Markdown 格式的審題報告 */
export function buildReport(doc: ParsedDocument, result: LintResult, ruleSetName: string): string {
  const lines: string[] = [];
  const date = new Date().toLocaleString("zh-TW");
  lines.push(`# 審題報告：${doc.fileName}`, "", `- 檢查時間：${date}`, `- 規則集：${ruleSetName}`);
  lines.push(`- 結果：**${result.passed ? "通過" : "未通過"}**`, "");

  lines.push("## 試卷結構", "", "| 大題 | 題數 | 配分 |", "|---|---:|---:|");
  for (const s of result.structure.sections) {
    lines.push(`| ${s.title} | ${s.questions.length} | ${sectionScore(s).total ?? "?"} |`);
  }
  lines.push("");

  lines.push(`## 問題清單（${result.issues.length}）`, "");
  if (!result.issues.length) lines.push("沒有發現任何問題。");
  for (const i of result.issues) {
    const where = i.paragraphIndex !== undefined ? `第 ${i.paragraphIndex + 1} 段` : "全卷";
    lines.push(`- **[${LABEL[i.severity]}]** ${i.message}（${where}，規則 \`${i.ruleId}\`）`);
    if (i.excerpt) lines.push(`  > ${i.excerpt.replace(/\n/g, " ")}`);
  }
  return lines.join("\n") + "\n";
}

export function download(fileName: string, content: string, type = "text/markdown") {
  const url = URL.createObjectURL(new Blob([content], { type: `${type};charset=utf-8` }));
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
