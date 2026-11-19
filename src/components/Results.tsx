"use client";

import { useState } from "react";
import type { LintResult } from "@/lib/rules/engine";
import type { SectionKind, Severity } from "@/lib/types";
import { sectionScore } from "@/lib/structure";
import { SeverityBadge, SEVERITY_LABEL } from "./SeverityBadge";

const KIND_LABEL: Record<SectionKind, string> = {
  choice: "選擇題",
  truefalse: "是非題",
  fill: "填充題",
  short: "非選題",
  other: "其他",
};

interface Props {
  result: LintResult;
  onJump(index: number): void;
  onDownload(): void;
}

export function Results({ result, onJump, onDownload }: Props) {
  const [filter, setFilter] = useState<Severity | "all">("all");
  const counts = { error: 0, warning: 0, info: 0 };
  for (const i of result.issues) counts[i.severity]++;
  const shown = result.issues.filter((i) => filter === "all" || i.severity === filter);
  const total = result.structure.sections.reduce((a, s) => a + (sectionScore(s).total ?? 0), 0);

  return (
    <>
      <section className={`card verdict ${result.passed ? "pass" : "fail"}`}>
        <div>
          <div className="verdict-title">{result.passed ? "✓ 通過檢查" : "✗ 未通過檢查"}</div>
          <div className="small">
            {counts.error} 個錯誤、{counts.warning} 個警告、{counts.info} 個提醒
            {result.passed && counts.warning + counts.info > 0 && "（警告與提醒不影響通過）"}
          </div>
        </div>
        <button onClick={onDownload}>下載報告</button>
      </section>

      <section className="card">
        <h2>試卷結構</h2>
        {result.structure.sections.length === 0 ? (
          <p className="muted">找不到大題或題目。大題標題請用「一、選擇題」格式，題目請用「1.」開頭。</p>
        ) : (
          <div className="table-wrap">
            <table className="structure">
              <thead>
                <tr>
                  <th>大題</th>
                  <th>題型</th>
                  <th>題號</th>
                  <th className="num">題數</th>
                  <th className="num">配分</th>
                </tr>
              </thead>
              <tbody>
                {result.structure.sections.map((s, i) => {
                  const { total: score } = sectionScore(s);
                  const nums = s.questions.map((q) => q.number);
                  return (
                    <tr key={i} onClick={() => s.paragraphIndex >= 0 && onJump(s.paragraphIndex)}>
                      <td>{s.title}</td>
                      <td>{KIND_LABEL[s.kind]}</td>
                      <td className="muted">{nums.length ? `${nums[0]}–${nums[nums.length - 1]}` : "—"}</td>
                      <td className="num">{s.questions.length}</td>
                      <td className="num">{score ?? "?"}</td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={3}>合計</td>
                  <td className="num">{result.structure.sections.reduce((a, s) => a + s.questions.length, 0)}</td>
                  <td className="num">{total}</td>
                </tr>
              </tfoot>
            </table>
          </div>
        )}
      </section>

      <section className="card">
        <div className="card-head">
          <h2>檢查結果</h2>
          <div className="filters" role="tablist">
            {(["all", "error", "warning", "info"] as const).map((f) => (
              <button
                key={f}
                role="tab"
                aria-selected={filter === f}
                className={filter === f ? "active" : ""}
                onClick={() => setFilter(f)}
              >
                {f === "all" ? `全部 ${result.issues.length}` : `${SEVERITY_LABEL[f]} ${counts[f]}`}
              </button>
            ))}
          </div>
        </div>
        {shown.length === 0 ? (
          <p className="muted">{result.issues.length === 0 ? "沒有發現任何問題 🎉" : "這個分類沒有問題。"}</p>
        ) : (
          <ul className="issues">
            {shown.map((issue, i) => (
              <li key={i}>
                <button
                  className="issue"
                  disabled={issue.paragraphIndex === undefined}
                  onClick={() => issue.paragraphIndex !== undefined && onJump(issue.paragraphIndex)}
                >
                  <SeverityBadge severity={issue.severity} />
                  <span className="issue-body">
                    <span>{issue.message}</span>
                    {issue.excerpt && <span className="excerpt">{issue.excerpt}</span>}
                  </span>
                  <span className="muted small rule-tag">{issue.ruleId}</span>
                </button>
              </li>
            ))}
          </ul>
        )}
      </section>
    </>
  );
}
