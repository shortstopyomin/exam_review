"use client";

import { useMemo } from "react";
import type { LintResult } from "@/lib/rules/engine";
import type { Issue, ParsedDocument, Severity } from "@/lib/types";

interface Props {
  doc: ParsedDocument;
  result: LintResult;
  selected: number | null;
  onSelect(index: number): void;
}

const RANK: Record<Severity, number> = { error: 0, warning: 1, info: 2 };

export function DocumentView({ doc, result, selected, onSelect }: Props) {
  const { byPara, sectionAt, questionAt, instructionAt } = useMemo(() => {
    const byPara = new Map<number, Issue[]>();
    for (const i of result.issues) {
      if (i.paragraphIndex === undefined) continue;
      byPara.set(i.paragraphIndex, [...(byPara.get(i.paragraphIndex) ?? []), i]);
    }
    // 大題標題，以及標題下一行的配分說明
    const sectionAt = new Set(result.structure.sections.flatMap((s) => [s.paragraphIndex, ...(s.scoreLines ?? [])]));
    const questionAt = new Set(result.structure.sections.flatMap((s) => s.questions.map((q) => q.startParagraph)));
    const instructionAt = new Set(result.structure.instructions);
    return { byPara, sectionAt, questionAt, instructionAt };
  }, [result]);

  return (
    <section className="card">
      <h2>試卷內容</h2>
      <p className="muted small">有標色的段落代表有問題，點一下可以看詳細說明。</p>
      <div className="doc">
        {doc.paragraphs.map((p) => {
          if (!p.text.trim() && !p.imageCount) return null;
          const issues = byPara.get(p.index);
          const worst = issues?.reduce<Severity>((w, i) => (RANK[i.severity] < RANK[w] ? i.severity : w), "info");
          const cls = [
            "para",
            sectionAt.has(p.index) && "is-section",
            questionAt.has(p.index) && "is-question",
            instructionAt.has(p.index) && "is-instruction",
            p.inTable && "in-table",
            worst && `has-${worst}`,
            selected === p.index && "selected",
          ]
            .filter(Boolean)
            .join(" ");
          return (
            <div key={p.index} id={`para-${p.index}`} className={cls} onClick={() => issues && onSelect(p.index)}>
              <div className="para-text">
                {p.text}
                {p.imageCount > 0 && <span className="chip">圖片 ×{p.imageCount}</span>}
                {p.hasMath && <span className="chip">方程式</span>}
                {instructionAt.has(p.index) && <span className="chip">作答說明</span>}
              </div>
              {issues && selected === p.index && (
                <ul className="para-issues">
                  {issues.map((i, k) => (
                    <li key={k} className={i.severity}>
                      {i.message}
                    </li>
                  ))}
                </ul>
              )}
            </div>
          );
        })}
      </div>
    </section>
  );
}
