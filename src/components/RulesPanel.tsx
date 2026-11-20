"use client";

import { useRef, useState } from "react";
import type { RuleSet } from "@/lib/types";
import type { LintResult } from "@/lib/rules/engine";
import { download } from "@/lib/report";
import { SeverityBadge } from "./SeverityBadge";

interface Props {
  rulesText: string;
  onChange(text: string): void;
  onReset(): void;
  error?: string;
  ruleSet?: RuleSet;
  summary?: LintResult["ruleSummary"];
}

export function RulesPanel({ rulesText, onChange, onReset, error, ruleSet, summary }: Props) {
  const [editing, setEditing] = useState(false);
  const input = useRef<HTMLInputElement>(null);

  return (
    <section className="card">
      <h2>
        2. 規則檔
        <span className="title-alt">
          <span lang="en">Rule file</span> · <span lang="vi">Tệp quy tắc</span>
        </span>
      </h2>
      <p className="small">
        目前使用：<strong>{ruleSet?.name ?? "（規則檔有錯誤）"}</strong>
        {ruleSet && <span className="muted">（{ruleSet.rules.filter((r) => r.enabled !== false).length} 條規則）</span>}
      </p>

      <div className="btn-row">
        <button onClick={() => setEditing((v) => !v)}>{editing ? "收起編輯器" : "編輯規則"}</button>
        <button onClick={() => input.current?.click()}>上傳規則檔</button>
        <button onClick={() => download("exam-rules.toml", rulesText, "application/toml")}>下載規則檔</button>
        <button onClick={() => confirm("確定要還原成預設規則？目前的修改會被覆蓋。") && onReset()}>還原預設</button>
      </div>
      <input
        ref={input}
        type="file"
        accept=".toml,application/toml"
        hidden
        onChange={async (e) => {
          const f = e.target.files?.[0];
          if (f) onChange(await f.text());
          e.target.value = "";
        }}
      />

      {error && <p className="alert error">{error}</p>}

      {editing && (
        <textarea
          className="rules-editor"
          value={rulesText}
          onChange={(e) => onChange(e.target.value)}
          spellCheck={false}
          aria-label="規則檔 TOML"
        />
      )}

      {ruleSet && (
        <ul className="rule-list">
          {ruleSet.rules.map((r) => {
            const s = summary?.find((x) => x.id === r.id);
            const off = r.enabled === false;
            return (
              <li key={r.id} className={off ? "off" : ""} title={s?.description}>
                <span className="rule-status">
                  {off ? "—" : !s ? "·" : s.count === 0 ? <span className="ok">✓</span> : <SeverityBadge severity={s.severity} count={s.count} />}
                </span>
                <span className="rule-id">{r.id}</span>
                <span className="muted small">{r.type}</span>
              </li>
            );
          })}
        </ul>
      )}
    </section>
  );
}
