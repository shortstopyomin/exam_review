import type { Severity } from "@/lib/types";

export const SEVERITY_LABEL: Record<Severity, string> = {
  error: "錯誤",
  warning: "警告",
  info: "提醒",
};

export function SeverityBadge({ severity, count }: { severity: Severity; count?: number }) {
  return (
    <span className={`badge ${severity}`}>
      {SEVERITY_LABEL[severity]}
      {count !== undefined && ` ${count}`}
    </span>
  );
}
