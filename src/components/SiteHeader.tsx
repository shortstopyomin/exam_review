"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";

/*
 * 每一頁最上方的網站列：左邊「越聰越懶的小島」回主選單，右邊直接切換到各功能。
 * 新增功能頁時，在 TOOLS 加一筆（主選單卡片在 src/app/page.tsx）。
 */
const TOOLS = [
  { href: "/review/", label: "審題", title: "試卷審題小幫手 Exam Review" },
  { href: "/compose/", label: "組卷", title: "組卷小幫手 Exam Composer" },
  { href: "/leave/", label: "請假單", title: "請假單小幫手 Leave Form" },
  { href: "/goout/", label: "外出單", title: "外出單小幫手 Going Out Form" },
];

/** usePathname 不含 basePath；trailingSlash 開著時可能有、也可能沒有結尾斜線 */
const normalize = (p: string) => (p.endsWith("/") ? p : `${p}/`);

export function SiteHeader() {
  const path = normalize(usePathname() ?? "/");
  return (
    <header className="site-bar">
      <nav className="site-bar-inner" aria-label="網站導覽">
        <Link href="/" className="site-brand" aria-current={path === "/" ? "page" : undefined}>
          <span className="site-brand-mark" aria-hidden="true" />
          越聰越懶的小島
          <span className="site-brand-alt" lang="en">
            Lazy Island
          </span>
        </Link>
        <ul className="site-tools">
          {TOOLS.map((t) => (
            <li key={t.href}>
              <Link href={t.href} title={t.title} aria-current={path === t.href ? "page" : undefined}>
                {t.label}
              </Link>
            </li>
          ))}
        </ul>
      </nav>
    </header>
  );
}
