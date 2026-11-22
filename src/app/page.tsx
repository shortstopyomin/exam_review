import Link from "next/link";
import type { ReactNode } from "react";
import "./home.css";

/*
 * 主選單：越聰越懶的小島。每個功能是一張卡片，依用途分成「考卷」和「表單」兩區。
 * 之後新增功能時，在 GROUPS 裡加一筆就好。
 */

interface Tool {
  href: string;
  name: string;
  en: string;
  vi: string;
  desc: string;
  icon: ReactNode;
}

const icon = (children: ReactNode) => (
  <svg viewBox="0 0 24 24" width="26" height="26" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    {children}
  </svg>
);

const GROUPS: { title: string; alt: string; tools: Tool[] }[] = [
  {
    title: "考卷",
    alt: "Exams · Đề thi",
    tools: [
      {
        href: "/review/",
        name: "試卷審題小幫手",
        en: "Exam Review",
        vi: "Kiểm tra đề thi",
        desc: "上傳 PDF 或 Word 試卷，自動檢查題號、配分、選項與文字。",
        icon: icon(
          <>
            <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h4" />
            <path d="M14 3v4h4" />
            <path d="M14 3l4 4v3" />
            <circle cx="16.5" cy="16.5" r="3.5" />
            <path d="M19 19l2.5 2.5" />
          </>,
        ),
      },
      {
        href: "/compose/",
        name: "組卷小幫手",
        en: "Exam Composer",
        vi: "Ghép đề thi",
        desc: "從幾份考卷挑題目，組成一份新的 Word 試卷。",
        icon: icon(
          <>
            <rect x="8" y="3" width="12" height="15" rx="2" />
            <path d="M16 21H6a2 2 0 0 1-2-2V7" />
            <path d="M11 8h6M11 11.5h6M11 15h3" />
          </>,
        ),
      },
    ],
  },
  {
    title: "表單",
    alt: "Forms · Đơn từ",
    tools: [
      {
        href: "/leave/",
        name: "請假單小幫手",
        en: "Leave Form",
        vi: "Đơn xin phép",
        desc: "選假別和時間，自動算天數與時數，下載填好的請假單。",
        icon: icon(
          <>
            <rect x="3" y="5" width="18" height="16" rx="2" />
            <path d="M3 10h18M8 3v4M16 3v4" />
            <path d="M9 15.5l2 2 4-4" />
          </>,
        ),
      },
      {
        href: "/goout/",
        name: "外出單小幫手",
        en: "Going Out Form",
        vi: "Đơn xin ra ngoài",
        desc: "填外出時間、地點與事由，下載填好的外出單。",
        icon: icon(
          <>
            <path d="M13 4h4a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2h-4" />
            <path d="M3 12h11" />
            <path d="M10 8l4 4-4 4" />
          </>,
        ),
      },
    ],
  },
];

/** 小島插圖：太陽、海浪、椰子樹和一張吊床（懶） */
function Island() {
  return (
    <svg className="island-art" viewBox="0 0 240 160" role="img" aria-label="小島上的椰子樹與吊床">
      <circle cx="190" cy="42" r="20" fill="var(--accent)" />
      <circle cx="190" cy="42" r="28" fill="none" stroke="var(--accent)" strokeOpacity="0.35" strokeWidth="3" />
      <ellipse cx="118" cy="122" rx="78" ry="18" fill="#e9c98f" />
      <path d="M86 120 C 92 90, 92 70, 82 48" stroke="#8a5a2b" strokeWidth="6" fill="none" strokeLinecap="round" />
      <g fill="#3f9a6b">
        <path d="M82 48 C 64 36, 46 40, 38 52 C 54 46, 68 48, 82 48 Z" />
        <path d="M82 48 C 72 30, 76 18, 88 12 C 84 26, 86 38, 82 48 Z" />
        <path d="M82 48 C 98 32, 116 34, 124 46 C 108 42, 94 44, 82 48 Z" />
        <path d="M82 48 C 66 50, 56 62, 56 74 C 64 62, 72 54, 82 48 Z" />
        <path d="M82 48 C 100 52, 108 64, 106 76 C 100 64, 92 56, 82 48 Z" />
      </g>
      <path d="M150 120 V 88" stroke="#8a5a2b" strokeWidth="4" strokeLinecap="round" />
      <path d="M90 92 Q 120 112 150 92" stroke="var(--accent)" strokeWidth="4" fill="none" strokeLinecap="round" />
      <text x="128" y="80" fill="var(--header-muted)" fontSize="13" fontWeight="700" fontFamily="system-ui, sans-serif">
        z
        <tspan x="138" y="70" fontSize="11">z</tspan>
        <tspan x="146" y="62" fontSize="9">z</tspan>
      </text>
      <path d="M8 140 q 14 -10 28 0 t 28 0 t 28 0 t 28 0 t 28 0 t 28 0 t 28 0 t 28 0 t 28 0" stroke="#5fb3d4" strokeWidth="4" fill="none" strokeLinecap="round" />
      <path d="M30 154 q 14 -8 28 0 t 28 0 t 28 0 t 28 0 t 28 0 t 28 0 t 28 0" stroke="#5fb3d4" strokeOpacity="0.55" strokeWidth="3" fill="none" strokeLinecap="round" />
    </svg>
  );
}

export default function Home() {
  return (
    <div className="app island">
      <header className="app-header island-hero">
        <div>
          <h1>
            越聰越懶的小島
            <span className="title-alt">
              <span lang="en">The Smart &amp; Lazy Island</span> · <span lang="vi">Hòn đảo càng thông minh càng lười</span>
            </span>
          </h1>
          <p className="muted">把重複的文書工作交給小島，省下來的時間，拿去偷懶。</p>
          <p className="muted" lang="en">
            Hand the repetitive paperwork to the island, and spend the time you save being lazy.
          </p>
          <p className="muted" lang="vi">
            Giao việc giấy tờ lặp đi lặp lại cho hòn đảo, dành thời gian tiết kiệm được để nghỉ ngơi.
          </p>
        </div>
        <Island />
      </header>

      {GROUPS.map((g) => (
        <section key={g.title} className="island-group" aria-labelledby={`group-${g.title}`}>
          <h2 id={`group-${g.title}`}>
            {g.title} <span className="muted">{g.alt}</span>
          </h2>
          <div className="island-tools">
            {g.tools.map((t) => (
              <Link key={t.href} href={t.href} className="island-tool">
                <span className="island-tool-icon">{t.icon}</span>
                <span className="island-tool-text">
                  <strong>{t.name}</strong>
                  <span className="island-tool-alt">
                    <span lang="en">{t.en}</span> · <span lang="vi">{t.vi}</span>
                  </span>
                  <span className="island-tool-desc">{t.desc}</span>
                </span>
                <span className="island-tool-arrow" aria-hidden="true">
                  →
                </span>
              </Link>
            ))}
          </div>
        </section>
      ))}

      <p className="island-note muted small">
        所有檔案都只在你的瀏覽器中處理，不會上傳到任何伺服器。
        <span lang="en"> Files never leave your browser.</span>
      </p>
    </div>
  );
}
