"use client";

import { useEffect, useLayoutEffect, useRef, useState } from "react";
import type { SourceQuestion } from "@/lib/compose/compose";

interface OrderSection {
  title: string;
  questions: SourceQuestion[];
  /** 已經手動排過（顯示「恢復原順序」） */
  customized: boolean;
}

interface Props {
  sections: OrderSection[];
  /** 題目來源標籤，例如「試卷 A」 */
  sourceLabel: (q: SourceQuestion) => string;
  onMove: (section: number, from: number, to: number) => void;
  onReset: (section: number) => void;
}

/**
 * 新考卷的題目順序：拖曳題目排序，或用 ↑↓ 按鈕。
 * 用 pointer events 自己做拖曳（HTML5 drag and drop 在手機、平板上不可靠）。
 * 滑鼠可以拖整列；觸控只能拖把手，否則滑過列表時頁面不能捲動。
 *
 * 拖曳效果：抓起的題目放大、加陰影，平順地跟著指標移動（transform，不影響排版）；
 * 經過其他題目時直接改順序，其他題目用 FLIP 動畫滑開讓位；放開時滑回新位置。
 * 靠近畫面上下緣時自動捲動。
 */
const LIFT = "scale(1.02)";
const SLIDE_MS = 160;
const reducedMotion = () => typeof window !== "undefined" && window.matchMedia("(prefers-reduced-motion: reduce)").matches;

export function QuestionOrder({ sections, sourceLabel, onMove, onReset }: Props) {
  const [drag, setDrag] = useState<{ section: number; id: string } | null>(null);
  /** 放開後滑回定位的那一題（保留陰影到動畫結束） */
  const [settling, setSettling] = useState<string | null>(null);
  const lists = useRef<(HTMLOListElement | null)[]>([]);
  // 抓起時的位置（頁面座標）與題目在列表裡的位置；最後一次的指標位置
  const grab = useRef({ startPageY: 0, startTop: 0, clientY: 0 });
  // 改順序前各題的位置，排版更新後用來做滑動動畫
  const before = useRef<{ section: number; tops: Map<string, number> } | null>(null);
  // 拖曳中的事件處理掛在 window 上，透過 ref 拿到最新的 onMove
  const moveRef = useRef(onMove);
  useEffect(() => {
    moveRef.current = onMove;
  });

  const itemsOf = (section: number) => [...(lists.current[section]?.querySelectorAll<HTMLElement>(":scope > li") ?? [])];
  const itemById = (section: number, id: string) => itemsOf(section).find((li) => li.dataset.id === id);

  /** 改順序：先記下各題位置，排版更新後再滑過去 */
  const move = (section: number, from: number, to: number) => {
    before.current = { section, tops: new Map(itemsOf(section).map((li) => [li.dataset.id!, li.offsetTop])) };
    moveRef.current(section, from, to);
  };

  /** 抓起的題目跟著指標：目標位置減掉它現在排版上的位置 */
  const followPointer = (section: number, id: string) => {
    const li = itemById(section, id);
    if (!li) return;
    const g = grab.current;
    const dy = g.startTop + (g.clientY + window.scrollY - g.startPageY) - li.offsetTop;
    li.style.transform = `translateY(${dy}px) ${LIFT}`;
  };

  // 每次排版更新後：其他題目從舊位置滑到新位置（FLIP），抓起的題目重新對準指標
  useLayoutEffect(() => {
    const b = before.current;
    before.current = null;
    if (b && !reducedMotion()) {
      for (const li of itemsOf(b.section)) {
        if (li.dataset.id === drag?.id) continue;
        const dy = (b.tops.get(li.dataset.id!) ?? li.offsetTop) - li.offsetTop;
        if (!dy) continue;
        li.style.transition = "none";
        li.style.transform = `translateY(${dy}px)`;
        void li.offsetHeight; // 先套用舊位置，下一步才會有動畫
        li.style.transition = `transform ${SLIDE_MS}ms ease`;
        li.style.transform = "";
      }
    }
    if (drag) followPointer(drag.section, drag.id);
  });

  useEffect(() => {
    if (!drag) return;
    const update = () => {
      const li = itemById(drag.section, drag.id);
      if (!li) return;
      const items = itemsOf(drag.section);
      const from = items.indexOf(li);
      // 新位置 = 其他題目裡，中線在抓起題目中線上方的有幾題
      const g = grab.current;
      const center = g.startTop + (g.clientY + window.scrollY - g.startPageY) + li.offsetHeight / 2;
      let to = 0;
      for (const other of items) if (other !== li && other.offsetTop + other.offsetHeight / 2 < center) to++;
      if (to !== from) move(drag.section, from, to);
      followPointer(drag.section, drag.id);
    };
    const onPointerMove = (e: PointerEvent) => {
      e.preventDefault();
      grab.current.clientY = e.clientY;
      update();
      // 上方有網站列、下方有下載列，邊緣留寬一點
      if (e.clientY < 90) window.scrollBy(0, -14);
      else if (e.clientY > window.innerHeight - 110) window.scrollBy(0, 14);
    };
    const end = () => {
      const li = itemById(drag.section, drag.id);
      setDrag(null);
      if (!li || reducedMotion()) {
        if (li) li.style.transform = "";
        return;
      }
      // 從指標的位置滑回排版上的位置
      setSettling(drag.id);
      li.style.transition = `transform ${SLIDE_MS}ms ease, box-shadow ${SLIDE_MS}ms ease`;
      li.style.transform = "";
      window.setTimeout(() => {
        li.style.transition = "";
        setSettling((s) => (s === drag.id ? null : s));
      }, SLIDE_MS);
    };
    window.addEventListener("pointermove", onPointerMove, { passive: false });
    window.addEventListener("scroll", update);
    window.addEventListener("pointerup", end);
    window.addEventListener("pointercancel", end);
    return () => {
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("scroll", update);
      window.removeEventListener("pointerup", end);
      window.removeEventListener("pointercancel", end);
    };
    // itemsOf / move / followPointer 只讀 ref，不需要列為依賴
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [drag]);

  return (
    <div className={`order${drag ? " is-dragging" : ""}`}>
      {sections.map((sec, si) => (
        <div key={si} className="order-group">
          <div className="order-group-head">
            <h3>
              {sec.title} <span className="muted small">{sec.questions.length} 題</span>
            </h3>
            {sec.customized && (
              <button className="link-btn" onClick={() => onReset(si)}>
                恢復原順序
              </button>
            )}
          </div>
          {sec.questions.length === 0 ? (
            <p className="muted small">這個大題沒有勾選題目。</p>
          ) : (
            <ol className="order-list" ref={(el) => void (lists.current[si] = el)}>
              {sec.questions.map((q, i) => {
                const last = i === sec.questions.length - 1;
                return (
                  <li
                    key={q.id}
                    data-id={q.id}
                    className={drag?.id === q.id ? "dragging" : settling === q.id ? "settling" : undefined}
                    onPointerDown={(e) => {
                      if (e.button !== 0) return;
                      const target = e.target as HTMLElement;
                      // ↑↓ 按鈕照常點；手機上只有把手能拖，滑過列表才能正常捲動頁面
                      if (target.closest(".order-arrows")) return;
                      if (e.pointerType === "touch" && !target.closest(".order-handle")) return;
                      e.preventDefault();
                      const li = e.currentTarget;
                      grab.current = { startPageY: e.clientY + window.scrollY, startTop: li.offsetTop, clientY: e.clientY };
                      // 抓起時只有陰影有動畫；位置要即時跟著指標，不能有 transition
                      li.style.transition = `box-shadow ${SLIDE_MS}ms ease`;
                      li.style.transform = LIFT;
                      setDrag({ section: si, id: q.id });
                    }}
                  >
                    <button
                      className="order-handle"
                      aria-label={`第 ${i + 1} 題，拖曳或用上下方向鍵調整順序`}
                      onKeyDown={(e) => {
                        if (e.key === "ArrowUp" && i > 0) move(si, i, i - 1);
                        else if (e.key === "ArrowDown" && !last) move(si, i, i + 1);
                        else return;
                        e.preventDefault();
                      }}
                    >
                      <svg viewBox="0 0 10 16" width="10" height="16" aria-hidden="true">
                        {[3, 8, 13].flatMap((y) => [2, 8].map((x) => <circle key={`${x}-${y}`} cx={x} cy={y} r="1.5" />))}
                      </svg>
                    </button>
                    <span className="order-number">{i + 1}.</span>
                    <span className="order-body">
                      <span className="order-meta">
                        <span className="source-tag">{sourceLabel(q)}</span>
                        <span className="muted small">原第 {q.number} 題</span>
                      </span>
                      <span className="order-text">{q.text || "（只有圖片或算式）"}</span>
                    </span>
                    <span className="order-arrows">
                      <button onClick={() => move(si, i, i - 1)} disabled={i === 0} aria-label="往上移">
                        ↑
                      </button>
                      <button onClick={() => move(si, i, i + 1)} disabled={last} aria-label="往下移">
                        ↓
                      </button>
                    </span>
                  </li>
                );
              })}
            </ol>
          )}
        </div>
      ))}
    </div>
  );
}
