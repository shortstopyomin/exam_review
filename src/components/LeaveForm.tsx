"use client";

import Link from "next/link";
import { useEffect, useMemo, useState, useSyncExternalStore } from "react";
import { Alt, DropZone } from "@/components/DropZone";
import {
  estimateDuration,
  fillLeaveForm,
  formatDuration,
  formatPeriod,
  inspectLeaveForm,
  LEAVE_TYPES,
  leaveFileName,
  LeaveFormError,
  slashDate,
  type LeaveInfo,
  type LeaveType,
  WORK_END,
  WORK_START,
} from "@/lib/leave/leave";

interface Template {
  name: string;
  data: ArrayBuffer;
  types: LeaveType[];
  /** 程式內建的版型（public/leave/），不是使用者上傳的 */
  builtin?: boolean;
}

/** 內建的請假單版型（F-QL-013G），打開頁面就自動載入，不用上傳 */
const BASE = process.env.NEXT_PUBLIC_BASE_PATH ?? "";
const BUILTIN_URL = `${BASE}/leave/F-QL-013G-leave-form.docx`;
const BUILTIN_NAME = "F-QL-013G 請假單";

async function loadBuiltin(): Promise<Template> {
  const res = await fetch(BUILTIN_URL);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.arrayBuffer();
  const t = await inspectLeaveForm(data.slice(0));
  return { name: BUILTIN_NAME, data, types: t.types, builtin: true };
}

/** 姓名、單位、職務記在這台瀏覽器，下次不用重打 */
const PROFILE_KEY = "leave-form-profile";

interface Profile {
  name?: string;
  dept?: string;
  position?: string;
}

const noSubscribe = () => () => {};
function readProfileRaw(): string | null {
  try {
    return localStorage.getItem(PROFILE_KEY);
  } catch {
    return null; // 無痕模式或被封鎖時就不記
  }
}
function parseProfile(raw: string | null): Profile {
  try {
    const p = JSON.parse(raw ?? "{}");
    const str = (v: unknown) => (typeof v === "string" ? v : undefined);
    return { name: str(p.name), dept: str(p.dept), position: str(p.position) };
  } catch {
    return {};
  }
}

function todayIso() {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

export function LeaveForm() {
  const [template, setTemplate] = useState<Template | null>(null);
  const [builtinState, setBuiltinState] = useState<"loading" | "ready" | "failed">("loading");
  const [busy, setBusy] = useState<"template" | "fill" | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    loadBuiltin().then(
      (t) => {
        if (cancelled) return;
        setBuiltinState("ready");
        // 使用者已經自己上傳了就不要蓋掉
        setTemplate((cur) => cur ?? t);
      },
      () => !cancelled && setBuiltinState("failed"),
    );
    return () => {
      cancelled = true;
    };
  }, []);

  async function restoreBuiltin() {
    setError(null);
    try {
      setTemplate(await loadBuiltin());
    } catch (e) {
      setError(`內建請假單載入失敗：${(e as Error).message}`);
    }
  }

  // 沒改過的欄位顯示上次記住的值
  const storedRaw = useSyncExternalStore(noSubscribe, readProfileRaw, () => null);
  const stored = useMemo(() => parseProfile(storedRaw), [storedRaw]);
  const [nameEdit, setName] = useState<string | null>(null);
  const [deptEdit, setDept] = useState<string | null>(null);
  const [positionEdit, setPosition] = useState<string | null>(null);
  const name = nameEdit ?? stored.name ?? "";
  const dept = deptEdit ?? stored.dept ?? "";
  const position = positionEdit ?? stored.position ?? "";
  const [type, setType] = useState<LeaveType | null>(null);
  const [otherText, setOtherText] = useState("");

  const [startDate, setStartDate] = useState("");
  const [startTime, setStartTime] = useState(WORK_START);
  const [endDate, setEndDate] = useState("");
  const [endTime, setEndTime] = useState(WORK_END);
  // 請假時間與總計會依起訖自動產生；手動改過就不再覆蓋
  const [periodText, setPeriodText] = useState<string | null>(null);
  const [totalText, setTotalText] = useState<string | null>(null);

  const [reason, setReason] = useState("");
  const [applicant, setApplicant] = useState<string | null>(null);
  const [date, setDate] = useState(todayIso);


  const autoPeriod = formatPeriod(startDate, startTime, endDate || startDate, endTime);
  const duration = estimateDuration(startDate, startTime, endDate || startDate, endTime);
  const autoTotal = duration ? formatDuration(duration) : "";
  // 選了日期卻算不出時數時，說明原因（不要默默留白）
  const totalProblem = (() => {
    if (!startDate || duration) return null;
    const end = endDate || startDate;
    if (end < startDate || (end === startDate && endTime && startTime && endTime <= startTime))
      return "結束時間要晚於開始時間";
    return `選的時間不在上班時間 ${WORK_START}–${WORK_END} 內，不算請假時數`;
  })();
  const period = periodText ?? autoPeriod;
  const total = totalText ?? autoTotal;
  const signer = applicant ?? name;
  const available = template?.types ?? LEAVE_TYPES.map((t) => t.zh);

  const info: LeaveInfo = {
    name,
    dept,
    position,
    type,
    otherText,
    period,
    total,
    reason,
    applicant: signer,
    date: slashDate(date),
  };

  const missing = useMemo(() => {
    if (!template) return builtinState === "loading" ? "載入請假單中…" : "請先上傳請假單";
    if (!name.trim()) return "請填姓名";
    if (!type) return "請選擇假別";
    if (!period.trim()) return "請填請假時間";
    if (!date) return "請填申請日期";
    return null;
  }, [template, builtinState, name, type, period, date]);

  async function onTemplate(files: File[]) {
    const file = files[0];
    if (!file) return;
    setBusy("template");
    setError(null);
    try {
      const data = await file.arrayBuffer();
      const t = await inspectLeaveForm(data.slice(0));
      setTemplate({ name: file.name, data, types: t.types });
      setType((prev) => (prev && t.types.includes(prev) ? prev : (t.prefill.type ?? null)));
      // 上傳的是填過的請假單時，帶入原本的內容（只填空白的欄位）
      const fill = (cur: string, set: (v: string) => void, v?: string) => !cur.trim() && v && set(v);
      fill(name, setName, t.prefill.name);
      fill(dept, setDept, t.prefill.dept);
      fill(position, setPosition, t.prefill.position);
      fill(reason, setReason, t.prefill.reason);
      if (t.prefill.period && !startDate) setPeriodText(t.prefill.period);
      if (t.prefill.total && !startDate) setTotalText(t.prefill.total);
    } catch (e) {
      setError(e instanceof LeaveFormError ? e.message : `讀取失敗：${(e as Error).message}`);
    } finally {
      setBusy(null);
    }
  }

  async function download() {
    if (!template || missing) return;
    setBusy("fill");
    setError(null);
    try {
      try {
        localStorage.setItem(PROFILE_KEY, JSON.stringify({ name, dept, position }));
      } catch {
        /* 記不住也沒關係 */
      }
      const bytes = await fillLeaveForm(template.data.slice(0), info);
      const url = URL.createObjectURL(
        new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
      );
      const a = document.createElement("a");
      a.href = url;
      a.download = leaveFileName(info);
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (e) {
      setError(e instanceof LeaveFormError ? e.message : `產生失敗：${(e as Error).message}`);
    } finally {
      setBusy(null);
    }
  }

  return (
    <div className="app compose leave">
      <header className="app-header">
        <div>
          <h1>
            請假單小幫手
            <span className="title-alt">
              <span lang="en">Leave Application</span> · <span lang="vi">Đơn xin nghỉ phép</span>
            </span>
          </h1>
          <p className="muted">填好姓名、假別、時間與事由，下載填好的請假單 Word 檔。資料只在你的瀏覽器中處理。</p>
          <p className="muted" lang="en">
            Fill in your details and download the completed leave form as a Word file. Your data never leaves your browser.
          </p>
          <p className="muted" lang="vi">
            Điền thông tin và tải về đơn xin phép đã điền dưới dạng tệp Word. Dữ liệu không rời khỏi trình duyệt của bạn.
          </p>
          <p className="nav-link">
            <Link href="/">← 試卷審題 Exam Review</Link>
            <Link href="/compose/">組卷小幫手 Exam Composer</Link>
          </p>
        </div>
      </header>

      {error && <p className="alert error">{error}</p>}

      <div className="compose-steps">
        <section className="card">
          <h2>
            1. 請假單版型
            <Alt en="Leave form" vi="Mẫu đơn" />
          </h2>
          {template?.builtin ? (
            <p className="leave-builtin">
              <strong>使用內建 {BUILTIN_NAME}</strong>
              <span className="muted small">不用上傳，直接往下填寫</span>
            </p>
          ) : (
            builtinState === "loading" && !template && <p className="muted small leave-found">載入內建請假單中…</p>
          )}
          {template && !template.builtin && (
            <p className="leave-builtin">
              <strong>使用上傳的 {template.name}</strong>
              {builtinState === "ready" && (
                <button type="button" onClick={restoreBuiltin}>
                  改回內建版型
                </button>
              )}
            </p>
          )}
          {builtinState === "failed" && !template && (
            <p className="muted small leave-found">內建請假單載入失敗，請上傳請假單 Word 檔。</p>
          )}
          <DropZone
            label={template ? "更換版型（選填）" : "點這裡選擇或拖放請假單 .docx"}
            hint={template ? "請假單改版時，再上傳新的 Word 檔" : "表格、勾選框、頁首都會照原樣保留"}
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            busy={busy === "template"}
            onFiles={onTemplate}
          />
        </section>

        <section className="card">
          <h2>
            2. 基本資料
            <Alt en="Employee" vi="Thông tin" />
          </h2>
          <div className="info-fields">
            <label>
              <span>
                姓名 <span className="muted">Full name</span>
              </span>
              <input value={name} onChange={(e) => setName(e.target.value)} autoComplete="name" />
            </label>
            <label>
              <span>
                單位 <span className="muted">Dept</span>
              </span>
              <input value={dept} onChange={(e) => setDept(e.target.value)} autoComplete="organization" />
            </label>
            <label>
              <span>
                職務 <span className="muted">Position</span>
              </span>
              <input value={position} onChange={(e) => setPosition(e.target.value)} autoComplete="organization-title" />
            </label>
          </div>
          <p className="muted small">姓名、單位、職務會記在這台電腦的瀏覽器裡，下次自動帶入。</p>
        </section>

        <section className="card">
          <h2>
            3. 申請人簽名
            <Alt en="Applicant" vi="Người xin phép" />
          </h2>
          <div className="info-fields">
            <label>
              <span>
                申請人 <span className="muted">Applicant</span>
              </span>
              <input value={signer} onChange={(e) => setApplicant(e.target.value)} placeholder={name} />
            </label>
            <label>
              <span>
                申請日期 <span className="muted">Application date</span>
              </span>
              <input type="date" value={date} onChange={(e) => setDate(e.target.value)} />
            </label>
          </div>
          <div className="signature-preview" aria-label="簽名欄預覽">
            <span>{signer || "　"}</span>
            <span>{date ? slashDate(date) : "yyyy/mm/dd"}</span>
          </div>
        </section>
      </div>

      <section className="card leave-details">
        <h2>
          4. 請假內容
          <Alt en="Leave details" vi="Nội dung nghỉ phép" />
        </h2>

        <h3>
          假別 <span className="muted">Leave type · Loại phép</span>
        </h3>
        <div className="leave-types" role="radiogroup" aria-label="假別">
          {LEAVE_TYPES.filter((t) => available.includes(t.zh)).map((t) => (
            <label key={t.zh} className={type === t.zh ? "on" : ""}>
              <input type="radio" name="leave-type" checked={type === t.zh} onChange={() => setType(t.zh)} />
              <span>
                <strong>{t.zh}</strong>
                <span className="muted small">
                  {t.en} · <span lang="vi">{t.vi}</span>
                </span>
              </span>
            </label>
          ))}
        </div>
        {type === "其他" && (
          <label className="other-text">
            其他說明 <span className="muted">Other</span>
            <input value={otherText} onChange={(e) => setOtherText(e.target.value)} placeholder="例如：防疫假" />
          </label>
        )}

        <h3>
          請假時間與日數 <span className="muted">Time and days · Thời gian và số ngày nghỉ</span>
        </h3>
        <div className="period-grid">
          <label>
            <span>
              開始 <span className="muted">From</span>
            </span>
            <span className="dt">
              <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
              <input type="time" value={startTime} onChange={(e) => setStartTime(e.target.value)} />
            </span>
          </label>
          <label>
            <span>
              結束 <span className="muted">To</span>
            </span>
            <span className="dt">
              <input type="date" value={endDate} min={startDate} onChange={(e) => setEndDate(e.target.value)} />
              <input type="time" value={endTime} onChange={(e) => setEndTime(e.target.value)} />
            </span>
          </label>
        </div>
        <div className="period-grid wide">
          <label>
            <span>
              寫在請假單上的時間 <span className="muted">Shown on the form</span>
            </span>
            <span className="dt">
              <input value={period} onChange={(e) => setPeriodText(e.target.value)} placeholder="選擇開始與結束時間，或直接輸入" />
              {periodText !== null && (
                <button type="button" onClick={() => setPeriodText(null)} title="依開始與結束時間重新產生">
                  重算
                </button>
              )}
            </span>
          </label>
          <label>
            <span>
              總計 <span className="muted">Total</span>
            </span>
            <span className="dt">
              <input value={total} onChange={(e) => setTotalText(e.target.value)} />
              {totalText !== null && (
                <button type="button" onClick={() => setTotalText(null)} title="依開始與結束時間重新計算">
                  重算
                </button>
              )}
            </span>
            {totalProblem && totalText === null && (
              <span className="total-problem" role="status">
                {totalProblem}
              </span>
            )}
          </label>
        </div>
        <p className="muted small">
          依上班時間 {WORK_START}–{WORK_END} 計算（一天 9 小時，週六、週日也照算），不滿一天以 1 小時為單位、無條件進位，最少 1
          小時。總計可以直接修改。
        </p>

        <h3>
          請假事由 <span className="muted">Reason · Lý do xin phép</span>
        </h3>
        <textarea value={reason} onChange={(e) => setReason(e.target.value)} rows={3} />
      </section>

      <div className="compose-bar">
        <div>
          <strong>{name || "（未填姓名）"}</strong>
          <span className="muted small">
            {" "}
            {[type === "其他" && otherText ? `其他：${otherText}` : type, period, total]
              .filter(Boolean)
              .join(" · ")}
          </span>
        </div>
        <button className="primary" onClick={download} disabled={!!missing || busy !== null}>
          {busy === "fill" ? "產生中…" : (missing ?? "下載請假單 Word 檔")}
        </button>
      </div>
    </div>
  );
}
