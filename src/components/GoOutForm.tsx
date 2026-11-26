"use client";

import { useEffect, useMemo, useState } from "react";
import { Alt, DropZone } from "@/components/DropZone";
import { saveProfile, todayIso, useStoredProfile } from "@/components/formProfile";
import { slashDate } from "@/lib/forms/docx";
import {
  fillGoOutForm,
  GoOutFormError,
  goOutFileName,
  goOutTimeProblem,
  inspectGoOutForm,
  WORK_END,
  WORK_START,
  type GoOutInfo,
} from "@/lib/goout/goout";

/** 外出單版型（F-QL-034）內建在程式裡（public/goout/），打開頁面就在背景載入；改版時可以在第 3 區上傳新的 */
const BASE = process.env.NEXT_PUBLIC_BASE_PATH ?? "";
const BUILTIN_NAME = "F-QL-034 外出單";
const TEMPLATE_URL = `${BASE}/goout/F-QL-034-going-out-form.docx`;

async function loadTemplate(): Promise<ArrayBuffer> {
  const res = await fetch(TEMPLATE_URL);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.arrayBuffer();
  await inspectGoOutForm(data.slice(0));
  return data;
}

export function GoOutForm() {
  const [template, setTemplate] = useState<ArrayBuffer | null>(null);
  const [templateFailed, setTemplateFailed] = useState(false);
  const [busy, setBusy] = useState<"template" | "fill" | null>(null);
  /** 使用者上傳的版型檔名；null 表示用內建版型 */
  const [uploaded, setUploaded] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    loadTemplate().then(
      // 使用者已經自己上傳了就不要蓋掉
      (t) => !cancelled && setTemplate((cur) => cur ?? t),
      () => !cancelled && setTemplateFailed(true),
    );
    return () => {
      cancelled = true;
    };
  }, []);

  // 沒改過的欄位顯示上次記住的值（跟請假單共用）
  const stored = useStoredProfile();
  const [nameEdit, setName] = useState<string | null>(null);
  const [deptEdit, setDept] = useState<string | null>(null);
  const [positionEdit, setPosition] = useState<string | null>(null);
  const [contactEdit, setContact] = useState<string | null>(null);
  const name = nameEdit ?? stored.name ?? "";
  const dept = deptEdit ?? stored.dept ?? "";
  const position = positionEdit ?? stored.position ?? "";
  const contact = contactEdit ?? stored.contact ?? "";

  const [outDate, setOutDate] = useState(todayIso);
  const [outTime, setOutTime] = useState(WORK_START);
  // 返回日期沒改過就跟外出日期一樣
  const [backDateEdit, setBackDate] = useState<string | null>(null);
  const backDate = backDateEdit ?? outDate;
  const [backTime, setBackTime] = useState(WORK_END);
  const [destination, setDestination] = useState("");
  const [reason, setReason] = useState("");
  const [applicant, setApplicant] = useState<string | null>(null);
  const [date, setDate] = useState(todayIso);
  // 沒改過就用上次記住的申請人，沒記過就跟姓名一樣
  const signer = applicant ?? (stored.applicant || name);

  const info: GoOutInfo = {
    name,
    dept,
    position,
    outDate,
    outTime,
    backDate,
    backTime,
    destination,
    contact,
    reason,
    applicant: signer,
    date,
  };
  const timeProblem = goOutTimeProblem(info);

  const missing = useMemo(() => {
    if (!template) return templateFailed ? "外出單載入失敗，請重新整理頁面" : "載入外出單中…";
    if (!name.trim()) return "請填姓名";
    if (!outDate || !outTime) return "請選擇外出時間";
    if (!backDate || !backTime) return "請選擇預計返回時間";
    if (timeProblem) return timeProblem;
    if (!destination.trim()) return "請填外出地點";
    if (!reason.trim()) return "請填事由";
    if (!date) return "請填填表日期";
    return null;
  }, [template, templateFailed, name, outDate, outTime, backDate, backTime, timeProblem, destination, reason, date]);

  async function onTemplate(files: File[]) {
    const file = files[0];
    if (!file) return;
    setBusy("template");
    setError(null);
    try {
      const data = await file.arrayBuffer();
      await inspectGoOutForm(data.slice(0));
      setTemplate(data);
      setUploaded(file.name);
    } catch (e) {
      setError(e instanceof GoOutFormError ? e.message : `讀取失敗：${(e as Error).message}`);
    } finally {
      setBusy(null);
    }
  }

  async function restoreBuiltin() {
    setError(null);
    try {
      setTemplate(await loadTemplate());
      setUploaded(null);
    } catch (e) {
      setError(`內建外出單載入失敗：${(e as Error).message}`);
    }
  }

  async function download() {
    if (!template || missing) return;
    setBusy("fill");
    setError(null);
    try {
      saveProfile({ name, dept, position, contact, applicant: signer.trim() === name.trim() ? "" : signer });
      const bytes = await fillGoOutForm(template.slice(0), info);
      const url = URL.createObjectURL(
        new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
      );
      const a = document.createElement("a");
      a.href = url;
      a.download = goOutFileName(info);
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (e) {
      setError(e instanceof GoOutFormError ? e.message : `產生失敗：${(e as Error).message}`);
    } finally {
      setBusy(null);
    }
  }

  const when = (d: string, t: string) => [d ? slashDate(d) : "", t].filter(Boolean).join(" ");

  return (
    <div className="app compose leave goout">
      <header className="app-header">
        <div>
          <h1>
            外出單小幫手
            <span className="title-alt">
              <span lang="en">Permission for Going Out</span> · <span lang="vi">Đơn xin ra ngoài</span>
            </span>
          </h1>
          <p className="muted">填好姓名、外出時間、地點與事由，下載填好的外出單 Word 檔。資料只在你的瀏覽器中處理。</p>
          <p className="muted" lang="en">
            Fill in your details and download the completed going-out form as a Word file. Your data never leaves your
            browser.
          </p>
          <p className="muted" lang="vi">
            Điền thông tin và tải về đơn xin ra ngoài đã điền dưới dạng tệp Word. Dữ liệu không rời khỏi trình duyệt của bạn.
          </p>
        </div>
      </header>

      {error && <p className="alert error">{error}</p>}

      <div className="compose-steps">
        <section className="card">
          <h2>
            1. 基本資料
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
                部門 <span className="muted">Division</span>
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
          <p className="muted small">姓名、部門、職務、聯繫方式、申請人會在下載時記在這台電腦的瀏覽器裡（跟請假單共用），下次自動帶入。</p>
        </section>

        <section className="card">
          <h2>
            2. 申請人簽名
            <Alt en="Applicant" vi="Người làm đơn" />
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
                填表日期 <span className="muted">Date</span>
              </span>
              <input type="date" value={date} onChange={(e) => setDate(e.target.value)} />
            </label>
          </div>
          <div className="signature-preview" aria-label="簽名欄預覽">
            <span>{signer || "　"}</span>
            <span>{date ? slashDate(date) : "dd/mm/yyyy"}</span>
          </div>
          <p className="muted small">填表日期也會填在外出單上方的「填表日期 Date」。</p>
        </section>

        <section className="card">
          <h2>
            3. 更換版型（選填）
            <Alt en="Form template" vi="Đổi mẫu đơn" />
          </h2>
          {uploaded ? (
            <p className="leave-builtin">
              <strong>使用上傳的 {uploaded}</strong>
              <button type="button" onClick={restoreBuiltin}>
                改回內建版型
              </button>
            </p>
          ) : (
            <p className="muted small leave-builtin">目前使用內建的 {BUILTIN_NAME}，不用上傳。</p>
          )}
          <DropZone
            label="上傳新的外出單 .docx"
            hint="外出單改版時再上傳，表格、日期欄、頁首都會照原樣保留"
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            busy={busy === "template"}
            onFiles={onTemplate}
          />
        </section>
      </div>

      <section className="card leave-details">
        <h2>
          4. 外出內容
          <Alt en="Going out" vi="Nội dung ra ngoài" />
        </h2>

        <div className="period-grid">
          <label>
            <span>
              外出時間 <span className="muted">Time out</span>
            </span>
            <span className="dt">
              <input type="date" value={outDate} onChange={(e) => setOutDate(e.target.value)} />
              <input type="time" value={outTime} min={WORK_START} max={WORK_END} onChange={(e) => setOutTime(e.target.value)} />
            </span>
          </label>
          <label>
            <span>
              預計返回時間 <span className="muted">Intended time to come back</span>
            </span>
            <span className="dt">
              <input type="date" value={backDate} min={outDate} onChange={(e) => setBackDate(e.target.value)} />
              <input type="time" value={backTime} min={WORK_START} max={WORK_END} onChange={(e) => setBackTime(e.target.value)} />
            </span>
            {timeProblem && (
              <span className="total-problem" role="status">
                {timeProblem}
              </span>
            )}
          </label>
        </div>

        <div className="period-grid wide">
          <label>
            <span>
              外出地點 <span className="muted">Destination</span>
            </span>
            <input value={destination} onChange={(e) => setDestination(e.target.value)} />
          </label>
          <label>
            <span>
              聯繫方式 <span className="muted">Contact</span>
            </span>
            <input value={contact} onChange={(e) => setContact(e.target.value)} autoComplete="tel" placeholder="手機號碼" />
          </label>
        </div>

        <h3>
          外出事由 <span className="muted">Reason · Lý do ra ngoài</span>
        </h3>
        <textarea value={reason} onChange={(e) => setReason(e.target.value)} rows={3} />
        <p className="muted small">外出與返回時間要在上班時間 {WORK_START}–{WORK_END} 內。保安確認欄留白，由保安簽名。</p>
      </section>

      <div className="compose-bar">
        <div>
          <strong>{name || "（未填姓名）"}</strong>
          <span className="muted small">
            {" "}
            {[
              outTime ? `${when(outDate, outTime)} ~ ${backDate === outDate ? backTime : when(backDate, backTime)}` : "",
              destination,
            ]
              .filter(Boolean)
              .join(" · ")}
          </span>
        </div>
        <button className="primary" onClick={download} disabled={!!missing || busy !== null}>
          {busy === "fill" ? "產生中…" : (missing ?? "下載外出單 Word 檔")}
        </button>
      </div>
    </div>
  );
}
