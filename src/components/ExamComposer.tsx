"use client";

import Link from "next/link";
import { useMemo, useRef, useState } from "react";
import {
  ComposeError,
  composeExam,
  defaultSectionFor,
  inspectTemplate,
  loadDocxSource,
  loadPdfSource,
  markDuplicates,
  outputFileName,
  type ExamInfo,
  type ExamType,
  type SourceExam,
  type SourceQuestion,
  type TemplateInfo,
} from "@/lib/compose/compose";
import { DocxParseError } from "@/lib/docx/parse";
import { isPdf, parsePdf, PdfParseError } from "@/lib/pdf/parse";

interface Template {
  name: string;
  data: ArrayBuffer;
  info: TemplateInfo;
}

const LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
const sourceLabel = (i: number) => `試卷 ${LETTERS[i % 26]}${i >= 26 ? Math.floor(i / 26) : ""}`;

/** 大題標題去掉配分說明，只留名稱 */
const shortTitle = (t: string) => t.replace(/\s*[（(][^)）]*(?:分|points?|marks?|%|％)[^)）]*[)）]\s*$/i, "").trim();

function errorText(e: unknown): string {
  if (e instanceof ComposeError || e instanceof DocxParseError || e instanceof PdfParseError) return e.message;
  return `解析失敗：${(e as Error).message}`;
}

function Alt({ en, vi }: { en: string; vi: string }) {
  return (
    <span className="title-alt">
      <span lang="en">{en}</span> · <span lang="vi">{vi}</span>
    </span>
  );
}

function DropZone({
  label,
  hint,
  accept,
  multiple,
  busy,
  onFiles,
}: {
  label: string;
  hint: string;
  accept: string;
  multiple?: boolean;
  busy?: boolean;
  onFiles: (files: File[]) => void;
}) {
  const input = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  return (
    <>
      <div
        className={`dropzone${dragging ? " dragging" : ""}`}
        onClick={() => input.current?.click()}
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDragging(false);
          onFiles(Array.from(e.dataTransfer.files));
        }}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && input.current?.click()}
      >
        <strong>{busy ? "解析中…" : label}</strong>
        <span className="muted small">{hint}</span>
      </div>
      <input
        ref={input}
        type="file"
        accept={accept}
        multiple={multiple}
        hidden
        onChange={(e) => {
          onFiles(Array.from(e.target.files ?? []));
          e.target.value = "";
        }}
      />
    </>
  );
}

export function ExamComposer() {
  const [template, setTemplate] = useState<Template | null>(null);
  const [examType, setExamType] = useState<ExamType | null>(null);
  // 學年度（開始年份）、學期、年級：上傳版型時從標題帶入，可以改
  const [startYear, setStartYear] = useState("");
  const [semester, setSemester] = useState("");
  const [grade, setGrade] = useState("");
  const [sources, setSources] = useState<SourceExam[]>([]);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [placement, setPlacement] = useState<Record<string, number>>({});
  const [filter, setFilter] = useState<number | "all">("all");
  const [busy, setBusy] = useState<"template" | "sources" | "compose" | null>(null);
  const [error, setError] = useState<string | null>(null);
  const nextIndex = useRef(0);

  const allQuestions = useMemo(() => sources.flatMap((s) => s.questions), [sources]);
  const byId = useMemo(() => new Map(allQuestions.map((q) => [q.id, q])), [allQuestions]);
  const labelOf = (q: SourceQuestion) => `${sourceLabel(q.sourceIndex)} 第 ${q.number} 題`;

  const sectionOf = (q: SourceQuestion) =>
    template ? (placement[q.id] ?? defaultSectionFor(q.kind, template.info.sections)) : 0;

  /** 依版型的大題分組，組內依上傳順序、原題號 */
  const groups = useMemo(() => {
    if (!template) return [];
    return template.info.sections.map((sec, si) => ({
      sec,
      questions: allQuestions.filter(
        (q) => (placement[q.id] ?? defaultSectionFor(q.kind, template.info.sections)) === si,
      ),
    }));
  }, [template, allQuestions, placement]);

  async function onTemplate(files: File[]) {
    const f = files[0];
    if (!f) return;
    setError(null);
    if (!/\.docx$/i.test(f.name)) {
      setError("版型請上傳 Word（.docx）檔。");
      return;
    }
    setBusy("template");
    try {
      const data = await f.arrayBuffer();
      const info = await inspectTemplate(data, f.name);
      setTemplate({ name: f.name, data, info });
      setStartYear(info.startYear !== undefined ? String(info.startYear) : "");
      setSemester(info.semester !== undefined ? String(info.semester) : "");
      setGrade(info.grade !== undefined ? String(info.grade) : "");
      setPlacement({});
    } catch (e) {
      setError(errorText(e));
    } finally {
      setBusy(null);
    }
  }

  async function addSources(files: { name: string; data: ArrayBuffer }[]) {
    setError(null);
    setBusy("sources");
    const added: SourceExam[] = [];
    const errors: string[] = [];
    for (const f of files) {
      if (/\.doc$/i.test(f.name)) {
        errors.push(`${f.name}：不支援舊版 .doc，請先另存成 .docx。`);
        continue;
      }
      const index = nextIndex.current++;
      try {
        const exam = isPdf(f.data) ? loadPdfSource(await parsePdf(f.data, f.name), index) : await loadDocxSource(f.data, f.name, index);
        if (!exam.questions.length) errors.push(`${f.name}：找不到題目。`);
        else added.push(exam);
      } catch (e) {
        errors.push(`${f.name}：${errorText(e)}`);
      }
    }
    const next = [...sources, ...added];
    markDuplicates(next);
    setSources(next);
    // 新加入的題目預設勾選；重複題預設不勾
    setSelected((prev) => {
      const s = new Set(prev);
      for (const ex of added) for (const q of ex.questions) if (!q.duplicateOf) s.add(q.id);
      return s;
    });
    if (errors.length) setError(errors.join("\n"));
    setBusy(null);
  }

  function removeSource(index: number) {
    const next = sources.filter((s) => s.index !== index);
    markDuplicates(next);
    setSources(next);
    setSelected((prev) => new Set([...prev].filter((id) => !id.startsWith(`${index}:`))));
    if (filter === index) setFilter("all");
  }

  const toggle = (id: string, on: boolean) =>
    setSelected((prev) => {
      const s = new Set(prev);
      if (on) s.add(id);
      else s.delete(id);
      return s;
    });

  const setMany = (qs: SourceQuestion[], on: boolean) =>
    setSelected((prev) => {
      const s = new Set(prev);
      for (const q of qs) {
        if (on) s.add(q.id);
        else s.delete(q.id);
      }
      return s;
    });

  const picked = groups.map((g) => g.questions.filter((q) => selected.has(q.id)));
  const pickedCount = picked.reduce((a, g) => a + g.length, 0);
  const pdfPicked = picked.some((g) => g.some((q) => q.lines));

  async function download() {
    if (!template || !examInfo) return;
    setBusy("compose");
    setError(null);
    try {
      const bytes = await composeExam({
        template: template.data,
        templateName: template.name,
        info: examInfo,
        sections: picked,
        sources,
      });
      const url = URL.createObjectURL(
        new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
      );
      const a = document.createElement("a");
      a.href = url;
      a.download = outputFileName(template.name, examInfo);
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (e) {
      setError(errorText(e));
    } finally {
      setBusy(null);
    }
  }

  const typeLabel = (t: ExamType) => (t === "midterm" ? "期中考 Midterm" : "期末考 Final");
  // 版型標題裡有這個欄位才需要填
  const has = {
    year: template?.info.startYear !== undefined,
    semester: template?.info.semester !== undefined,
    grade: template?.info.grade !== undefined,
  };
  const yearOk = !has.year || /^\d{4}$/.test(startYear);
  const semesterOk = !has.semester || /^[1-4]$/.test(semester);
  const gradeOk = !has.grade || (/^\d{1,2}$/.test(grade) && Number(grade) >= 1 && Number(grade) <= 12);
  const examInfo: ExamInfo | null = examType
    ? {
        examType,
        startYear: has.year && yearOk ? Number(startYear) : undefined,
        semester: has.semester && semesterOk ? Number(semester) : undefined,
        grade: has.grade && gradeOk ? Number(grade) : undefined,
      }
    : null;
  const missing = !template
    ? "請先上傳版型"
    : !examType
      ? "請選擇期中或期末考"
      : !yearOk
        ? "學年度請填四位數年份"
        : !semesterOk
          ? "請選擇學期"
          : !gradeOk
            ? "年級請填 1–12"
            : !pickedCount
              ? "請至少勾選一題"
              : null;

  return (
    <div className="app compose">
      <header className="app-header">
        <div>
          <h1>
            組卷小幫手
            <span className="title-alt">
              <span lang="en">Exam Composer</span> · <span lang="vi">Trợ lý ghép đề thi</span>
            </span>
          </h1>
          <p className="muted">
            上傳考卷版型與多份試卷，勾選要的題目，產生一份新的 Word 考卷。頁首、頁尾與版面沿用版型，檔案只在你的瀏覽器中處理。
          </p>
          <p className="muted" lang="en">
            Upload a template and several exams, tick the questions you want, and download a new Word exam that keeps the
            template&apos;s header, footer, and layout. Files never leave your browser.
          </p>
          <p className="nav-link">
            <Link href="/">← 試卷審題 Exam Review</Link>
          </p>
        </div>
      </header>

      {error && <p className="alert error">{error}</p>}

      <div className="compose-steps">
        <section className="card">
          <h2>
            1. 考卷版型
            <Alt en="Template" vi="Mẫu đề thi" />
          </h2>
          <DropZone
            label={template ? template.name : "點這裡選擇或拖放版型 .docx"}
            hint={template ? "點擊可更換版型" : "標題、作答說明、大題標題、頁首頁尾都會沿用；版型裡的題目不會保留"}
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            busy={busy === "template"}
            onFiles={onTemplate}
          />
          {template && (
            <>
              <ul className="tpl-sections">
                {template.info.sections.map((s, i) => (
                  <li key={i}>{shortTitle(s.title)}</li>
                ))}
              </ul>
              <button
                className="link-btn"
                disabled={busy !== null}
                onClick={() => addSources([{ name: template.name, data: template.data.slice(0) }])}
              >
                把版型裡的題目也加入題庫
              </button>
            </>
          )}
        </section>

        <section className="card">
          <h2>
            2. 考試資訊
            <Alt en="Exam details" vi="Thông tin kỳ thi" />
          </h2>
          <div className="segmented" role="radiogroup" aria-label="考試類型">
            {(["midterm", "final"] as const).map((t) => (
              <button
                key={t}
                role="radio"
                aria-checked={examType === t}
                className={examType === t ? "active" : ""}
                onClick={() => setExamType(t)}
              >
                {t === "midterm" ? "期中考" : "期末考"}
                <span lang="en">{t === "midterm" ? "Midterm" : "Final"}</span>
              </button>
            ))}
          </div>

          {template && (
            <div className="info-fields">
              <label className={has.year ? "" : "off"}>
                <span>
                  學年度 <span className="muted" lang="en">School year</span>
                </span>
                <span className="year-input">
                  <input
                    type="text"
                    inputMode="numeric"
                    maxLength={4}
                    value={startYear}
                    disabled={!has.year}
                    aria-invalid={!yearOk}
                    onChange={(e) => setStartYear(e.target.value.replace(/\D/g, ""))}
                  />
                  <span className="muted">– {/^\d{4}$/.test(startYear) ? Number(startYear) + 1 : "____"}</span>
                </span>
              </label>
              <div className={`field${has.semester ? "" : " off"}`}>
                <span>
                  學期 <span className="muted" lang="en">Semester</span>
                </span>
                <div className="segmented small" role="radiogroup" aria-label="學期">
                  {["1", "2"].map((v) => (
                    <button
                      key={v}
                      role="radio"
                      aria-checked={semester === v}
                      className={semester === v ? "active" : ""}
                      disabled={!has.semester}
                      onClick={() => setSemester(v)}
                    >
                      第 {v} 學期
                    </button>
                  ))}
                </div>
              </div>
              <label className={has.grade ? "" : "off"}>
                <span>
                  年級 <span className="muted" lang="en">Grade</span>
                </span>
                <select value={grade} disabled={!has.grade} aria-invalid={!gradeOk} onChange={(e) => setGrade(e.target.value)}>
                  {!gradeOk && <option value={grade}>請選擇</option>}
                  {Array.from({ length: 12 }, (_, i) => String(i + 1)).map((g) => (
                    <option key={g} value={g}>
                      {g} 年級 / Grade {g}
                    </option>
                  ))}
                </select>
              </label>
            </div>
          )}

          <p className="muted small">
            {template
              ? `版型標題與檔名裡的期中／期末、學年度、學期、年級會換成這裡的設定。${
                  !has.year || !has.semester || !has.grade
                    ? `版型標題裡找不到的欄位（${[!has.year && "學年度", !has.semester && "學期", !has.grade && "年級"].filter(Boolean).join("、")}）無法修改。`
                    : ""
                }`
              : "上傳版型後，會從標題帶入學年度、學期與年級。"}
            {template?.info.examType && ` 版型目前是${typeLabel(template.info.examType)}。`}
          </p>
          {template && examInfo && (
            <p className="small file-preview">
              檔名：<code>{outputFileName(template.name, examInfo)}</code>
            </p>
          )}
        </section>

        <section className="card">
          <h2>
            3. 題目來源試卷
            <Alt en="Source exams" vi="Đề thi nguồn" />
          </h2>
          <DropZone
            label="點這裡選擇或拖放試卷（可一次選多份）"
            hint="Word .docx 會完整保留算式與圖片；PDF 只能取出文字"
            accept=".docx,.pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/pdf"
            multiple
            busy={busy === "sources"}
            onFiles={async (files) =>
              addSources(await Promise.all(files.map(async (f) => ({ name: f.name, data: await f.arrayBuffer() }))))
            }
          />
          {sources.length > 0 && (
            <ul className="source-list">
              {sources.map((s) => (
                <li key={s.index}>
                  <span className="source-tag">{sourceLabel(s.index)}</span>
                  <span className="source-name">{s.name}</span>
                  <span className="muted small">{s.questions.length} 題</span>
                  <button onClick={() => removeSource(s.index)} aria-label={`移除 ${s.name}`}>
                    移除
                  </button>
                </li>
              ))}
            </ul>
          )}
        </section>
      </div>

      {template && allQuestions.length > 0 && (
        <section className="card pick">
          <div className="card-head">
            <h2>
              4. 挑選題目
              <Alt en="Choose questions" vi="Chọn câu hỏi" />
            </h2>
            {sources.length > 1 && (
              <div className="filters">
                <button className={filter === "all" ? "active" : ""} onClick={() => setFilter("all")}>
                  全部
                </button>
                {sources.map((s) => (
                  <button key={s.index} className={filter === s.index ? "active" : ""} onClick={() => setFilter(s.index)}>
                    {sourceLabel(s.index)}
                  </button>
                ))}
              </div>
            )}
          </div>

          {groups.map((g, si) => {
            const shown = g.questions.filter((q) => filter === "all" || q.sourceIndex === filter);
            const count = g.questions.filter((q) => selected.has(q.id)).length;
            return (
              <div key={si} className="pick-group">
                <div className="pick-group-head">
                  <h3>
                    {shortTitle(g.sec.title)}
                    <span className="muted small">
                      {" "}
                      已選 {count} / {g.questions.length}
                    </span>
                  </h3>
                  <div className="btn-row plain">
                    <button onClick={() => setMany(shown, true)} disabled={!shown.length}>
                      全選
                    </button>
                    <button onClick={() => setMany(shown, false)} disabled={!shown.length}>
                      全不選
                    </button>
                  </div>
                </div>
                {shown.length === 0 && <p className="muted small">沒有題目。</p>}
                <ol className="questions">
                  {shown.map((q) => {
                    const dup = q.duplicateOf ? byId.get(q.duplicateOf) : undefined;
                    const on = selected.has(q.id);
                    return (
                      <li key={q.id} className={`q${on ? " on" : ""}`}>
                        <label className="q-main">
                          <input type="checkbox" checked={on} onChange={(e) => toggle(q.id, e.target.checked)} />
                          <span className="q-body">
                            <span className="q-meta">
                              <span className="source-tag">{sourceLabel(q.sourceIndex)}</span>
                              <span>原第 {q.number} 題</span>
                              <span className="muted">{shortTitle(q.sectionTitle)}</span>
                              {dup && <span className="badge warning">與{labelOf(dup)}重複</span>}
                              {q.hasMath && <span className="badge info">含算式</span>}
                              {q.imageCount > 0 && <span className="badge info">含圖 {q.imageCount}</span>}
                              {q.lines && <span className="badge warning">PDF 純文字</span>}
                            </span>
                            <span className="q-text">{q.text || "（只有圖片或算式）"}</span>
                          </span>
                        </label>
                        {template.info.sections.length > 1 && (
                          <select
                            className="q-move"
                            value={sectionOf(q)}
                            aria-label="放到哪一個大題"
                            onChange={(e) => setPlacement((p) => ({ ...p, [q.id]: Number(e.target.value) }))}
                          >
                            {template.info.sections.map((s, i) => (
                              <option key={i} value={i}>
                                放到：{shortTitle(s.title)}
                              </option>
                            ))}
                          </select>
                        )}
                      </li>
                    );
                  })}
                </ol>
              </div>
            );
          })}
        </section>
      )}

      <div className="compose-bar">
        <div>
          <strong>已選 {pickedCount} 題</strong>
          {picked.length > 0 && (
            <span className="muted small">
              {" "}
              （{groups.map((g, i) => `${shortTitle(g.sec.title).split(/\s*\/\s*/)[0]} ${picked[i].length}`).join("、")}）
            </span>
          )}
          {pdfPicked && <span className="muted small"> · PDF 題目只有文字，算式和圖要再補</span>}
        </div>
        <button className="primary" onClick={download} disabled={!!missing || busy !== null}>
          {busy === "compose" ? "產生中…" : missing ?? `下載 ${examType ? typeLabel(examType) : ""} Word 檔`}
        </button>
      </div>
    </div>
  );
}
