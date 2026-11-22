"use client";

import Link from "next/link";
import { useEffect, useMemo, useRef, useState } from "react";
import { parseDocx, DocxParseError } from "@/lib/docx/parse";
import { isPdf, parsePdf, PdfParseError } from "@/lib/pdf/parse";
import { lint, parseRuleSet, RuleSetError } from "@/lib/rules/engine";
import { DEFAULT_RULES_TOML } from "@/lib/rules/default-rules";
import type { ParsedDocument, RuleSet } from "@/lib/types";
import { RulesPanel } from "./RulesPanel";
import { Results } from "./Results";
import { DocumentView } from "./DocumentView";
import { buildReport, download } from "@/lib/report";

const STORAGE_KEY = "exam-review:rules-toml";
const BASE = process.env.NEXT_PUBLIC_BASE_PATH ?? "";

function loadSavedRules(): string {
  try {
    return localStorage.getItem(STORAGE_KEY) ?? DEFAULT_RULES_TOML;
  } catch {
    return DEFAULT_RULES_TOML;
  }
}

export function ExamReviewer() {
  const [doc, setDoc] = useState<ParsedDocument | null>(null);
  const [fileError, setFileError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [rulesText, setRulesText] = useState(DEFAULT_RULES_TOML);
  const [selected, setSelected] = useState<number | null>(null);
  const [dragging, setDragging] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);

  // localStorage 只能在瀏覽器讀取，所以在掛載後才載入
  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setRulesText(loadSavedRules());
  }, []);

  const updateRules = (text: string) => {
    setRulesText(text);
    try {
      if (text === DEFAULT_RULES_TOML) localStorage.removeItem(STORAGE_KEY);
      else localStorage.setItem(STORAGE_KEY, text);
    } catch {
      /* 無痕模式等情況無法儲存，忽略 */
    }
  };

  const ruleParse = useMemo((): { set?: RuleSet; error?: string } => {
    try {
      return { set: parseRuleSet(rulesText) };
    } catch (e) {
      return { error: e instanceof RuleSetError ? e.message : String(e) };
    }
  }, [rulesText]);

  const result = useMemo(
    () => (doc && ruleParse.set ? lint(doc, ruleParse.set) : null),
    [doc, ruleParse.set],
  );

  const mainRef = useRef<HTMLElement>(null);
  // 手機版（單欄）結果在上傳區、規則檔下方，上傳完自動捲到結果
  useEffect(() => {
    if (doc && window.matchMedia("(max-width: 900px)").matches) {
      mainRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }, [doc]);

  async function openFile(data: ArrayBuffer, name: string) {
    setLoading(true);
    setFileError(null);
    setSelected(null);
    try {
      setDoc(isPdf(data) ? await parsePdf(data, name) : await parseDocx(data, name));
    } catch (e) {
      setDoc(null);
      setFileError(e instanceof DocxParseError || e instanceof PdfParseError ? e.message : `解析失敗：${(e as Error).message}`);
    } finally {
      setLoading(false);
    }
  }

  async function handleFiles(files: FileList | null) {
    const f = files?.[0];
    if (!f) return;
    if (/\.doc$/i.test(f.name)) {
      setFileError("不支援舊版 .doc 檔，請先在 Word 中「另存新檔」為 .docx 再上傳。");
      return;
    }
    await openFile(await f.arrayBuffer(), f.name);
  }

  async function loadSample() {
    setLoading(true);
    try {
      const res = await fetch(`${BASE}/samples/sample-exam.docx`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      await openFile(await res.arrayBuffer(), "範例試卷.docx");
    } catch (e) {
      setFileError(`無法載入範例試卷：${(e as Error).message}`);
      setLoading(false);
    }
  }

  function jumpTo(index: number) {
    setSelected(index);
    document.getElementById(`para-${index}`)?.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  return (
    <div className="app">
      <header className="app-header">
        <div>
          <h1>
            試卷審題小幫手
            <span className="title-alt">
              <span lang="en">Exam Review Assistant</span> · <span lang="vi">Trợ lý kiểm tra đề thi</span>
            </span>
          </h1>
          <p className="muted">上傳 PDF 或 Word 試卷，依規則檔自動檢查題號、配分、選項與文字。檔案只在你的瀏覽器中處理，不會上傳到任何伺服器。</p>
          <p className="muted" lang="en">
            Upload a PDF or Word exam to automatically check question numbers, scores, answer options, and wording against
            the rule file. Files are processed only in your browser and are never uploaded to any server.
          </p>
          <p className="muted" lang="vi">
            Tải lên đề thi PDF hoặc Word để tự động kiểm tra số thứ tự câu hỏi, điểm số, các phương án lựa chọn và nội dung
            theo tệp quy tắc. Tệp chỉ được xử lý trong trình duyệt của bạn và không được tải lên bất kỳ máy chủ nào.
          </p>
          {/* 另一個功能：組卷（/compose/），審題本身不依賴它 */}
          <Link href="/compose/" className="feature-link">
            <span>組卷小幫手</span>
            <span className="feature-link-alt">
              <span lang="en">Exam Composer</span> · <span lang="vi">Ghép đề thi</span>
            </span>
            <span aria-hidden="true">→</span>
          </Link>
        </div>
      </header>

      <div className="layout">
        <aside className="side">
          <section className="card">
            <h2>
              1. 上傳試卷
              <span className="title-alt">
                <span lang="en">Upload exam</span> · <span lang="vi">Tải lên đề thi</span>
              </span>
            </h2>
            <div
              className={`dropzone${dragging ? " dragging" : ""}`}
              onClick={() => fileInput.current?.click()}
              onDragOver={(e) => {
                e.preventDefault();
                setDragging(true);
              }}
              onDragLeave={() => setDragging(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragging(false);
                handleFiles(e.dataTransfer.files);
              }}
              role="button"
              tabIndex={0}
              onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && fileInput.current?.click()}
            >
              <strong>{loading ? "解析中…" : doc ? doc.fileName : "點這裡選擇或拖放 .pdf 或 .docx 檔"}</strong>
              {!doc && !loading && (
                <>
                  <span className="drop-alt" lang="en">Click here to choose, or drag and drop a .pdf or .docx file</span>
                  <span className="drop-alt" lang="vi">Nhấn vào đây để chọn, hoặc kéo thả tệp .pdf hoặc .docx</span>
                </>
              )}
              {doc ? (
                <span className="muted small">點擊可更換檔案</span>
              ) : (
                <span className="drop-hint">
                  <span className="muted small">支援可選取文字的 PDF 與 Word 2007 以後的 .docx</span>
                  <span className="muted small" lang="en">Supports text-based PDFs and .docx files from Word 2007 or later</span>
                  <span className="muted small" lang="vi">Hỗ trợ tệp PDF có văn bản chọn được và tệp .docx từ Word 2007 trở lên</span>
                </span>
              )}
            </div>
            <input
              ref={fileInput}
              type="file"
              accept=".docx,.pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/pdf"
              hidden
              onChange={(e) => {
                handleFiles(e.target.files);
                e.target.value = "";
              }}
            />
            <button className="link-btn" onClick={loadSample} disabled={loading}>
              沒有試卷？先用範例試卷試試看
            </button>
            {fileError && <p className="alert error">{fileError}</p>}
          </section>

          <RulesPanel
            rulesText={rulesText}
            onChange={updateRules}
            onReset={() => updateRules(DEFAULT_RULES_TOML)}
            error={ruleParse.error}
            ruleSet={ruleParse.set}
            summary={result?.ruleSummary}
          />
        </aside>

        <main className="main" ref={mainRef}>
          {!doc && !loading && (
            <div className="card empty">
              <p>上傳試卷後，檢查結果會顯示在這裡。</p>
              <p lang="en">Upload an exam and the check results will appear here.</p>
              <p lang="vi">Sau khi tải lên đề thi, kết quả kiểm tra sẽ hiển thị tại đây.</p>
            </div>
          )}
          {doc && result && (
            <>
              <Results
                result={result}
                onJump={jumpTo}
                onDownload={() =>
                  download(
                    `${doc.fileName.replace(/\.(docx|pdf)$/i, "")}-審題報告.md`,
                    buildReport(doc, result, ruleParse.set!.name),
                  )
                }
              />
              <DocumentView doc={doc} result={result} selected={selected} onSelect={setSelected} />
            </>
          )}
          {doc && !result && (
            <div className="card empty">
              <p>規則檔有錯誤，修正後會自動重新檢查。</p>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
