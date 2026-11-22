import type { Metadata } from "next";
import { ExamComposer } from "@/components/ExamComposer";
import "./compose.css";

export const metadata: Metadata = {
  title: "組卷小幫手",
  description: "上傳考卷版型與多份試卷，勾選題目後產生新的 Word 考卷。全部在瀏覽器中處理。",
};

export default function ComposePage() {
  return <ExamComposer />;
}
