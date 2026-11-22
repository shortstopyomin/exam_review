import type { Metadata } from "next";
import { LeaveForm } from "@/components/LeaveForm";
import "../compose/compose.css";
import "./leave.css";

export const metadata: Metadata = {
  title: "請假單小幫手",
  description: "填寫姓名、假別、時間與事由，下載填好的請假單 Word 檔。全部在瀏覽器中處理。",
};

export default function LeavePage() {
  return <LeaveForm />;
}
