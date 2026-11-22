import type { Metadata } from "next";
import { GoOutForm } from "@/components/GoOutForm";
import "../compose/compose.css";
import "../leave/leave.css";

export const metadata: Metadata = {
  title: "外出單小幫手",
  description: "填寫姓名、外出時間、地點與事由，下載格式一樣的外出單 Word 檔。全部在瀏覽器中處理。",
};

export default function GoOutPage() {
  return <GoOutForm />;
}
