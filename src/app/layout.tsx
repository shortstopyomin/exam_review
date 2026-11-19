import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "試卷審題小幫手",
  description: "上傳 Word 試卷，依規則檔自動檢查題號、配分、選項與文字。全部在瀏覽器中處理。",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="zh-Hant-TW">
      <body>{children}</body>
    </html>
  );
}
