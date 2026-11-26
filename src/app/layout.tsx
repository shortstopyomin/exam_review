import type { Metadata } from "next";
import { SiteHeader } from "@/components/SiteHeader";
import "./globals.css";

export const metadata: Metadata = {
  title: "越聰越懶的小島",
  description: "試卷審題、組卷、請假單、外出單小幫手。全部在瀏覽器中處理，檔案不會上傳。",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="zh-Hant-TW">
      <body>
        <SiteHeader />
        {children}
      </body>
    </html>
  );
}
