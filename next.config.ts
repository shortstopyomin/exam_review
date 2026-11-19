import type { NextConfig } from "next";

// GitHub Pages 會把專案放在 https://<user>.github.io/<repo>/ 底下，
// 所以部署時要設定 basePath。GitHub Actions 會透過 PAGES_BASE_PATH 傳入；
// 本機開發時為空字串。
const basePath = process.env.PAGES_BASE_PATH ?? "";

const nextConfig: NextConfig = {
  output: "export",
  basePath,
  trailingSlash: true,
  images: { unoptimized: true },
  env: {
    NEXT_PUBLIC_BASE_PATH: basePath,
  },
};

export default nextConfig;
