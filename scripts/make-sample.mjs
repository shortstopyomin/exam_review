// 產生範例試卷 public/samples/sample-exam.docx（刻意埋入一些問題供展示）
// 執行：npm run sample
import { writeFileSync, mkdirSync } from "node:fs";
import { Document, Packer, Paragraph, TextRun, LevelFormat, AlignmentType } from "docx";

const p = (text, opts = {}) => new Paragraph({ children: [new TextRun(text)], ...opts });
const title = (text) => new Paragraph({ children: [new TextRun({ text, bold: true })], spacing: { before: 240 } });

const choice = [
  ["( )1. 下列哪一個數是質數？", "(A) 21 (B) 27 (C) 31 (D) 33"],
  ["( )2. 若 x + 3 = 7，則 x = ？", "(A) 3 (B) 4 (C) 7 (D) 10"],
  ["( )3. 下列何者為負數？", "(A) 0 (B) -2 (C) 5"], // 只有 3 個選項
  ["( )4. 數線上 -3 與 5 的距離為多少？", "(A) 2 (B) 8 (C) -8 (D) 15"],
  ["( B )5. 計算 (-4) × (-3) = ？", "(A) -12 (B) 12 (C) -7 (D) 7"], // 留下答案
  ["( )6. 下列哪一個是偶數？", "(A) 4 (B) 9 (C) 4 (D) 15"], // 選項重複
  ["( )8. 將 0.25 化為分數，最簡分數為何？", "(A) 1/4 (B) 2/8 (C) 25/10 (D) 1/2"], // 跳過第 7 題
  ["( )9. 1 公尺等於幾公分？", "(A) 10 (B) 100 (C) 1000 (D) 10000"],
  ["( )10. 若 a = -2，則 a² = ？", "(A) -4 (B) 4 (C) -2 (D) 2"],
  ["( )11. 下列敘述何者正確？", "(A) 0 是正數 (B) -1 > 0 (C) 1 > -1 (D) 2 < -3"],
];

const trueFalse = [
  "0 是整數。",
  "任何數乘以 0 都等於 0。",
  "-5 比 -3 大。",
  "兩個負數相乘,結果是正數。", // 半形逗號
  "1 不是質數。",
  "3/6 與 1/2 相等。",
  "所有偶數都是合數。",
  "絕對值一定大於或等於 0。",
  "-0 與 0 相等。",
  "7 的倍數都是奇數。",
];

const fill = [
  "1. 計算 12 ÷ (-3) = ____。",
  "2. 15 與 20 的最大公因數為 ____。",
  "3. 15 與 20 的最小公倍數為 ____。",
  "4. 數線上與原點距離 4 的點為 ____ 與 ____。",
  "5. 將 36 做質因數分解：36 = ____。",
  "6. |-7| = ____。。", // 重複標點
  "7. 比 -2 小 3 的數是 ____。",
];

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "tf",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "( )%1.", alignment: AlignmentType.LEFT }],
      },
    ],
  },
  sections: [
    {
      children: [
        p("越華國際學校 2026-2027 學年度 第 1 學期 7年級 數學 期末評量試卷", { alignment: AlignmentType.CENTER }),
        // 拼錯：International、Mathematics
        p("Viet Hoa Internatinal School Grade 7 Mathmatics Final Exam of Semester 1 Year 2026-2027", {
          alignment: AlignmentType.CENTER,
        }),
        p("班級：________ 座號：____ 姓名：____________"),

        title("一、選擇題（每題 4 分，共 40 分）"),
        ...choice.flatMap(([q, o]) => [p(q), p(o)]),

        title("二、是非題（每題 2 分，共 20 分）"),
        // 這一大題使用 Word 自動編號，測試解析器能否還原題號
        ...trueFalse.map((t) => new Paragraph({ children: [new TextRun(t)], numbering: { reference: "tf", level: 0 } })),

        title("三、填充題（每格 3 分，共 24 分）"),
        ...fill.map((t) => p(t)),

        title("四、計算題（共 16 分）"),
        p("1. 計算 (-3)² + 4 × (-2) 的值。（8 分）"),
        p("2. 小明有 120 元，買了 3 枝筆（每枝 15 元，剩下多少錢？（6 分）"), // 括號不成對
        p("答案：75 元"), // 留下答案
      ],
    },
  ],
});

mkdirSync("public/samples", { recursive: true });
writeFileSync("public/samples/sample-exam.docx", await Packer.toBuffer(doc));
console.log("已產生 public/samples/sample-exam.docx");
