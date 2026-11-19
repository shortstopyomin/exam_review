/** 從 docx 解析出來的一個段落 */
export interface Paragraph {
  index: number;
  /** 段落文字（已含 Word 自動編號產生的題號） */
  text: string;
  /** Word 自動編號產生的標籤，例如 "1."、"（一）" */
  numberingLabel?: string;
  inTable: boolean;
  hasMath: boolean;
  imageCount: number;
}

export interface ParsedDocument {
  fileName: string;
  paragraphs: Paragraph[];
}

export type SectionKind = "choice" | "truefalse" | "fill" | "short" | "other";

export interface Question {
  number: number;
  /** 題號前的作答括號內容，例如 "( B )1." 會得到 "B" */
  answerMark: string;
  startParagraph: number;
  /** 最後一個屬於本題的段落（含） */
  endParagraph: number;
  /** 本題所有段落文字（去掉題號前綴） */
  body: string;
  /** 題目中標示的配分，例如「（5分）」 */
  score?: number;
}

export interface Section {
  title: string;
  /** 標題所在段落；-1 代表沒有大題標題（題目出現在任何大題之前） */
  paragraphIndex: number;
  kind: SectionKind;
  perQuestionScore?: number;
  totalScore?: number;
  questions: Question[];
}

export interface ExamStructure {
  sections: Section[];
}

export type Severity = "error" | "warning" | "info";

export interface Issue {
  ruleId: string;
  severity: Severity;
  message: string;
  paragraphIndex?: number;
  excerpt?: string;
}

/** 規則檔中一條規則的原始設定 */
export interface RuleConfig {
  id: string;
  type: string;
  severity?: Severity;
  message?: string;
  enabled?: boolean;
  [key: string]: unknown;
}

export interface RuleSet {
  name: string;
  rules: RuleConfig[];
}
