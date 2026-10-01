export interface OfficeDocumentMetadata {
  id: string;
  name: string;
  path?: string;
  mimeType?: string;
  createdAt: string;
  modifiedAt: string;
}

export interface OfficeDocument<TContent = unknown> {
  metadata: OfficeDocumentMetadata;
  content: TContent;
}

export type WriterBlock = {
  type: "paragraph" | "heading" | "list" | "table" | "image" | "equation";
  data: unknown;
};

export type WriterSettings = {
  fontFamily: string;
  fontSize: number;
  textColor: string;
  highlightColor: string;
  lineSpacing: number;
  indent: number;
  pageSize: "A4" | "Letter";
  margin: number;
  header: string;
  footer: string;
};

export type WriterContent = {
  html: string;
  plainText: string;
  blocks: WriterBlock[];
  settings?: WriterSettings;
};

export type SpreadsheetCell = string | number | boolean | null;

export type SpreadsheetFormatting = {
  bold?: boolean;
  italic?: boolean;
  color?: string;
  background?: string;
  align?: "left" | "center" | "right";
};

export type SpreadsheetSheet = {
  name: string;
  cells: Record<string, SpreadsheetCell>;
  formats?: Record<string, SpreadsheetFormatting>;
  columnWidths?: Record<string, number>;
  rowHeights?: Record<string, number>;
  frozenRows?: number;
  frozenColumns?: number;
};

export type SpreadsheetContent = {
  activeSheet: number;
  sheets: SpreadsheetSheet[];
};

export type PresentationElement = {
  id: string;
  type: "text" | "image" | "shape";
  value: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  rotation?: number;
  fontSize?: number;
  color?: string;
  background?: string;
};

export type PresentationSlide = {
  id: string;
  title: string;
  body: string;
  notes: string;
  theme: "light" | "dark" | "midnight" | "paper";
  background: string;
  layout: "title" | "title-body" | "blank";
  elements: PresentationElement[];
  transition?: "none" | "fade" | "slide" | "zoom";
};

export type PresentationContent = {
  activeSlide: number;
  slides: PresentationSlide[];
};
