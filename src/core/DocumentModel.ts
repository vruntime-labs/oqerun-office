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

export type WriterContent = {
  html: string;
  plainText: string;
  blocks: WriterBlock[];
};

export type SpreadsheetCell = string | number | boolean | null;

export type SpreadsheetContent = {
  activeSheet: number;
  sheets: Array<{
    name: string;
    cells: Record<string, SpreadsheetCell>;
  }>;
};

export type PresentationElement = {
  id: string;
  type: "text" | "image" | "shape";
  value: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
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
};

export type PresentationContent = {
  activeSlide: number;
  slides: PresentationSlide[];
};
