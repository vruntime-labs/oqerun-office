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

export type WriterContent = {
  blocks: Array<{
    type: "paragraph" | "heading" | "list" | "table" | "image";
    data: unknown;
  }>;
};

export type SpreadsheetContent = {
  sheets: Array<{
    name: string;
    cells: Record<string, unknown>;
  }>;
};

export type PresentationContent = {
  slides: Array<{
    id: string;
    elements: unknown[];
  }>;
};