import type { OfficeModule } from "./Office";

export const FILE_ASSOCIATIONS: Record<string, OfficeModule> = {
  ".doc": "writer",
  ".docx": "writer",
  ".odt": "writer",
  ".rtf": "writer",
  ".xls": "sheets",
  ".xlsx": "sheets",
  ".ods": "sheets",
  ".csv": "sheets",
  ".ppt": "slides",
  ".pptx": "slides",
  ".odp": "slides",
};

export function moduleForFile(path: string): OfficeModule | null {
  const name = path.toLowerCase();
  const dot = name.lastIndexOf(".");
  if (dot === -1) return null;
  return FILE_ASSOCIATIONS[name.slice(dot)] ?? null;
}