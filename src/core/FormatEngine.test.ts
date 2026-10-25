import { afterEach, describe, expect, it, vi } from "vitest";
import {
  importSheets,
  importSlidesPptx,
  importWriterDocx,
} from "./FormatEngine";

function oversizedFile(name: string): File {
  return {
    name,
    size: 25 * 1024 * 1024 + 1,
    arrayBuffer: vi.fn(),
  } as unknown as File;
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe("Office import size guards", () => {
  it("rejects oversized Writer documents before reading their bytes", async () => {
    const file = oversizedFile("large.docx");
    await expect(importWriterDocx(file)).rejects.toThrow(/25 MB import limit/i);
    expect(file.arrayBuffer).not.toHaveBeenCalled();
  });

  it("rejects oversized spreadsheets before reading their bytes", async () => {
    const file = oversizedFile("large.xlsx");
    await expect(importSheets(file)).rejects.toThrow(/25 MB import limit/i);
    expect(file.arrayBuffer).not.toHaveBeenCalled();
  });

  it("rejects oversized presentations before reading their bytes", async () => {
    const file = oversizedFile("large.pptx");
    await expect(importSlidesPptx(file)).rejects.toThrow(/25 MB import limit/i);
    expect(file.arrayBuffer).not.toHaveBeenCalled();
  });
});
