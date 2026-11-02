import { describe, expect, it } from "vitest";
import { columnIndex, columnLabel, formatAddress, parseAddress, parseRange, rangeCellCount, rangeAddresses } from "./cell-address";
import { tableRowGridCols } from "../../slides/engine/table-grid";

describe("spreadsheet address engine", () => {
  it("round-trips A1 references across column boundaries", () => {
    expect(parseAddress("$AA$27")).toEqual({ row: 26, column: 26 });
    expect(formatAddress(26, 26)).toBe("AA27");
    expect(columnLabel(701)).toBe("ZZ");
    expect(columnIndex("ZZ")).toBe(701);
  });
  it("normalizes reversed ranges and expands addresses", () => {
    const bounds = parseRange("C3:A1");
    expect(bounds).toEqual({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 });
    expect(rangeCellCount(bounds)).toBe(9);
    expect(rangeAddresses(bounds)).toEqual(["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"]);
  });
  it("rejects malformed A1 references", () => {
    expect(() => parseAddress("a1")).toThrow();
    expect(() => parseAddress("A0")).toThrow();
    expect(() => parseAddress("A1:B2")).toThrow();
  });
});

describe("presentation table grid", () => {
  it("accounts for grid-span continuation cells", () => {
    expect(tableRowGridCols([{ gridSpan: 2 }, { merged: true }, {}])).toEqual([0, 1, 2]);
  });
  it("advances by a span when continuation cells are omitted", () => {
    expect(tableRowGridCols([{ gridSpan: 3 }, {}])).toEqual([0, 3]);
  });
});
