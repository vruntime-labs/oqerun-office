import type { OfficeDocument, SpreadsheetCell, SpreadsheetContent } from "../core/DocumentModel";
import { DocumentStore } from "../core/DocumentStore";

type SheetState = { name: string; cells: Record<string, SpreadsheetCell> };

export class Sheets {
  private readonly root: HTMLElement;
  private readonly store = new DocumentStore();
  private readonly state: SpreadsheetContent = {
    activeSheet: 0,
    sheets: [{ name: "Sheet 1", cells: {} }],
  };

  constructor(root: HTMLElement) {
    this.root = root;
  }

  render(): void {
    this.root.innerHTML = `
      <main class="sheet-shell">
        <header class="sheet-topbar">
          <button class="writer-back" data-action="back" type="button">← Oqerun Office</button>
          <div class="sheet-title"><span class="sheet-app-icon">S</span><div><strong>Oqerun Sheets</strong><span>Untitled workbook</span></div></div>
          <span class="sheet-status" data-status>Ready</span>
        </header>
        <div class="sheet-toolbar">
          <button data-action="add-sheet" type="button">＋ Sheet</button>
          <button data-action="formula" type="button">Formula</button>
          <button data-action="save" class="primary-tool" type="button">Save</button>
        </div>
        <div class="sheet-tabs" data-tabs></div>
        <section class="sheet-grid-wrap">
          <table class="sheet-grid">
            <thead><tr><th></th>${Array.from({ length: 10 }, (_, i) => `<th>${this.columnName(i)}</th>`).join("")}</tr></thead>
            <tbody>${Array.from({ length: 20 }, (_, row) => `<tr><th>${row + 1}</th>${Array.from({ length: 10 }, (_, col) => {
              const ref = this.cellRef(row, col);
              return `<td><input data-cell="${ref}" aria-label="${ref}" /></td>`;
            }).join("")}</tr>`).join("")}</tbody>
          </table>
        </section>
      </main>`;
    this.bind();
    this.renderTabs();
    this.loadSheet();
  }

  private bind(): void {
    this.root.querySelectorAll<HTMLInputElement>("[data-cell]").forEach((input) =>
      input.addEventListener("input", () => {
        this.activeSheet().cells[input.dataset.cell!] = this.parseValue(input.value);
        this.setStatus("Unsaved changes");
      }),
    );

    this.root.querySelector('[data-action="back"]')?.addEventListener("click", () =>
      this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true })),
    );

    this.root.querySelector('[data-action="add-sheet"]')?.addEventListener("click", () => {
      this.state.sheets.push({ name: `Sheet ${this.state.sheets.length + 1}`, cells: {} });
      this.state.activeSheet = this.state.sheets.length - 1;
      this.renderTabs();
      this.loadSheet();
    });

    this.root.querySelector('[data-action="formula"]')?.addEventListener("click", () => this.showFormulaHelp());
    this.root.querySelector('[data-action="save"]')?.addEventListener("click", () => void this.save());
  }

  private activeSheet(): SheetState {
    return this.state.sheets[this.state.activeSheet];
  }

  private loadSheet(): void {
    this.root.querySelectorAll<HTMLInputElement>("[data-cell]").forEach((input) => {
      const value = this.activeSheet().cells[input.dataset.cell!];
      input.value = value === undefined || value === null ? "" : String(value);
    });
    this.setStatus(`Editing ${this.activeSheet().name}`);
  }

  private renderTabs(): void {
    const tabs = this.root.querySelector<HTMLElement>("[data-tabs]");
    if (!tabs) return;

    tabs.innerHTML = this.state.sheets
      .map((sheet, index) => `<button class="sheet-tab ${index === this.state.activeSheet ? "active" : ""}" data-sheet="${index}" type="button">${sheet.name}</button>`)
      .join("");

    tabs.querySelectorAll<HTMLButtonElement>("[data-sheet]").forEach((button) =>
      button.addEventListener("click", () => {
        this.state.activeSheet = Number(button.dataset.sheet);
        this.renderTabs();
        this.loadSheet();
      }),
    );
  }

  private async save(): Promise<void> {
    const now = new Date().toISOString();
    const document: OfficeDocument<SpreadsheetContent> = {
      metadata: {
        id: "oqerun-sheet-workbook",
        name: "Untitled workbook",
        path: "/home/user/Documents/untitled-workbook.oqsheet",
        mimeType: "application/x-oqerun-spreadsheet",
        createdAt: now,
        modifiedAt: now,
      },
      content: this.state,
    };
    await this.store.save(document);
    this.setStatus("Saved locally");
  }

  private showFormulaHelp(): void {
    window.alert("Supported formulas: =A1+10, =A1*2, and =SUM(A1:A5).");
  }

  private parseValue(value: string): SpreadsheetCell {
    const trimmed = value.trim();
    if (trimmed.startsWith("=")) return this.evaluateFormula(trimmed.slice(1));
    if (trimmed === "") return null;
    const number = Number(trimmed);
    return Number.isFinite(number) ? number : trimmed;
  }

  private evaluateFormula(formula: string): SpreadsheetCell {
    const sum = formula.match(/^SUM\(([A-Z]+\d+):([A-Z]+\d+)\)$/i);
    if (sum) {
      const start = this.refToCoord(sum[1]);
      const end = this.refToCoord(sum[2]);
      let total = 0;
      for (let row = start.row; row <= end.row; row += 1) {
        for (let col = start.col; col <= end.col; col += 1) {
          const value = this.activeSheet().cells[this.cellRef(row, col)];
          if (typeof value === "number") total += value;
        }
      }
      return total;
    }

    const arithmetic = formula.match(/^([A-Z]+\d+)\s*([+*\-/])\s*(-?\d+(?:\.\d+)?)$/i);
    if (arithmetic) {
      const left = this.activeSheet().cells[arithmetic[1].toUpperCase()];
      const right = Number(arithmetic[3]);
      if (typeof left !== "number") return 0;
      if (arithmetic[2] === "+") return left + right;
      if (arithmetic[2] === "-") return left - right;
      if (arithmetic[2] === "*") return left * right;
      return right === 0 ? 0 : left / right;
    }

    return formula;
  }

  private refToCoord(ref: string): { row: number; col: number } {
    const match = ref.toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!match) return { row: 0, col: 0 };
    let col = 0;
    for (const char of match[1]) col = col * 26 + char.charCodeAt(0) - 64;
    return { row: Number(match[2]) - 1, col: col - 1 };
  }

  private cellRef(row: number, col: number): string {
    return `${this.columnName(col)}${row + 1}`;
  }

  private columnName(index: number): string {
    let n = index + 1;
    let result = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      result = String.fromCharCode(65 + rem) + result;
      n = Math.floor((n - 1) / 26);
    }
    return result;
  }

  private setStatus(text: string): void {
    const status = this.root.querySelector<HTMLElement>("[data-status]");
    if (status) status.textContent = text;
  }
}
