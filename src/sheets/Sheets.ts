import type { OfficeDocument, SpreadsheetCell, SpreadsheetContent } from "../core/DocumentModel";
import { DocumentStore } from "../core/DocumentStore";
import { FileMenu } from "../ui/FileMenu";

type SheetState = { name: string; cells: Record<string, SpreadsheetCell> };

export class Sheets {
  private readonly root: HTMLElement;
  private readonly store = new DocumentStore();
  private readonly state: SpreadsheetContent = {
    activeSheet: 0,
    sheets: [{ name: "Sheet 1", cells: {} }],
  };
  private documentId = crypto.randomUUID();
  private createdAt = new Date().toISOString();
  private fileName = "Untitled workbook";

  constructor(root: HTMLElement) {
    this.root = root;
  }

  render(): void {
    this.root.innerHTML = `
      <main class="sheet-shell">
        <header class="sheet-topbar">
          <button class="writer-back" data-action="back" type="button">← Oqerun Office</button>
          <div class="sheet-title"><span class="sheet-app-icon">S</span><div><strong>Oqerun Sheets</strong><span data-file-name>Untitled workbook</span></div></div>
          <span class="sheet-status" data-status>Ready</span>
        </header>
        <div class="sheet-toolbar">
          <button data-action="add-sheet" type="button">＋ Sheet</button>
          <button data-action="formula" type="button">ƒx Formula</button>
          <label class="sheet-open-button">Open from<input data-open-file type="file" accept=".oqsheet,.csv" hidden></label>
          <button data-action="save" class="primary-tool" type="button">Save</button>
        </div>
        <div class="sheet-formula-bar">
          <span class="formula-label">ƒx</span>
          <input data-formula-bar placeholder="Select a cell or enter a formula such as =SUM(A1:A5)" />
        </div>
        <div class="sheet-tabs" data-tabs></div>
        <section class="sheet-grid-wrap">
          <table class="sheet-grid">
            <thead><tr><th></th>${Array.from({ length: 12 }, (_, i) => `<th>${this.columnName(i)}</th>`).join("")}</tr></thead>
            <tbody>${Array.from({ length: 24 }, (_, row) => `<tr><th>${row + 1}</th>${Array.from({ length: 12 }, (_, col) => {
              const ref = this.cellRef(row, col);
              return `<td><input data-cell="${ref}" aria-label="${ref}" /></td>`;
            }).join("")}</tr>`).join("")}</tbody>
          </table>
        </section>
      </main>`;
    this.bind();
    this.renderTabs();
    this.loadSheet();
    this.mountFileMenu();
  }

  async openPath(path: string): Promise<void> {
    const document = await this.store.load<SpreadsheetContent>(path);
    this.documentId = document.metadata.id;
    this.createdAt = document.metadata.createdAt;
    this.fileName = document.metadata.name;
    this.state.activeSheet = document.content.activeSheet;
    this.state.sheets.splice(0, this.state.sheets.length, ...document.content.sheets);
    const name = this.root.querySelector<HTMLElement>("[data-file-name]");
    if (name) name.textContent = document.metadata.name;
    this.renderTabs();
    this.loadSheet();
    this.setStatus("Opened");
  }

  private mountFileMenu(): void {
    new FileMenu(this.root, {
      module: "sheets",
      title: "Oqerun Sheets",
      acceptedTypes: ".oqsheet,.csv",
      onHome: () => this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true })),
      onNew: () => {
        this.documentId = crypto.randomUUID();
        this.createdAt = new Date().toISOString();
        this.fileName = "Untitled workbook";
        this.state.activeSheet = 0;
        this.state.sheets.splice(0, this.state.sheets.length, { name: "Sheet 1", cells: {} });
        this.renderTabs();
        this.loadSheet();
        this.setStatus("New workbook");
      },
      onSave: () => void this.save(),
      onSaveAs: () => void this.saveAs(),
      onDownload: () => this.download(),
      onPrint: () => window.print(),
      onOpenRecent: (path) => void this.openPath(path),
      onOpenFile: (file) => void this.importFile(file),
    }).mount();
  }
  private bind(): void {
    this.root.querySelectorAll<HTMLInputElement>("[data-cell]").forEach((input) =>
      input.addEventListener("focus", () => this.updateFormulaBar(input.dataset.cell!)),
    );

    this.root.querySelectorAll<HTMLInputElement>("[data-cell]").forEach((input) =>
      input.addEventListener("input", () => {
        this.activeSheet().cells[input.dataset.cell!] = this.parseValue(input.value);
        this.updateFormulaBar(input.dataset.cell!);
        this.setStatus("Unsaved changes");
      }),
    );

    this.root.querySelector<HTMLInputElement>("[data-formula-bar]")?.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") return;
      const selected = this.root.querySelector<HTMLInputElement>("[data-cell]:focus");
      if (!selected) return;
      selected.value = (event.currentTarget as HTMLInputElement).value;
      selected.dispatchEvent(new Event("input", { bubbles: true }));
      selected.blur();
    });

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

    this.root.querySelector<HTMLInputElement>("[data-open-file]")?.addEventListener("change", (event) => {
      const file = (event.currentTarget as HTMLInputElement).files?.[0];
      if (!file) return;
      void this.importFile(file);
    });
  }

  private async importFile(file: File): Promise<void> {
    if (file.name.toLowerCase().endsWith(".oqsheet")) {
      const text = await file.text();
      try {
        const parsed = JSON.parse(text) as OfficeDocument<SpreadsheetContent>;
        if (parsed.content?.sheets) {
          this.state.activeSheet = parsed.content.activeSheet;
          this.state.sheets.splice(0, this.state.sheets.length, ...parsed.content.sheets);
          this.renderTabs();
          this.loadSheet();
          this.setStatus("Imported workbook");
        }
      } catch {
        this.setStatus("Could not import workbook");
      }
      return;
    }

    const cells: Record<string, SpreadsheetCell> = {};
    const text = await file.text();
    text.split(/\r?\n/).forEach((line, row) => {
      this.parseCsvLine(line).forEach((value, col) => {
        cells[this.cellRef(row, col)] = this.parseValue(value);
      });
    });
    this.state.activeSheet = 0;
    this.state.sheets.splice(0, this.state.sheets.length, { name: file.name.replace(/\.csv$/i, ""), cells });
    this.renderTabs();
    this.loadSheet();
    this.setStatus("Imported CSV");
  }

  private parseCsvLine(line: string): string[] {
    const result: string[] = [];
    let current = "";
    let quoted = false;
    for (let index = 0; index < line.length; index += 1) {
      const char = line[index];
      if (char === '"') {
        if (quoted && line[index + 1] === '"') {
          current += '"';
          index += 1;
        } else quoted = !quoted;
      } else if (char === "," && !quoted) {
        result.push(current);
        current = "";
      } else current += char;
    }
    result.push(current);
    return result;
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

  private updateFormulaBar(ref: string): void {
    const bar = this.root.querySelector<HTMLInputElement>("[data-formula-bar]");
    if (!bar) return;
    const value = this.activeSheet().cells[ref];
    bar.value = value === undefined || value === null ? "" : String(value);
  }

  private renderTabs(): void {
    const tabs = this.root.querySelector<HTMLElement>("[data-tabs]");
    if (!tabs) return;

    tabs.innerHTML = this.state.sheets
      .map((sheet, index) => `<button class="sheet-tab ${index === this.state.activeSheet ? "active" : ""}" data-sheet="${index}" type="button">${this.escape(sheet.name)}</button>`)
      .join("");

    tabs.querySelectorAll<HTMLButtonElement>("[data-sheet]").forEach((button) =>
      button.addEventListener("click", () => {
        this.state.activeSheet = Number(button.dataset.sheet);
        this.renderTabs();
        this.loadSheet();
      }),
    );
  }

  private async saveAs(): Promise<void> {
    const value = window.prompt("Save As — workbook name", this.fileName);
    if (!value?.trim()) return;
    this.fileName = value.trim().replace(/\.oqsheet$/i, "");
    await this.save();
  }

  private download(): void {
    const payload = JSON.stringify({
      metadata: { id: this.documentId, name: this.fileName, createdAt: this.createdAt, modifiedAt: new Date().toISOString() },
      content: this.state
    }, null, 2);
    const url = URL.createObjectURL(new Blob([payload], { type: "application/json" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = this.fileName.replace(/[^a-z0-9._-]+/gi, "-") + ".oqsheet";
    anchor.click();
    URL.revokeObjectURL(url);
    this.setStatus("Downloaded");
  }

  private async save(): Promise<void> {
    const now = new Date().toISOString();
    const document: OfficeDocument<SpreadsheetContent> = {
      metadata: {
        id: this.documentId,
        name: this.fileName,
        path: `/home/user/Documents/${this.fileName.replace(/[^a-z0-9._-]+/gi, "-")}.oqsheet`,
        mimeType: "application/x-oqerun-spreadsheet",
        createdAt: this.createdAt,
        modifiedAt: now,
      },
      content: this.state,
    };
    await this.store.save(document);
    const name = this.root.querySelector<HTMLElement>("[data-file-name]");
    if (name) name.textContent = this.fileName;
    this.setStatus("Saved locally");
  }

  private showFormulaHelp(): void {
    window.alert(
      "Supported formulas:\n=SUM(A1:A5)\n=AVERAGE(A1:A5)\n=MIN(A1:A5)\n=MAX(A1:A5)\n=A1+B1\n=A1*2",
    );
  }

  private parseValue(value: string): SpreadsheetCell {
    const trimmed = value.trim();
    if (trimmed.startsWith("=")) return this.evaluateFormula(trimmed.slice(1));
    if (trimmed === "") return null;
    const number = Number(trimmed);
    return Number.isFinite(number) ? number : trimmed;
  }

  private evaluateFormula(formula: string): SpreadsheetCell {
    const range = formula.match(/^(SUM|AVERAGE|MIN|MAX)\(([A-Z]+\d+):([A-Z]+\d+)\)$/i);
    if (range) {
      const start = this.refToCoord(range[2]);
      const end = this.refToCoord(range[3]);
      const values: number[] = [];
      for (let row = start.row; row <= end.row; row += 1) {
        for (let col = start.col; col <= end.col; col += 1) {
          const value = this.activeSheet().cells[this.cellRef(row, col)];
          if (typeof value === "number") values.push(value);
        }
      }
      if (!values.length) return 0;
      if (range[1].toUpperCase() === "AVERAGE") return values.reduce((a, b) => a + b, 0) / values.length;
      if (range[1].toUpperCase() === "MIN") return Math.min(...values);
      if (range[1].toUpperCase() === "MAX") return Math.max(...values);
      return values.reduce((a, b) => a + b, 0);
    }

    const cellArithmetic = formula.match(/^([A-Z]+\d+)\s*([+*\-/])\s*([A-Z]+\d+|[-]?\d+(?:\.\d+)?)$/i);
    if (cellArithmetic) {
      const left = this.activeSheet().cells[cellArithmetic[1].toUpperCase()];
      const rightToken = cellArithmetic[3];
      const rightValue = /^[A-Z]+\d+$/i.test(rightToken)
        ? this.activeSheet().cells[rightToken.toUpperCase()]
        : Number(rightToken);
      if (typeof left !== "number" || typeof rightValue !== "number") return 0;
      if (cellArithmetic[2] === "+") return left + rightValue;
      if (cellArithmetic[2] === "-") return left - rightValue;
      if (cellArithmetic[2] === "*") return left * rightValue;
      return rightValue === 0 ? 0 : left / rightValue;
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

  private escape(value: string): string {
    return value.replace(/[&<>"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[char] ?? char));
  }

  private setStatus(text: string): void {
    const status = this.root.querySelector<HTMLElement>("[data-status]");
    if (status) status.textContent = text;
  }
}
