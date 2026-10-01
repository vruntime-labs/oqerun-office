import type { OfficeDocument, SpreadsheetCell, SpreadsheetContent, SpreadsheetFormatting } from "../core/DocumentModel";
import { DocumentStore } from "../core/DocumentStore";
import { exportSheets, importSheets, downloadText } from "../core/FormatEngine";
import { FileMenu } from "../ui/FileMenu";

type SheetState = SpreadsheetContent["sheets"][number];

export class Sheets {
  private readonly root: HTMLElement;
  private readonly store = new DocumentStore();
  private readonly state: SpreadsheetContent = {
    activeSheet: 0,
    sheets: [{ name: "Sheet 1", cells: {}, formats: {} }],
  };
  private documentId: string = crypto.randomUUID();
  private createdAt = new Date().toISOString();
  private fileName = "Untitled workbook";
  private selectedRef = "A1";
  private filterTerm = "";
  private filterColumn = "";

  constructor(root: HTMLElement) {
    this.root = root;
  }

  render(): void {
    this.root.innerHTML = `
      <main class="sheet-shell">
        <header class="sheet-topbar">
          <button class="writer-back" data-action="back" type="button">← Oqerun Office</button>
          <div class="sheet-title"><span class="sheet-app-icon">S</span><div><strong>Oqerun Sheets</strong><span data-file-name>${this.escape(this.fileName)}</span></div></div>
          <span class="sheet-status" data-status>Ready</span>
        </header>
        <div class="sheet-toolbar">
          <button data-action="add-sheet" type="button">＋ Sheet</button>
          <button data-action="insert-row" type="button">＋ Row</button>
          <button data-action="insert-column" type="button">＋ Column</button>
          <button data-action="sort" type="button">Sort</button>
          <button data-action="filter" type="button">Filter</button>
          <button data-action="chart" type="button">Chart</button>
          <button data-action="format-bold" type="button"><b>B</b></button>
          <label class="sheet-color-tool">Text<input data-format-color="color" type="color" value="#202634"></label>
          <label class="sheet-color-tool">Fill<input data-format-color="background" type="color" value="#ffffff"></label>
          <button data-action="formula" type="button">ƒx Formula</button>
          <label class="sheet-open-button">Open from<input data-open-file type="file" accept=".oqsheet,.csv,.xlsx,.xls,.ods" hidden></label>
          <button data-action="save" class="primary-tool" type="button">Save</button>
        </div>
        <div class="sheet-formula-bar">
          <span class="formula-label">ƒx</span>
          <input data-formula-bar placeholder="Select a cell or enter a formula such as =SUM(A1:A5)" />
          <span class="sheet-selection" data-selection>A1</span>
        </div>
        <div class="sheet-tabs" data-tabs></div>
        <section class="sheet-grid-wrap">
          <table class="sheet-grid">
            <thead><tr><th></th>${Array.from({ length: 20 }, (_, i) => `<th>${this.columnName(i)}</th>`).join("")}</tr></thead>
            <tbody data-grid-body></tbody>
          </table>
        </section>
        <section class="sheet-chart-panel" data-chart-panel hidden>
          <div class="sheet-chart-heading"><strong>Quick chart</strong><button data-action="close-chart" type="button">×</button></div>
          <div class="sheet-chart" data-chart></div>
        </section>
      </main>`;
    this.bind();
    this.renderTabs();
    this.renderGrid();
    this.mountFileMenu();
  }

  async openPath(path: string): Promise<void> {
    const document = await this.store.load<SpreadsheetContent>(path);
    this.documentId = document.metadata.id;
    this.createdAt = document.metadata.createdAt;
    this.fileName = document.metadata.name;
    this.state.activeSheet = Math.max(0, Math.min(document.content.activeSheet, document.content.sheets.length - 1));
    this.state.sheets.splice(0, this.state.sheets.length, ...document.content.sheets.map((sheet) => ({
      ...sheet,
      formats: sheet.formats ?? {},
    })));
    this.updateFileName();
    this.renderTabs();
    this.renderGrid();
    this.setStatus("Opened");
  }

  private mountFileMenu(): void {
    new FileMenu(this.root, {
      module: "sheets",
      title: "Oqerun Sheets",
      acceptedTypes: ".oqsheet,.csv,.xlsx,.xls,.ods",
      onHome: () => this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true })),
      onNew: () => this.newWorkbook(),
      onSave: () => void this.save(),
      onSaveAs: () => void this.saveAs(),
      onDownload: () => void this.download(),
      onPrint: () => window.print(),
      onOpenRecent: (path) => void this.openPath(path),
      onOpenFile: (file) => void this.importFile(file),
    }).mount();
  }

  private bind(): void {
    this.root.querySelector('[data-action="back"]')?.addEventListener("click", () =>
      this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true })),
    );
    this.root.querySelector('[data-action="add-sheet"]')?.addEventListener("click", () => {
      this.state.sheets.push({ name: `Sheet ${this.state.sheets.length + 1}`, cells: {}, formats: {} });
      this.state.activeSheet = this.state.sheets.length - 1;
      this.renderTabs();
      this.renderGrid();
      this.setStatus("Sheet added");
    });
    this.root.querySelector('[data-action="insert-row"]')?.addEventListener("click", () => this.insertRow());
    this.root.querySelector('[data-action="insert-column"]')?.addEventListener("click", () => this.insertColumn());
    this.root.querySelector('[data-action="sort"]')?.addEventListener("click", () => this.sortSelectedColumn());
    this.root.querySelector('[data-action="filter"]')?.addEventListener("click", () => this.applyFilter());
    this.root.querySelector('[data-action="chart"]')?.addEventListener("click", () => this.renderChart());
    this.root.querySelector('[data-action="close-chart"]')?.addEventListener("click", () => {
      const panel = this.root.querySelector<HTMLElement>("[data-chart-panel]");
      if (panel) panel.hidden = true;
    });
    this.root.querySelector('[data-action="formula"]')?.addEventListener("click", () => this.showFormulaHelp());
    this.root.querySelector('[data-action="format-bold"]')?.addEventListener("click", () => this.toggleBold());
    this.root.querySelectorAll<HTMLInputElement>("[data-format-color]").forEach((input) =>
      input.addEventListener("input", () => this.applyFormat(input.dataset.formatColor as "color" | "background", input.value)),
    );
    this.root.querySelector<HTMLInputElement>("[data-open-file]")?.addEventListener("change", (event) => {
      const file = (event.currentTarget as HTMLInputElement).files?.[0];
      if (file) void this.importFile(file);
      (event.currentTarget as HTMLInputElement).value = "";
    });
    this.root.querySelector<HTMLInputElement>("[data-formula-bar]")?.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") return;
      const bar = event.currentTarget as HTMLInputElement;
      this.activeSheet().cells[this.selectedRef] = this.parseInput(bar.value);
      this.renderGrid(this.selectedRef);
      this.setStatus("Unsaved changes");
    });

    document.addEventListener("copy", (event) => {
      if (!this.root.contains(document.activeElement)) return;
      const input = document.activeElement as HTMLInputElement;
      if (!input.matches?.("[data-cell]")) return;
      event.preventDefault();
      void navigator.clipboard?.writeText(input.value);
    });
    document.addEventListener("paste", (event) => {
      if (!this.root.contains(document.activeElement)) return;
      const input = document.activeElement as HTMLInputElement;
      if (!input.matches?.("[data-cell]")) return;
      const text = event.clipboardData?.getData("text/plain");
      if (text == null) return;
      event.preventDefault();
      this.pasteText(text);
    });
  }

  private renderGrid(focusRef = this.selectedRef): void {
    const body = this.root.querySelector<HTMLElement>("[data-grid-body]");
    if (!body) return;
    const rows: string[] = [];
    for (let row = 0; row < 30; row += 1) {
      const refForRow = `A${row + 1}`;
      const firstValue = this.activeSheet().cells[refForRow];
      const hidden = this.filterTerm && this.filterColumn
        ? !String(this.activeSheet().cells[`${this.filterColumn}${row + 1}`] ?? "").toLowerCase().includes(this.filterTerm.toLowerCase())
        : false;
      rows.push(`<tr data-row="${row + 1}" ${hidden ? 'hidden' : ''}><th>${row + 1}</th>${Array.from({ length: 20 }, (_, col) => {
        const ref = this.cellRef(row, col);
        const raw = this.activeSheet().cells[ref];
        const value = raw == null ? "" : String(this.evaluateCell(ref) ?? "");
        const format = this.activeSheet().formats?.[ref] ?? {};
        const style = [format.color ? `color:${format.color}` : "", format.background ? `background:${format.background}` : "", format.align ? `text-align:${format.align}` : "", format.bold ? "font-weight:700" : "", format.italic ? "font-style:italic" : ""].filter(Boolean).join(";");
        return `<td><input data-cell="${ref}" value="${this.escapeAttribute(value)}" style="${style}" aria-label="${ref}" /></td>`;
      }).join("")}</tr>`);
    }
    body.innerHTML = rows.join("");
    body.querySelectorAll<HTMLInputElement>("[data-cell]").forEach((input) => {
      input.addEventListener("focus", () => {
        this.selectedRef = input.dataset.cell ?? "A1";
        this.updateFormulaBar(this.selectedRef);
      });
      input.addEventListener("input", () => {
        this.selectedRef = input.dataset.cell ?? "A1";
        this.activeSheet().cells[this.selectedRef] = this.parseInput(input.value);
        this.updateFormulaBar(this.selectedRef);
        this.renderGrid(this.selectedRef);
        this.setStatus("Unsaved changes");
      });
      input.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          const coord = this.refToCoord(this.selectedRef);
          const next = this.cellRef(coord.row + 1, coord.col);
          this.focusCell(next);
        }
      });
    });
    this.updateFormulaBar(focusRef);
    this.updateSelectionLabel();
    this.focusCell(focusRef, false);
  }

  public async importFile(file: File): Promise<void> {
    if (file.name.toLowerCase().endsWith(".oqsheet")) {
      try {
        const parsed = JSON.parse(await file.text()) as OfficeDocument<SpreadsheetContent>;
        this.documentId = parsed.metadata?.id ?? crypto.randomUUID();
        this.createdAt = parsed.metadata?.createdAt ?? new Date().toISOString();
        this.fileName = parsed.metadata?.name ?? file.name.replace(/\.oqsheet$/i, "");
        this.state.activeSheet = parsed.content.activeSheet;
        this.state.sheets.splice(0, this.state.sheets.length, ...parsed.content.sheets);
        this.renderTabs(); this.renderGrid(); this.setStatus("Imported workbook");
      } catch {
        this.setStatus("Could not import workbook");
      }
      return;
    }

    try {
      const imported = await importSheets(file);
      this.documentId = crypto.randomUUID();
      this.createdAt = new Date().toISOString();
      this.fileName = file.name.replace(/\.(csv|xlsx|xls|ods)$/i, "") || "Untitled workbook";
      this.state.activeSheet = imported.activeSheet;
      this.state.sheets.splice(0, this.state.sheets.length, ...imported.sheets);
      this.renderTabs(); this.renderGrid(); this.setStatus(`Imported ${file.name.split(".").pop()?.toUpperCase()}`);
      await this.save();
    } catch {
      window.alert("The selected spreadsheet could not be imported.");
    }
  }

  private newWorkbook(): void {
    this.documentId = crypto.randomUUID();
    this.createdAt = new Date().toISOString();
    this.fileName = "Untitled workbook";
    this.state.activeSheet = 0;
    this.state.sheets.splice(0, this.state.sheets.length, { name: "Sheet 1", cells: {}, formats: {} });
    this.filterTerm = "";
    this.filterColumn = "";
    this.renderTabs(); this.renderGrid(); this.updateFileName();
    this.setStatus("New workbook");
  }

  private activeSheet(): SheetState {
    return this.state.sheets[this.state.activeSheet];
  }

  private updateFormulaBar(ref: string): void {
    const bar = this.root.querySelector<HTMLInputElement>("[data-formula-bar]");
    if (!bar) return;
    const value = this.activeSheet().cells[ref];
    bar.value = value == null ? "" : String(value);
    this.selectedRef = ref;
    this.updateSelectionLabel();
  }

  private updateSelectionLabel(): void {
    const label = this.root.querySelector<HTMLElement>("[data-selection]");
    if (label) label.textContent = this.selectedRef;
  }

  private renderTabs(): void {
    const tabs = this.root.querySelector<HTMLElement>("[data-tabs]");
    if (!tabs) return;
    tabs.innerHTML = this.state.sheets.map((sheet, index) =>
      `<button class="sheet-tab ${index === this.state.activeSheet ? "active" : ""}" data-sheet="${index}" type="button">${this.escape(sheet.name)}</button>`
    ).join("");
    tabs.querySelectorAll<HTMLButtonElement>("[data-sheet]").forEach((button) =>
      button.addEventListener("click", () => {
        this.state.activeSheet = Number(button.dataset.sheet);
        this.filterTerm = ""; this.filterColumn = "";
        this.renderTabs(); this.renderGrid();
      }),
    );
  }

  private parseInput(value: string): SpreadsheetCell {
    const trimmed = value.trim();
    if (trimmed === "") return null;
    if (trimmed.startsWith("=")) return trimmed;
    if (trimmed.toLowerCase() === "true") return true;
    if (trimmed.toLowerCase() === "false") return false;
    const number = Number(trimmed);
    return Number.isFinite(number) ? number : trimmed;
  }

  private evaluateCell(ref: string, stack = new Set<string>()): SpreadsheetCell {
    const raw = this.activeSheet().cells[ref.toUpperCase()];
    if (raw == null) return null;
    if (typeof raw !== "string" || !raw.startsWith("=")) return raw;
    if (stack.has(ref.toUpperCase())) return "#CYCLE!";
    const nextStack = new Set(stack).add(ref.toUpperCase());
    return this.evaluateFormula(raw.slice(1), nextStack);
  }

  private evaluateFormula(formula: string, stack = new Set<string>()): SpreadsheetCell {
    const expression = formula.trim();
    const functionMatch = expression.match(/^([A-Z.]+)\((.*)\)$/i);
    if (functionMatch) {
      const name = functionMatch[1].toUpperCase();
      const args = this.splitArgs(functionMatch[2]);
      if (name === "IF") {
        const condition = this.evaluateCondition(args[0] ?? "", stack);
        return condition ? this.evaluateToken(args[1] ?? "", stack) : this.evaluateToken(args[2] ?? "", stack);
      }
      if (["SUM", "AVERAGE", "MIN", "MAX", "COUNT"].includes(name)) {
        const values = this.valuesFromArgs(args, stack);
        const numbers = values.filter((value): value is number => typeof value === "number");
        if (name === "COUNT") return numbers.length;
        if (!numbers.length) return 0;
        if (name === "AVERAGE") return numbers.reduce((a, b) => a + b, 0) / numbers.length;
        if (name === "MIN") return Math.min(...numbers);
        if (name === "MAX") return Math.max(...numbers);
        return numbers.reduce((a, b) => a + b, 0);
      }
      return "#NAME?";
    }

    const comparison = expression.match(/^(.+?)\s*(=|<>|>=|<=|>|<)\s*(.+)$/);
    if (comparison) return this.evaluateCondition(expression, stack);

    const arithmetic = expression.replace(/([A-Z]+\d+)/gi, (ref) => {
      const value = this.evaluateCell(ref.toUpperCase(), stack);
      return typeof value === "number" ? String(value) : "0";
    });
    if (/^[0-9+\-*/().\s]+$/.test(arithmetic)) {
      try {
        const result = Function(`"use strict"; return (${arithmetic});`)();
        return typeof result === "number" && Number.isFinite(result) ? result : 0;
      } catch {
        return "#VALUE!";
      }
    }
    return this.evaluateToken(expression, stack);
  }

  private evaluateCondition(expression: string, stack: Set<string>): boolean {
    const match = expression.match(/^(.+?)\s*(=|<>|>=|<=|>|<)\s*(.+)$/);
    if (!match) return Boolean(this.evaluateToken(expression, stack));
    const left = this.evaluateToken(match[1], stack);
    const right = this.evaluateToken(match[3], stack);
    switch (match[2]) {
      case "=": return left === right;
      case "<>": return left !== right;
      case ">": return Number(left) > Number(right);
      case "<": return Number(left) < Number(right);
      case ">=": return Number(left) >= Number(right);
      case "<=": return Number(left) <= Number(right);
      default: return false;
    }
  }

  private evaluateToken(token: string, stack: Set<string>): SpreadsheetCell {
    const value = token.trim().replace(/^"(.*)"$/, "$1");
    if (/^[A-Z]+\d+$/i.test(value)) return this.evaluateCell(value.toUpperCase(), stack);
    const number = Number(value);
    if (Number.isFinite(number) && value !== "") return number;
    if (value.toLowerCase() === "true") return true;
    if (value.toLowerCase() === "false") return false;
    return value;
  }

  private valuesFromArgs(args: string[], stack: Set<string>): SpreadsheetCell[] {
    const values: SpreadsheetCell[] = [];
    for (const arg of args) {
      const range = arg.trim().match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);
      if (range) {
        const start = this.refToCoord(range[1]); const end = this.refToCoord(range[2]);
        for (let row = start.row; row <= end.row; row += 1) {
          for (let col = start.col; col <= end.col; col += 1) values.push(this.evaluateCell(this.cellRef(row, col), stack));
        }
      } else values.push(this.evaluateToken(arg, stack));
    }
    return values;
  }

  private splitArgs(value: string): string[] {
    const result: string[] = []; let current = ""; let depth = 0; let quoted = false;
    for (const char of value) {
      if (char === '"') quoted = !quoted;
      if (!quoted && char === "(") depth += 1;
      if (!quoted && char === ")") depth -= 1;
      if (!quoted && char === "," && depth === 0) { result.push(current); current = ""; } else current += char;
    }
    result.push(current);
    return result;
  }

  private insertRow(): void {
    const row = this.refToCoord(this.selectedRef).row + 1;
    const sheet = this.activeSheet();
    const next: Record<string, SpreadsheetCell> = {};
    const nextFormats: Record<string, SpreadsheetFormatting> = {};
    Object.entries(sheet.cells).forEach(([ref, value]) => {
      const coord = this.refToCoord(ref);
      next[this.cellRef(coord.row >= row ? coord.row + 1 : coord.row, coord.col)] = value;
    });
    Object.entries(sheet.formats ?? {}).forEach(([ref, value]) => {
      const coord = this.refToCoord(ref);
      nextFormats[this.cellRef(coord.row >= row ? coord.row + 1 : coord.row, coord.col)] = value;
    });
    sheet.cells = next; sheet.formats = nextFormats;
    this.renderGrid(this.selectedRef); this.setStatus("Row inserted");
  }

  private insertColumn(): void {
    const col = this.refToCoord(this.selectedRef).col + 1;
    const sheet = this.activeSheet();
    const next: Record<string, SpreadsheetCell> = {};
    const nextFormats: Record<string, SpreadsheetFormatting> = {};
    Object.entries(sheet.cells).forEach(([ref, value]) => {
      const coord = this.refToCoord(ref);
      next[this.cellRef(coord.row, coord.col >= col ? coord.col + 1 : coord.col)] = value;
    });
    Object.entries(sheet.formats ?? {}).forEach(([ref, value]) => {
      const coord = this.refToCoord(ref);
      nextFormats[this.cellRef(coord.row, coord.col >= col ? coord.col + 1 : coord.col)] = value;
    });
    sheet.cells = next; sheet.formats = nextFormats;
    this.renderGrid(this.selectedRef); this.setStatus("Column inserted");
  }

  private sortSelectedColumn(): void {
    const coord = this.refToCoord(this.selectedRef);
    const sheet = this.activeSheet();
    const rows = Array.from({ length: 30 }, (_, index) => index);
    rows.sort((a, b) => {
      const av = this.evaluateCell(this.cellRef(a, coord.col));
      const bv = this.evaluateCell(this.cellRef(b, coord.col));
      return String(av ?? "").localeCompare(String(bv ?? ""), undefined, { numeric: true });
    });
    const next: Record<string, SpreadsheetCell> = {};
    rows.forEach((sourceRow, targetRow) => {
      for (let col = 0; col < 20; col += 1) {
        const value = sheet.cells[this.cellRef(sourceRow, col)];
        if (value !== undefined) next[this.cellRef(targetRow, col)] = value;
      }
    });
    sheet.cells = next;
    this.renderGrid(this.selectedRef);
    this.setStatus("Sorted current sheet");
  }

  private applyFilter(): void {
    const column = window.prompt("Filter column", this.refToCoord(this.selectedRef).col >= 0 ? this.columnName(this.refToCoord(this.selectedRef).col) : "A");
    if (!column) return;
    const term = window.prompt(`Show rows where ${column} contains`, this.filterTerm);
    if (term === null) return;
    this.filterColumn = column.toUpperCase();
    this.filterTerm = term;
    this.renderGrid(this.selectedRef);
    this.setStatus(term ? "Filter applied" : "Filter cleared");
  }

  private toggleBold(): void {
    const format = this.activeSheet().formats ?? (this.activeSheet().formats = {});
    const current = format[this.selectedRef] ?? {};
    format[this.selectedRef] = { ...current, bold: !current.bold };
    this.renderGrid(this.selectedRef);
    this.setStatus("Cell formatting changed");
  }

  private applyFormat(kind: "color" | "background", value: string): void {
    const format = this.activeSheet().formats ?? (this.activeSheet().formats = {});
    format[this.selectedRef] = { ...(format[this.selectedRef] ?? {}), [kind]: value };
    this.renderGrid(this.selectedRef);
    this.setStatus("Cell formatting changed");
  }

  private pasteText(text: string): void {
    const start = this.refToCoord(this.selectedRef);
    text.split(/\r?\n/).forEach((line, rowOffset) => {
      line.split("\t").forEach((value, colOffset) => {
        this.activeSheet().cells[this.cellRef(start.row + rowOffset, start.col + colOffset)] = this.parseInput(value);
      });
    });
    this.renderGrid(this.selectedRef);
    this.setStatus("Pasted");
  }

  private focusCell(ref: string, focus = true): void {
    this.selectedRef = ref;
    const input = this.root.querySelector<HTMLInputElement>(`[data-cell="${CSS.escape(ref)}"]`);
    if (input && focus) input.focus();
    this.updateFormulaBar(ref);
  }

  private renderChart(): void {
    const panel = this.root.querySelector<HTMLElement>("[data-chart-panel]");
    const chart = this.root.querySelector<HTMLElement>("[data-chart]");
    if (!panel || !chart) return;
    const coord = this.refToCoord(this.selectedRef);
    const values: Array<{ label: string; value: number }> = [];
    for (let row = 0; row < 15; row += 1) {
      const label = String(this.activeSheet().cells[`A${row + 1}`] ?? `Row ${row + 1}`);
      const value = this.evaluateCell(this.cellRef(row, coord.col));
      if (typeof value === "number") values.push({ label, value });
    }
    if (!values.length) {
      window.alert("Select a column containing numeric values to create a chart.");
      return;
    }
    const max = Math.max(...values.map((item) => Math.abs(item.value)), 1);
    chart.innerHTML = values.slice(0, 10).map((item) =>
      `<div class="sheet-chart-row"><span>${this.escape(item.label)}</span><div><i style="width:${Math.max(3, Math.abs(item.value) / max * 100)}%"></i></div><b>${item.value}</b></div>`
    ).join("");
    panel.hidden = false;
  }

  private async saveAs(): Promise<void> {
    const value = window.prompt("Save As — workbook name", this.fileName);
    if (!value?.trim()) return;
    this.fileName = value.trim().replace(/\.oqsheet$/i, "");
    await this.save();
  }

  private async download(): Promise<void> {
    const choice = (window.prompt("Download format:\n1. XLSX\n2. XLS\n3. ODS\n4. CSV\n5. Oqerun .oqsheet", "1") ?? "1").trim();
    if (choice === "1" || choice === "2" || choice === "3" || choice === "4") {
      const formats = { "1": "xlsx", "2": "xls", "3": "ods", "4": "csv" } as const;
      exportSheets(this.state, this.fileName, formats[choice as keyof typeof formats]);
      this.setStatus("Spreadsheet downloaded");
      return;
    }
    const payload = JSON.stringify(this.currentDocument(), null, 2);
    downloadText(payload, this.safeName(this.fileName) + ".oqsheet", "application/json");
    this.setStatus("Oqerun workbook downloaded");
  }

  private currentDocument(): OfficeDocument<SpreadsheetContent> {
    return {
      metadata: {
        id: this.documentId,
        name: this.fileName,
        path: `/home/user/Documents/${this.safeName(this.fileName)}.oqsheet`,
        mimeType: "application/x-oqerun-spreadsheet",
        createdAt: this.createdAt,
        modifiedAt: new Date().toISOString(),
      },
      content: this.state,
    };
  }

  private async save(): Promise<void> {
    await this.store.save(this.currentDocument());
    this.updateFileName();
    this.setStatus("Saved locally");
  }

  private showFormulaHelp(): void {
    window.alert(`Supported formulas:\n=SUM(A1:A5)\n=AVERAGE(A1:A5)\n=MIN(A1:A5)\n=MAX(A1:A5)\n=COUNT(A1:A5)\n=IF(A1>10,"High","Low")\n=A1+B1*2`);
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
    let n = index + 1; let result = "";
    while (n > 0) { const rem = (n - 1) % 26; result = String.fromCharCode(65 + rem) + result; n = Math.floor((n - 1) / 26); }
    return result;
  }

  private updateFileName(): void {
    const name = this.root.querySelector<HTMLElement>("[data-file-name]");
    if (name) name.textContent = this.fileName;
  }

  private safeName(value: string): string {
    return value.trim().replace(/[^a-z0-9._-]+/gi, "-").replace(/^-+|-+$/g, "") || "untitled";
  }

  private escape(value: string): string {
    return value.replace(/[&<>"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[char] ?? char));
  }

  private escapeAttribute(value: string): string {
    return this.escape(value).replace(/'/g, "&#39;");
  }

  private setStatus(text: string): void {
    const status = this.root.querySelector<HTMLElement>("[data-status]");
    if (status) status.textContent = text;
  }
}
