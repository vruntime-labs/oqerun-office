import { Writer } from "../writer/Writer";
import { Sheets } from "../sheets/Sheets";
import { Slides } from "../slides/Slides";
import { DocumentStore } from "./DocumentStore";
import { getRecent, type RecentDocument } from "./RecentDocuments";
import type { OfficeDocument, SpreadsheetContent, WriterContent } from "./DocumentModel";

export type OfficeModule = "writer" | "sheets" | "slides";

export class Office {
  private readonly root: HTMLElement | null;
  private readonly store = new DocumentStore();

  constructor(root: HTMLElement | null) {
    this.root = root;
  }

  start(): void {
    const root = this.root;
    if (!root) throw new Error("Oqerun Office root element was not found.");
    this.renderHome();
    root.addEventListener("office:navigate", this.handleNavigation);
  }

  private handleNavigation = (event: Event): void => {
    const detail = (event as CustomEvent<string>).detail;
    if (detail === "home") this.renderHome();
  };

  private renderHome(): void {
    const root = this.root;
    if (!root) return;

    const recent = getRecent();

    root.innerHTML = `
      <main class="office-shell">
        <header class="office-header">
          <div>
            <div class="office-brand">Oqerun Office</div>
            <div class="office-subtitle">Office Suite</div>
          </div>
          <div class="office-header-actions">
            <label class="home-open-button">
              <span>Open from</span>
              <input data-open-file type="file" accept=".oqdoc,.oqsheet,.oqslide,.docx,.xlsx,.xls,.ods,.pptx,.txt,.csv,.html" hidden>
            </label>
            <div class="office-version">0.3.0</div>
          </div>
        </header>

        <section class="office-welcome">
          <p class="eyebrow">VRuntime Labs</p>
          <h1>One suite. Three workspaces.</h1>
          <p class="welcome-copy">Create documents, spreadsheets and presentations from one unified Oqerun application.</p>

          <div class="module-grid">
            ${this.moduleCard("writer", "Writer", "Documents, rich text, tables, images and equations.", "W")}
            ${this.moduleCard("sheets", "Sheets", "Spreadsheets, formulas, tables and multiple sheets.", "S")}
            ${this.moduleCard("slides", "Slides", "Presentations, layouts, themes, images and shapes.", "P")}
          </div>

          <section class="recent-section">
            <div class="section-heading">
              <div>
                <p class="eyebrow">Workspace</p>
                <h2>Open recent</h2>
              </div>
              ${recent.length ? '<button data-action="clear-recent" class="text-button" type="button">Clear</button>' : ""}
            </div>
            <div class="recent-list">
              ${recent.length ? recent.map((item) => this.recentCard(item)).join("") : '<div class="empty-recent">Your recently saved documents will appear here.</div>'}
            </div>
          </section>
        </section>

        <footer class="office-footer">
          <span>Open source • Designed for Oqerun OS</span>
          <span>VRuntime Labs</span>
        </footer>
      </main>`;

    root.querySelectorAll<HTMLButtonElement>("[data-module]").forEach((button) =>
      button.addEventListener("click", () => this.openModule(button.dataset.module as OfficeModule)),
    );

    root.querySelectorAll<HTMLButtonElement>("[data-recent-path]").forEach((button) =>
      button.addEventListener("click", () => {
        const item = recent.find((entry) => entry.path === button.dataset.recentPath);
        if (item) this.openRecent(item);
      }),
    );

    root.querySelector('[data-action="clear-recent"]')?.addEventListener("click", () => {
      localStorage.removeItem("oqerun-office-recent");
      this.renderHome();
    });

    root.querySelector<HTMLInputElement>("[data-open-file]")?.addEventListener("change", (event) => {
      const input = event.currentTarget as HTMLInputElement;
      const file = input.files?.[0];
      if (file) void this.openFile(file);
      input.value = "";
    });
  }

  private async openFile(file: File): Promise<void> {
    const extension = file.name.toLowerCase().slice(file.name.lastIndexOf("."));
    const textExtensions = [".txt", ".csv", ".html"];
    let module: OfficeModule | null =
      extension === ".oqsheet" || extension === ".csv" ? "sheets" :
      extension === ".oqslide" ? "slides" :
      "writer";

    if (![".oqdoc", ".oqsheet", ".oqslide", ".docx", ".xlsx", ".xls", ".ods", ".pptx", ".txt", ".csv", ".html"].includes(extension)) {
      window.alert("This file format is not supported by the current import engine yet.");
      return;
    }

    const store = this.store;
    const text = await file.text();

    if (textExtensions.includes(extension)) {
      if (extension === ".csv") {
        const cells: Record<string, string> = {};
        text.split(/\r?\n/).forEach((line, row) => {
          line.split(",").forEach((value, col) => {
            cells[`${this.columnName(col)}${row + 1}`] = value.trim();
          });
        });
        const document: OfficeDocument<SpreadsheetContent> = {
          metadata: {
            id: crypto.randomUUID(),
            name: file.name.replace(/\.csv$/i, ""),
            path: `/home/user/Documents/${this.safeName(file.name.replace(/\.csv$/i, ""))}.oqsheet`,
            mimeType: "application/x-oqerun-spreadsheet",
            createdAt: new Date().toISOString(),
            modifiedAt: new Date().toISOString(),
          },
          content: { activeSheet: 0, sheets: [{ name: "Sheet 1", cells }] },
        };
        const path = await store.save(document);
        this.openModule("sheets");
        window.setTimeout(() => void new Sheets(this.root!).openPath(path), 0);
        return;
      }

      const document: OfficeDocument<WriterContent> = {
        metadata: {
          id: crypto.randomUUID(),
          name: file.name.replace(/\.(txt|html)$/i, ""),
          path: `/home/user/Documents/${this.safeName(file.name.replace(/\.(txt|html)$/i, ""))}.oqdoc`,
          mimeType: "application/x-oqerun-document",
          createdAt: new Date().toISOString(),
          modifiedAt: new Date().toISOString(),
        },
        content: { html: `<p>${this.escape(text).replace(/\r?\n/g, "</p><p>")}</p>`, plainText: text, blocks: [] },
      };
      const path = await store.save(document);
      this.openModule("writer");
      window.setTimeout(() => void new Writer(this.root!).openPath(path), 0);
      return;
    }

    try {
      const parsed = JSON.parse(text) as OfficeDocument;
      if (parsed?.metadata?.path && parsed?.content) {
        const path = await store.save(parsed);
        module = extension === ".oqsheet" ? "sheets" : extension === ".oqslide" ? "slides" : "writer";
        this.openModule(module);
        window.setTimeout(() => {
          if (module === "writer") void new Writer(this.root!).openPath(path);
          else if (module === "sheets") void new Sheets(this.root!).openPath(path);
          else void new Slides(this.root!).openPath(path);
        }, 0);
      }
    } catch {
      window.alert("The selected Office file could not be read.");
    }
  }

  private openRecent(item: RecentDocument): void {
    this.openModule(item.module);
    window.setTimeout(() => {
      if (item.module === "writer") void new Writer(this.root!).openPath(item.path);
      else if (item.module === "sheets") void new Sheets(this.root!).openPath(item.path);
      else void new Slides(this.root!).openPath(item.path);
    }, 0);
  }

  private openModule(module: OfficeModule): void {
    if (module === "writer") new Writer(this.root!).render();
    else if (module === "sheets") new Sheets(this.root!).render();
    else new Slides(this.root!).render();
  }

  private moduleCard(module: OfficeModule, title: string, description: string, glyph: string): string {
    return `<button class="module-card" data-module="${module}" type="button">
      <span class="module-icon">${glyph}</span>
      <span class="module-content">
        <strong>Oqerun ${title}</strong>
        <span>${description}</span>
      </span>
      <span class="module-arrow">→</span>
    </button>`;
  }

  private recentCard(item: RecentDocument): string {
    const icon = item.module === "writer" ? "W" : item.module === "sheets" ? "S" : "P";
    const date = new Date(item.modifiedAt).toLocaleString();
    return `<button class="recent-card" data-recent-path="${this.escape(item.path)}" type="button">
      <span class="recent-icon">${icon}</span>
      <span class="recent-content"><strong>${this.escape(item.name)}</strong><span>${item.module} • ${date}</span></span>
      <span class="module-arrow">→</span>
    </button>`;
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

  private safeName(value: string): string {
    return value.trim().replace(/[^a-z0-9._-]+/gi, "-").replace(/^-+|-+$/g, "") || "untitled";
  }

  private escape(value: string): string {
    return value.replace(/[&<>"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[char] ?? char));
  }
}
