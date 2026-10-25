import { Writer } from "../writer/Writer";
import { Sheets } from "../sheets/Sheets";
import { Slides } from "../slides/Slides";
import { DocumentStore } from "./DocumentStore";
import { getRecent, type RecentDocument } from "./RecentDocuments";
import type { OfficeDocument, SpreadsheetContent, WriterContent } from "./DocumentModel";

export type OfficeModule = "writer" | "sheets" | "slides";
type EditorInstance = Writer | Sheets | Slides;

interface WorkspaceTab {
  id: string;
  module: OfficeModule;
  name: string;
  host: HTMLElement;
  instance: EditorInstance;
}

export class Office {
  private readonly root: HTMLElement | null;
  private readonly store = new DocumentStore();
  private readonly tabs = new Map<string, WorkspaceTab>();
  private activeTabId: string | null = null;
  private homeHost: HTMLElement | null = null;
  private workspaceHost: HTMLElement | null = null;
  private tabStrip: HTMLElement | null = null;

  constructor(root: HTMLElement | null) {
    this.root = root;
  }

  start(): void {
    const root = this.root;
    if (!root) throw new Error("Oqerun Office root element was not found.");

    root.innerHTML = `
      <main class="office-app-shell">
        <header class="office-global-bar">
          <button class="office-global-brand" data-global-home type="button">
            <span class="office-global-logo">◈</span>
            <span><strong>Oqerun Office</strong><small>VRuntime Labs</small></span>
          </button>
          <div class="office-global-tabs" data-workspace-tabs>
            <button class="office-home-tab is-active" data-global-home-tab type="button">Home</button>
          </div>
          <div class="office-global-actions">
            <label class="office-global-open" title="Open file">
              <span>Open</span>
              <input data-global-open type="file" accept=".oqdoc,.oqsheet,.oqslide,.docx,.xlsx,.xls,.ods,.pptx,.txt,.csv,.html" hidden>
            </label>
            <button type="button" class="office-global-icon" title="Comments">💬</button>
            <span class="office-presence" title="Workspace presence"><i></i><i></i><b>V</b></span>
          </div>
        </header>
        <section class="office-home-host" data-home-host></section>
        <section class="office-workspace-host" data-workspace-host hidden></section>
      </main>`;

    this.homeHost = root.querySelector<HTMLElement>("[data-home-host]");
    this.workspaceHost = root.querySelector<HTMLElement>("[data-workspace-host]");
    this.tabStrip = root.querySelector<HTMLElement>("[data-workspace-tabs]");

    root.querySelectorAll("[data-global-home], [data-global-home-tab]").forEach((button) =>
      button.addEventListener("click", () => this.showHome()),
    );

    root.querySelector<HTMLInputElement>("[data-global-open]")?.addEventListener("change", (event) => {
      const input = event.currentTarget as HTMLInputElement;
      const file = input.files?.[0];
      if (file) void this.openFile(file);
      input.value = "";
    });

    root.addEventListener("office:navigate", this.handleNavigation);
    this.renderHome();
  }

  private handleNavigation = (event: Event): void => {
    const detail = (event as CustomEvent<string>).detail;
    if (detail === "home") this.showHome();
  };

  private showHome(): void {
    if (this.homeHost) this.homeHost.hidden = false;
    if (this.workspaceHost) this.workspaceHost.hidden = true;
    this.setHomeTabActive(true);
  }

  private showWorkspace(): void {
    if (this.homeHost) this.homeHost.hidden = true;
    if (this.workspaceHost) this.workspaceHost.hidden = false;
    this.setHomeTabActive(false);
  }

  private renderHome(): void {
    const host = this.homeHost;
    if (!host) return;
    const recent = getRecent();

    host.innerHTML = `
      <main class="office-shell">
        <header class="office-home-topbar">
          <div class="office-home-brand">
            <span class="office-home-logo">◈</span>
            <div><strong>Oqerun Office</strong><span>Open-source productivity suite</span></div>
          </div>
          <div class="office-home-actions"><span class="office-version">0.3.0</span></div>
        </header>
        <section class="office-home">
          <div class="office-home-hero">
            <div>
              <p class="eyebrow">VRuntime Labs</p>
              <h1>Everything you need to create.</h1>
              <p>Documents, spreadsheets and presentations in one focused workspace — designed for Oqerun OS and built in the open.</p>
            </div>
            <div class="office-home-orbit" aria-hidden="true"><span>W</span><span>S</span><span>P</span></div>
          </div>
          <div class="office-module-grid">
            ${this.moduleCard("writer", "Writer", "Write, format and publish polished documents.", "W")}
            ${this.moduleCard("sheets", "Sheets", "Calculate, analyze and visualize your data.", "S")}
            ${this.moduleCard("slides", "Slides", "Build presentations with layouts, themes and media.", "P")}
          </div>
          <section class="office-recent-section">
            <div class="office-section-heading">
              <div><p class="eyebrow">Your workspace</p><h2>Recent files</h2></div>
              ${recent.length ? '<button data-action="clear-recent" class="text-button" type="button">Clear history</button>' : ""}
            </div>
            <div class="office-recent-grid">
              ${recent.length ? recent.map((item) => this.recentCard(item)).join("") : '<div class="office-empty-state"><strong>No recent files yet</strong><span>Create a document, workbook or presentation and it will appear here.</span></div>'}
            </div>
          </section>
        </section>
        <footer class="office-footer"><span>Open source · MIT License</span><span>VRuntime Labs · Oqerun Office</span><span class="office-legal-links"><a href="./privacy.html" target="_blank" rel="noopener noreferrer">Privacy Policy</a><a href="./terms.html" target="_blank" rel="noopener noreferrer">Terms</a></span></footer>
      </main>`;

    host.querySelectorAll<HTMLButtonElement>("[data-module]").forEach((button) =>
      button.addEventListener("click", () => this.openModule(button.dataset.module as OfficeModule)),
    );

    host.querySelectorAll<HTMLButtonElement>("[data-recent-path]").forEach((button) =>
      button.addEventListener("click", () => {
        const item = recent.find((entry) => entry.path === button.dataset.recentPath);
        if (item) this.openRecent(item);
      }),
    );

    host.querySelector('[data-action="clear-recent"]')?.addEventListener("click", () => {
      localStorage.removeItem("oqerun-office-recent");
      this.renderHome();
    });
  }

  private openModule(module: OfficeModule): void {
    const id = crypto.randomUUID();
    const name = module === "writer" ? "Untitled document" : module === "sheets" ? "Untitled spreadsheet" : "Untitled presentation";
    const host = document.createElement("div");
    host.className = "office-document-host";
    host.dataset.tabId = id;
    host.hidden = true;
    this.workspaceHost?.appendChild(host);

    const instance = module === "writer" ? new Writer(host) : module === "sheets" ? new Sheets(host) : new Slides(host);
    instance.render();

    this.tabs.set(id, { id, module, name, host, instance });
    this.activeTabId = id;
    this.showWorkspace();
    this.renderWorkspaceTabs();
    this.activateTab(id);
  }

  private activateTab(id: string): void {
    const tab = this.tabs.get(id);
    if (!tab) return;
    this.activeTabId = id;
    this.tabs.forEach((entry) => {
      entry.host.hidden = entry.id !== id;
    });
    this.renderWorkspaceTabs();
    this.showWorkspace();
  }

  private closeTab(id: string): void {
    const tab = this.tabs.get(id);
    if (!tab) return;
    tab.host.remove();
    this.tabs.delete(id);
    if (this.activeTabId === id) {
      const remaining = Array.from(this.tabs.values());
      this.activeTabId = remaining.at(-1)?.id ?? null;
    }
    if (this.activeTabId) this.activateTab(this.activeTabId);
    else this.showHome();
    this.renderWorkspaceTabs();
  }

  private renderWorkspaceTabs(): void {
    const strip = this.tabStrip;
    if (!strip) return;

    strip.innerHTML = `
      <button class="office-home-tab ${this.activeTabId ? "" : "is-active"}" data-global-home-tab type="button">Home</button>
      ${Array.from(this.tabs.values()).map((tab) => `
        <div class="office-document-tab ${tab.id === this.activeTabId ? "is-active" : ""}" data-tab-id="${tab.id}">
          <button class="office-document-tab-main" type="button">
            <span class="office-tab-icon">${tab.module === "writer" ? "W" : tab.module === "sheets" ? "S" : "P"}</span>
            <span class="office-tab-name">${this.escape(tab.name)}</span>
          </button>
          <button class="office-document-tab-close" type="button" aria-label="Close document">×</button>
        </div>
      `).join("")}
      <button class="office-new-tab" type="button" title="New document">+</button>`;

    strip.querySelector("[data-global-home-tab]")?.addEventListener("click", () => this.showHome());
    strip.querySelector(".office-new-tab")?.addEventListener("click", () => this.openModule("writer"));

    strip.querySelectorAll<HTMLElement>(".office-document-tab").forEach((element) => {
      const id = element.dataset.tabId;
      if (!id) return;
      element.querySelector(".office-document-tab-main")?.addEventListener("click", () => this.activateTab(id));
      element.querySelector(".office-document-tab-close")?.addEventListener("click", (event) => {
        event.stopPropagation();
        this.closeTab(id);
      });
    });
  }

  private setHomeTabActive(active: boolean): void {
    this.tabStrip?.querySelector(".office-home-tab")?.classList.toggle("is-active", active);
  }

  private async openFile(file: File): Promise<void> {
    const extension = file.name.toLowerCase().slice(file.name.lastIndexOf("."));
    const module: OfficeModule =
      [".oqsheet", ".csv", ".xlsx", ".xls", ".ods"].includes(extension) ? "sheets" :
      [".oqslide", ".pptx"].includes(extension) ? "slides" : "writer";

    const supported = [".oqdoc", ".oqsheet", ".oqslide", ".docx", ".xlsx", ".xls", ".ods", ".pptx", ".txt", ".csv", ".html"];
    if (!supported.includes(extension)) {
      window.alert("This file format is not supported by the current import engine yet.");
      return;
    }

    const id = crypto.randomUUID();
    const host = document.createElement("div");
    host.className = "office-document-host";
    host.dataset.tabId = id;
    host.hidden = true;
    this.workspaceHost?.appendChild(host);
    const instance = module === "writer" ? new Writer(host) : module === "sheets" ? new Sheets(host) : new Slides(host);
    instance.render();
    this.tabs.set(id, { id, module, name: file.name, host, instance });
    this.activeTabId = id;
    this.showWorkspace();
    this.renderWorkspaceTabs();
    this.activateTab(id);

    if ([".docx", ".xlsx", ".xls", ".ods", ".pptx"].includes(extension)) {
      await instance.importFile(file);
      this.updateTabName(id, file.name);
      return;
    }

    const text = await file.text();

    if (extension === ".csv") {
      const cells: Record<string, string> = {};
      text.split(/\r?\n/).forEach((line, row) => line.split(",").forEach((value, col) => {
        cells[`${this.columnName(col)}${row + 1}`] = value.trim();
      }));
      const document: OfficeDocument<SpreadsheetContent> = {
        metadata: {
          id: crypto.randomUUID(), name: file.name.replace(/\.csv$/i, ""),
          path: `/home/user/Documents/${this.safeName(file.name.replace(/\.csv$/i, ""))}.oqsheet`,
          mimeType: "application/x-oqerun-spreadsheet", createdAt: new Date().toISOString(), modifiedAt: new Date().toISOString(),
        },
        content: { activeSheet: 0, sheets: [{ name: "Sheet 1", cells, formats: {} }] },
      };
      const path = await this.store.save(document);
      await (instance as Sheets).openPath(path);
      this.updateTabName(id, file.name.replace(/\.csv$/i, ""));
      return;
    }

    if (extension === ".txt" || extension === ".html") {
      const document: OfficeDocument<WriterContent> = {
        metadata: {
          id: crypto.randomUUID(), name: file.name.replace(/\.(txt|html)$/i, ""),
          path: `/home/user/Documents/${this.safeName(file.name.replace(/\.(txt|html)$/i, ""))}.oqdoc`,
          mimeType: "application/x-oqerun-document", createdAt: new Date().toISOString(), modifiedAt: new Date().toISOString(),
        },
        content: {
          html: extension === ".html" ? text : `<p>${this.escape(text).replace(/\r?\n/g, "</p><p>")}</p>`,
          plainText: text, blocks: [],
        },
      };
      const path = await this.store.save(document);
      await (instance as Writer).openPath(path);
      this.updateTabName(id, file.name.replace(/\.(txt|html)$/i, ""));
      return;
    }

    try {
      const parsed = JSON.parse(text) as OfficeDocument;
      if (!parsed?.content || !parsed.metadata) throw new Error("Invalid Office document");
      const path = await this.store.save(parsed);
      await (instance as EditorInstance).openPath(path);
      this.updateTabName(id, file.name);
    } catch {
      window.alert("The selected Oqerun Office file could not be read.");
    }
  }

  private openRecent(item: RecentDocument): void {
    const id = crypto.randomUUID();
    const host = document.createElement("div");
    host.className = "office-document-host";
    host.dataset.tabId = id;
    host.hidden = true;
    this.workspaceHost?.appendChild(host);
    const instance = item.module === "writer" ? new Writer(host) : item.module === "sheets" ? new Sheets(host) : new Slides(host);
    instance.render();
    this.tabs.set(id, { id, module: item.module, name: item.name, host, instance });
    this.activeTabId = id;
    this.showWorkspace();
    this.renderWorkspaceTabs();
    this.activateTab(id);
    void instance.openPath(item.path);
  }

  private updateTabName(id: string, name: string): void {
    const tab = this.tabs.get(id);
    if (!tab) return;
    tab.name = name;
    this.renderWorkspaceTabs();
  }

  private moduleCard(module: OfficeModule, title: string, description: string, glyph: string): string {
    return `<button class="module-card" data-module="${module}" type="button">
      <span class="module-icon">${glyph}</span>
      <span class="module-content"><strong>Oqerun ${title}</strong><span>${description}</span></span>
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
