import { getRecent, type RecentModule } from "../core/RecentDocuments";

export interface FileMenuOptions {
  module: RecentModule;
  title: string;
  acceptedTypes: string;
  onHome: () => void;
  onNew: () => void;
  onSave: () => void;
  onOpenRecent: (path: string) => void;
  onOpenFile: (file: File) => void;
}

export class FileMenu {
  private readonly root: HTMLElement;
  private readonly options: FileMenuOptions;

  constructor(root: HTMLElement, options: FileMenuOptions) {
    this.root = root;
    this.options = options;
  }

  mount(): void {
    const topbar = this.root.querySelector<HTMLElement>(
      ".writer-topbar, .sheet-topbar, .slides-topbar",
    );
    if (!topbar) return;

    topbar.querySelector("[data-file-menu-button]")?.remove();

    const button = document.createElement("button");
    button.className = "office-file-button";
    button.type = "button";
    button.dataset.fileMenuButton = "true";
    button.textContent = "File";
    topbar.insertBefore(button, topbar.firstChild);

    const panel = document.createElement("section");
    panel.className = "office-file-menu";
    panel.hidden = true;
    panel.innerHTML = \`
      <div class="office-file-backdrop" data-file-menu-close></div>
      <div class="office-file-dialog" role="dialog" aria-label="\${this.escape(this.options.title)} File menu">
        <aside class="office-file-sidebar">
          <div class="office-file-sidebar-title">\${this.escape(this.options.title)}</div>
          <button data-file-action="home" type="button">⌂ Home</button>
          <button data-file-action="new" type="button">＋ New</button>
          <button data-file-action="open" type="button">↥ Open From</button>
          <button data-file-action="recent" type="button">◷ Open Recent</button>
          <button data-file-action="save" type="button">↓ Save</button>
          <button data-file-menu-close type="button">← Back to editor</button>
        </aside>
        <div class="office-file-content">
          <div class="office-file-heading">
            <div>
              <p class="eyebrow">File</p>
              <h2>Open Recent</h2>
              <p>Files saved or opened in \${this.escape(this.options.title)} appear here automatically.</p>
            </div>
          </div>
          <div class="office-file-recent" data-recent-list></div>
          <input data-file-menu-input type="file" accept="\${this.escape(this.options.acceptedTypes)}" hidden>
        </div>
      </div>\`;
    
    const fileInput = panel.querySelector<HTMLInputElement>("[data-file-menu-input]");

    button.addEventListener("click", () => {
      panel.hidden = !panel.hidden;
      if (!panel.hidden) this.refreshRecent(panel);
    });

    panel.querySelectorAll("[data-file-menu-close]").forEach((close) =>
      close.addEventListener("click", () => { panel.hidden = true; }),
    );

    panel.querySelector('[data-file-action="home"]')?.addEventListener("click", () => {
      panel.hidden = true;
      this.options.onHome();
    });

    panel.querySelector('[data-file-action="new"]')?.addEventListener("click", () => {
      panel.hidden = true;
      this.options.onNew();
    });

    panel.querySelector('[data-file-action="save"]')?.addEventListener("click", () => {
      panel.hidden = true;
      this.options.onSave();
    });

    panel.querySelector('[data-file-action="open"]')?.addEventListener("click", () => {
      fileInput?.click();
    });

    fileInput?.addEventListener("change", () => {
      const file = fileInput.files?.[0];
      fileInput.value = "";
      if (!file) return;
      panel.hidden = true;
      this.options.onOpenFile(file);
    });

    panel.querySelector('[data-file-action="recent"]')?.addEventListener("click", () => {
      this.refreshRecent(panel);
      panel.querySelector<HTMLElement>("[data-recent-list]")?.scrollIntoView({ behavior: "smooth", block: "nearest" });
    });

    document.addEventListener("keydown", (event) => {
      if ((event.altKey && event.key.toLowerCase() === "f") || (event.ctrlKey && event.key.toLowerCase() === "o")) {
        event.preventDefault();
        panel.hidden = false;
        this.refreshRecent(panel);
      }
      if (event.key === "Escape" && !panel.hidden) panel.hidden = true;
    });
  }

  private refreshRecent(panel: HTMLElement): void {
    const list = panel.querySelector<HTMLElement>("[data-recent-list]");
    if (!list) return;

    const recent = getRecent().filter((item) => item.module === this.options.module);
    if (!recent.length) {
      list.innerHTML = \`
        <div class="office-file-empty">
          <strong>No recent files yet</strong>
          <span>Save or open a \${this.escape(this.options.title)} file and it will appear here.</span>
          <button data-file-action="open-empty" type="button">Open From</button>
        </div>\`;
      list.querySelector('[data-file-action="open-empty"]')?.addEventListener("click", () => {
        panel.querySelector<HTMLInputElement>("[data-file-menu-input]")?.click();
      });
      return;
    }

    list.innerHTML = recent.map((item) => \`
      <button class="office-recent-file" data-recent-path="\${this.escapeAttribute(item.path)}" type="button">
        <span class="office-recent-file-icon">\${this.iconFor(item.module)}</span>
        <span class="office-recent-file-copy">
          <strong>\${this.escape(item.name)}</strong>
          <small>\${new Date(item.modifiedAt).toLocaleString()}</small>
        </span>
        <span class="office-recent-arrow">→</span>
      </button>\`).join("");

    list.querySelectorAll<HTMLButtonElement>("[data-recent-path]").forEach((recentButton) => {
      recentButton.addEventListener("click", () => {
        const path = recentButton.dataset.recentPath;
        if (!path) return;
        panel.hidden = true;
        this.options.onOpenRecent(path);
      });
    });
  }

  private iconFor(module: RecentModule): string {
    return module === "writer" ? "W" : module === "sheets" ? "S" : "P";
  }

  private escape(value: string): string {
    return value.replace(/[&<>"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[char] ?? char));
  }

  private escapeAttribute(value: string): string {
    return this.escape(value).replace(/'/g, "&#39;");
  }
}
