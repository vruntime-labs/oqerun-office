import { clearRecent, getRecent, removeRecent, togglePinned, type RecentModule } from "../core/RecentDocuments";

export interface FileMenuOptions {
  module: RecentModule;
  title: string;
  acceptedTypes: string;
  onHome: () => void;
  onNew: () => void;
  onSave: () => void;
  onSaveAs: () => void;
  onDownload: () => void;
  onPrint: () => void;
  onOpenRecent: (path: string) => void;
  onOpenFile: (file: File) => void;
}

export class FileMenu {
  private readonly root: HTMLElement;
  private readonly options: FileMenuOptions;
  private panel: HTMLElement | null = null;

  constructor(root: HTMLElement, options: FileMenuOptions) {
    this.root = root;
    this.options = options;
  }

  mount(): void {
    const topbar = this.root.querySelector<HTMLElement>(".writer-topbar, .sheet-topbar, .slides-topbar");
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
    panel.innerHTML = `
      <div class="office-file-backdrop" data-file-close></div>
      <div class="office-backstage" role="dialog" aria-label="${this.escape(this.options.title)} File">
        <aside class="office-backstage-nav">
          <div class="office-backstage-brand">
            <span class="office-backstage-icon">${this.iconFor(this.options.module)}</span>
            <div><strong>${this.escape(this.options.title)}</strong><small>File</small></div>
          </div>
          <button class="backstage-nav-active" data-page="home">⌂ Home</button>
          <button data-page="new">＋ New</button>
          <button data-page="open">↥ Open</button>
          <button data-page="recent">◷ Recent</button>
          <div class="backstage-nav-divider"></div>
          <button data-action="save">↓ Save</button>
          <button data-action="save-as">⇩ Save As</button>
          <button data-action="download">⇱ Download</button>
          <button data-action="print">⎙ Print</button>
          <button data-file-close>← Back to editor</button>
        </aside>
        <main class="office-backstage-main">
          <div class="office-backstage-page" data-backstage-page></div>
        </main>
      </div>
      <input data-file-input type="file" accept="${this.escape(this.options.acceptedTypes)}" hidden>
    `;

    this.root.appendChild(panel);
    this.panel = panel;

    const fileInput = panel.querySelector<HTMLInputElement>("[data-file-input]");

    button.addEventListener("click", () => this.setPanelOpen(!panel.hidden));
    panel.querySelectorAll("[data-file-close]").forEach((el) =>
      el.addEventListener("click", () => this.setPanelOpen(false)),
    );

    panel.querySelectorAll<HTMLButtonElement>("[data-page]").forEach((nav) =>
      nav.addEventListener("click", () => this.showPage(nav.dataset.page ?? "home")),
    );

    panel.querySelector('[data-action="save"]')?.addEventListener("click", () => this.run(this.options.onSave));
    panel.querySelector('[data-action="save-as"]')?.addEventListener("click", () => this.run(this.options.onSaveAs));
    panel.querySelector('[data-action="download"]')?.addEventListener("click", () => this.run(this.options.onDownload));
    panel.querySelector('[data-action="print"]')?.addEventListener("click", () => this.run(this.options.onPrint));

    fileInput?.addEventListener("change", () => {
      const file = fileInput.files?.[0];
      fileInput.value = "";
      if (!file) return;
      this.setPanelOpen(false);
      this.options.onOpenFile(file);
    });

    document.addEventListener("keydown", (event) => {
      if (event.altKey && event.key.toLowerCase() === "f") {
        event.preventDefault();
        this.setPanelOpen(true);
      }
      if (event.ctrlKey && event.key.toLowerCase() === "o") {
        event.preventDefault();
        this.setPanelOpen(true);
        this.showPage("open");
      }
      if (event.ctrlKey && event.shiftKey && event.key.toLowerCase() === "s") {
        event.preventDefault();
        this.options.onSaveAs();
        return;
      }
      if (event.ctrlKey && event.key.toLowerCase() === "s") {
        event.preventDefault();
        this.options.onSave();
      }
      if (event.key === "Escape" && !panel.hidden) this.setPanelOpen(false);
    });

    this.showPage("home");
  }

  private run(action: () => void): void {
    this.setPanelOpen(false);
    action();
  }

  private setPanelOpen(open: boolean): void {
    if (!this.panel) return;
    this.panel.hidden = !open;
    this.panel.style.display = open ? "block" : "none";
    if (open) this.showPage("home");
  }

  private showPage(page: string): void {
    if (!this.panel) return;
    this.panel.querySelectorAll<HTMLButtonElement>("[data-page]").forEach((button) =>
      button.classList.toggle("backstage-nav-active", button.dataset.page === page),
    );

    const content = this.panel.querySelector<HTMLElement>("[data-backstage-page]");
    if (!content) return;

    if (page === "new") {
      content.innerHTML = `
        <p class="eyebrow">New</p>
        <h1>Create a new ${this.escape(this.options.title.replace("Oqerun ", ""))}</h1>
        <p class="backstage-lead">Start with a clean file and continue editing in the current workspace.</p>
        <button class="backstage-primary" data-backstage-new>＋ Blank file</button>`;
      content.querySelector("[data-backstage-new]")?.addEventListener("click", () => this.run(this.options.onNew));
      return;
    }

    if (page === "open") {
      content.innerHTML = `
        <p class="eyebrow">Open</p>
        <h1>Open a file</h1>
        <p class="backstage-lead">Browse files from your computer. Oqerun storage providers can be connected through the VFS later.</p>
        <div class="backstage-open-card">
          <button class="backstage-primary" data-backstage-browse>↥ Browse this device</button>
          <span>Supported here: ${this.escape(this.options.acceptedTypes.replaceAll(",", ", "))}</span>
        </div>
        <div class="backstage-open-card">
          <strong>Open Recent</strong>
          <button data-backstage-recent>View recent files →</button>
        </div>`;
      content.querySelector("[data-backstage-browse]")?.addEventListener("click", () => this.panel?.querySelector<HTMLInputElement>("[data-file-input]")?.click());
      content.querySelector("[data-backstage-recent]")?.addEventListener("click", () => this.showPage("recent"));
      return;
    }

    if (page === "recent") {
      const recent = getRecent().filter((item) => item.module === this.options.module);
      content.innerHTML = `
        <div class="backstage-heading-row">
          <div><p class="eyebrow">Recent</p><h1>Recent ${this.escape(this.options.title.replace("Oqerun ", ""))} files</h1></div>
          <button data-clear-recent>Clear recent</button>
        </div>
        <div class="backstage-search-row"><input data-recent-search placeholder="Search recent files..." /></div>
        <div class="backstage-recent-list" data-backstage-recent-list></div>`;
      const list = content.querySelector<HTMLElement>("[data-backstage-recent-list]");
      const search = content.querySelector<HTMLInputElement>("[data-recent-search]");
      const render = () => {
        const term = (search?.value ?? "").toLowerCase();
        const filtered = recent.filter((item) => item.name.toLowerCase().includes(term));
        if (!filtered.length) {
          if (list) list.innerHTML = `<div class="backstage-empty"><strong>No matching recent files</strong><span>Open or save a file and it will appear here.</span></div>`;
          return;
        }
        if (list) list.innerHTML = filtered.map((item) => `
          <div class="backstage-recent-item">
            <button class="backstage-recent-main" data-open-recent="${this.escapeAttribute(item.path)}">
              <span class="backstage-recent-icon">${this.iconFor(item.module)}</span>
              <span><strong>${this.escape(item.name)}</strong><small>${new Date(item.modifiedAt).toLocaleString()}</small></span>
            </button>
            <button class="backstage-icon-action" title="${item.pinned ? "Unpin" : "Pin"}" data-pin="${this.escapeAttribute(item.path)}">${item.pinned ? "★" : "☆"}</button>
            <button class="backstage-icon-action" title="Remove" data-remove="${this.escapeAttribute(item.path)}">×</button>
          </div>`).join("");
        list?.querySelectorAll<HTMLElement>("[data-open-recent]").forEach((el) => el.addEventListener("click", () => {
          const path = el.dataset.openRecent;
          if (path) { this.setPanelOpen(false); this.options.onOpenRecent(path); }
        }));
        list?.querySelectorAll<HTMLElement>("[data-pin]").forEach((el) => el.addEventListener("click", () => {
          togglePinned(el.dataset.pin ?? "");
          this.showPage("recent");
        }));
        list?.querySelectorAll<HTMLElement>("[data-remove]").forEach((el) => el.addEventListener("click", () => {
          removeRecent(el.dataset.remove ?? "");
          this.showPage("recent");
        }));
      };
      search?.addEventListener("input", render);
      content.querySelector("[data-clear-recent]")?.addEventListener("click", () => {
        clearRecent();
        this.showPage("recent");
      });
      render();
      return;
    }

    content.innerHTML = `
      <p class="eyebrow">Home</p>
      <h1>What do you want to do?</h1>
      <p class="backstage-lead">Manage this file without leaving ${this.escape(this.options.title)}.</p>
      <div class="backstage-command-grid">
        <button data-home-command="new"><b>＋</b><span><strong>New</strong><small>Create a blank file</small></span></button>
        <button data-home-command="open"><b>↥</b><span><strong>Open</strong><small>Browse files on this device</small></span></button>
        <button data-home-command="recent"><b>◷</b><span><strong>Recent</strong><small>Return to your recent files</small></span></button>
        <button data-home-command="save"><b>↓</b><span><strong>Save</strong><small>Save the current file</small></span></button>
        <button data-home-command="save-as"><b>⇩</b><span><strong>Save As</strong><small>Choose a new name</small></span></button>
        <button data-home-command="download"><b>⇱</b><span><strong>Download</strong><small>Export a local copy</small></span></button>
      </div>`;
    content.querySelectorAll<HTMLButtonElement>("[data-home-command]").forEach((el) =>
      el.addEventListener("click", () => {
        const command = el.dataset.homeCommand;
        if (command === "new") return this.run(this.options.onNew);
        if (command === "open") return this.showPage("open");
        if (command === "recent") return this.showPage("recent");
        if (command === "save") return this.run(this.options.onSave);
        if (command === "save-as") return this.run(this.options.onSaveAs);
        if (command === "download") return this.run(this.options.onDownload);
      }),
    );
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
