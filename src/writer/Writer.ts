import type { OfficeDocument, WriterBlock, WriterContent, WriterSettings } from "../core/DocumentModel";
import { DocumentStore } from "../core/DocumentStore";
import { exportWriterDocx, importWriterDocx, downloadText } from "../core/FormatEngine";
import { FileMenu } from "../ui/FileMenu";

const DEFAULT_SETTINGS: WriterSettings = {
  fontFamily: "Georgia",
  fontSize: 17,
  textColor: "#202634",
  highlightColor: "#fff59d",
  lineSpacing: 1.7,
  indent: 0,
  pageSize: "A4",
  margin: 72,
  header: "",
  footer: "",
};

const EMPTY_HTML = "<h1>Untitled document</h1><p>Start writing with Oqerun Writer.</p><p>Use the toolbar to format your document.</p>";

export class Writer {
  private readonly root: HTMLElement;
  private readonly store = new DocumentStore();
  private documentId: string = crypto.randomUUID();
  private createdAt = new Date().toISOString();
  private fileName = "Untitled document";
  private settings: WriterSettings = { ...DEFAULT_SETTINGS };

  constructor(root: HTMLElement) {
    this.root = root;
  }

  render(): void {
    this.root.innerHTML = `
      <main class="writer-app">
        <header class="writer-titlebar">
          <div class="writer-titlebar-left">
            <button class="writer-back" data-action="back" type="button" aria-label="Back to Oqerun Office">←</button>
            <div class="writer-brand">
              <span class="writer-logo">W</span>
              <div class="writer-brand-text">
                <strong>Oqerun Writer</strong>
                <span data-file-name>${this.escape(this.fileName)}</span>
              </div>
            </div>
          </div>
          <div class="writer-quick-actions">
            <button type="button" data-action="undo" aria-label="Undo">↶</button>
            <button type="button" data-action="redo" aria-label="Redo">↷</button>
            <span class="writer-save-state" data-status>Ready</span>
            <button type="button" class="writer-save-button" data-action="save">Save</button>
            <button type="button" class="writer-account" aria-label="Account">V</button>
          </div>
        </header>

        <div class="writer-menu-row">
          <button type="button" class="writer-file-button" data-file-menu>File</button>
          <nav class="writer-tabs" aria-label="Writer tabs">
            <button type="button" class="writer-tab is-active" data-ribbon-tab="home">Home</button>
            <button type="button" class="writer-tab" data-ribbon-tab="insert">Insert</button>
            <button type="button" class="writer-tab" data-ribbon-tab="layout">Layout</button>
            <button type="button" class="writer-tab" data-ribbon-tab="review">Review</button>
            <button type="button" class="writer-tab" data-ribbon-tab="view">View</button>
          </nav>
          <button type="button" class="writer-search" aria-label="Search commands">⌕ <span>Search</span> <kbd>Ctrl K</kbd></button>
        </div>

        <section class="writer-ribbon" aria-label="Writer ribbon">
          <div class="writer-ribbon-scroll">
            <div class="writer-group" data-ribbon-group="Clipboard" data-ribbon-tab="home">
              <div class="writer-group-tools writer-clipboard-tools">
                <button type="button" data-command="undo" title="Undo">↶</button>
                <button type="button" data-command="redo" title="Redo">↷</button>
              </div>
            </div>

            <div class="writer-group writer-font-group" data-ribbon-group="Font" data-ribbon-tab="home">
              <div class="writer-group-tools writer-font-tools">
                <select data-command="formatBlock" aria-label="Style">
                  <option value="p">Normal</option><option value="h1">Title</option><option value="h2">Heading 1</option><option value="h3">Heading 2</option>
                </select>
                <select data-command="fontName" aria-label="Font">
                  <option value="Arial">Arial</option><option value="Calibri">Calibri</option><option value="Georgia">Georgia</option><option value="Times New Roman">Times New Roman</option><option value="Verdana">Verdana</option>
                </select>
                <select data-action="font-size" aria-label="Font size">
                  <option value="12">12</option><option value="14">14</option><option value="16">16</option><option value="18" selected>18</option><option value="24">24</option><option value="32">32</option><option value="48">48</option>
                </select>
              </div>
              <div class="writer-group-tools writer-format-tools">
                <button type="button" data-command="bold" title="Bold"><b>B</b></button>
                <button type="button" data-command="italic" title="Italic"><i>I</i></button>
                <button type="button" data-command="underline" title="Underline"><u>U</u></button>
                <button type="button" data-command="strikeThrough" title="Strikethrough"><s>S</s></button>
                <label class="writer-color" title="Text color">A<input data-color="text" type="color" value="#202634"></label>
                <label class="writer-color" title="Text highlight">H<input data-color="highlight" type="color" value="#fff59d"></label>
              </div>
            </div>

            <div class="writer-group" data-ribbon-group="Paragraph" data-ribbon-tab="home">
              <div class="writer-group-tools">
                <button type="button" data-command="insertUnorderedList" title="Bulleted list">•</button>
                <button type="button" data-command="insertOrderedList" title="Numbered list">1.</button>
                <button type="button" data-command="outdent" title="Decrease indent">⇤</button>
                <button type="button" data-command="indent" title="Increase indent">⇥</button>
                <button type="button" data-command="justifyLeft" title="Align left">≡</button>
                <button type="button" data-command="justifyCenter" title="Center">≣</button>
                <button type="button" data-command="justifyRight" title="Align right">≡</button>
                <select data-action="line-spacing" aria-label="Line spacing">
                  <option value="1">1.0</option><option value="1.15">1.15</option><option value="1.5">1.5</option><option value="1.7" selected>1.7</option><option value="2">2.0</option>
                </select>
              </div>
            </div>

            <div class="writer-group" data-ribbon-group="Editing" data-ribbon-tab="home">
              <div class="writer-group-tools">
                <button type="button" data-action="find">Find</button>
                <button type="button" data-action="replace">Replace</button>
              </div>
            </div>

            <div class="writer-group writer-insert-group" data-ribbon-group="Insert" data-ribbon-tab="insert">
              <div class="writer-group-tools">
                <button type="button" data-action="link">Link</button>
                <button type="button" data-action="table">Table</button>
                <button type="button" data-action="image">Pictures</button>
                <button type="button" data-action="equation">Equation</button>
                <button type="button" data-action="layout">Page</button>
              </div>
            </div>

            <div class="writer-group" data-ribbon-group="Review" data-ribbon-tab="review">
              <div class="writer-group-tools">
                <button type="button" data-action="find">Find</button>
                <button type="button" data-action="replace">Replace</button>
              </div>
            </div>

            <div class="writer-group writer-document-group" data-ribbon-group="Document" data-ribbon-tab="home">
              <div class="writer-group-tools">
                <button type="button" data-action="new">New</button>
                <button type="button" class="writer-ribbon-save" data-action="save">Save</button>
              </div>
            </div>
          </div>
        </section>

        <section class="writer-layout-strip" data-settings-strip hidden>
          <div class="writer-layout-controls">
            <label>Paper <select data-setting="pageSize"><option value="A4">A4</option><option value="Letter">Letter</option></select></label>
            <label>Margins <input data-setting="margin" type="number" min="36" max="180" step="6" value="72"></label>
            <label>Header <input data-setting="header" placeholder="Header text"></label>
            <label>Footer <input data-setting="footer" placeholder="Footer text"></label>
            <button data-action="close-layout" type="button">Done</button>
          </div>
        </section>

        <section class="writer-canvas" aria-label="Document workspace">
          <div class="writer-ruler" aria-hidden="true">
            <span>0</span><span>1</span><span>2</span><span>3</span><span>4</span><span>5</span><span>6</span><span>7</span><span>8</span>
          </div>
          <article class="writer-page" data-page>
            <div class="writer-page-header" data-page-header></div>
            <div class="writer-editor" contenteditable="true" spellcheck="true" role="textbox" aria-label="Document editor">${EMPTY_HTML}</div>
            <div class="writer-page-footer" data-page-footer></div>
          </article>
        </section>

        <footer class="writer-statusbar" aria-label="Document status">
          <span data-page-status>Page 1 of 1</span>
          <span data-word-count>Words: 0</span>
          <span class="writer-status-divider"></span>
          <span data-save-status>Ready</span>
          <span class="writer-status-spacer"></span>
          <span>English</span>
          <span class="writer-zoom">− &nbsp;100%&nbsp; +</span>
        </footer>

        <input data-image-input type="file" accept="image/*" hidden>
      </main>`;
    this.applySettings();
    this.bind();
    this.setupRibbon();
    this.mountFileMenu();
  }

  async openPath(path: string): Promise<void> {
    const document = await this.store.load<WriterContent>(path);
    this.documentId = document.metadata.id;
    this.createdAt = document.metadata.createdAt;
    this.fileName = document.metadata.name;
    this.settings = { ...DEFAULT_SETTINGS, ...(document.content.settings ?? {}) };

    const editor = this.root.querySelector<HTMLElement>(".writer-editor");
    if (!editor) {
      this.render();
      return this.openPath(path);
    }

    editor.innerHTML = document.content.html || EMPTY_HTML;
    this.applySettings();
    this.updateFileName();
    this.setStatus("Opened");
  }

  private setupRibbon(): void {
    const ribbon = this.root.querySelector<HTMLElement>(".writer-ribbon");
    const tabs = Array.from(this.root.querySelectorAll<HTMLButtonElement>("[data-ribbon-tab]"));
    const groups = Array.from(this.root.querySelectorAll<HTMLElement>(".writer-group"));
    const settingsStrip = this.root.querySelector<HTMLElement>("[data-settings-strip]");
    if (!ribbon || !tabs.length) return;

    const showTab = (tab: string): void => {
      tabs.forEach((button) => {
        const active = button.dataset.ribbonTab === tab;
        button.classList.toggle("is-active", active);
        button.setAttribute("aria-selected", String(active));
      });

      groups.forEach((group) => {
        group.hidden = group.dataset.ribbonTab !== tab;
      });

      if (settingsStrip) settingsStrip.hidden = tab !== "layout";

      ribbon.querySelectorAll<HTMLElement>(".writer-view-panel").forEach((node) => node.remove());

      if (tab === "layout") {
        const panel = document.createElement("div");
        panel.className = "writer-view-panel";
        panel.innerHTML = `<strong>Page layout</strong><span>Paper size, margins, headers and footers are shown below.</span>`;
        ribbon.querySelector(".writer-ribbon-scroll")?.appendChild(panel);
      }

      if (tab === "view") {
        const panel = document.createElement("div");
        panel.className = "writer-view-panel";
        panel.innerHTML = `<button type="button" data-view-action="focus">Focus</button><button type="button" data-view-action="width">Page width</button><button type="button" data-view-action="top">Print layout</button>`;
        ribbon.querySelector(".writer-ribbon-scroll")?.appendChild(panel);
        panel.querySelector('[data-view-action="focus"]')?.addEventListener("click", () => {
          this.root.querySelector(".writer-app")?.classList.toggle("writer-focus-mode");
        });
        panel.querySelector('[data-view-action="width"]')?.addEventListener("click", () => {
          this.root.querySelector(".writer-page")?.classList.toggle("writer-page-wide");
        });
        panel.querySelector('[data-view-action="top"]')?.addEventListener("click", () => {
          this.root.querySelector(".writer-canvas")?.scrollTo({ top: 0, behavior: "smooth" });
        });
      }
    };

    tabs.forEach((tab) => tab.addEventListener("click", () => {
      showTab(tab.dataset.ribbonTab ?? "home");
    }));

    showTab("home");
  }


  private mountFileMenu(): void {
    new FileMenu(this.root, {
      module: "writer",
      title: "Oqerun Writer",
      acceptedTypes: ".oqdoc,.docx,.txt,.html",
      onHome: () => this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true })),
      onNew: () => this.newDocument(),
      onSaveAs: () => { const editor = this.editor(); if (editor) void this.saveAs(editor); },
      onDownload: () => { const editor = this.editor(); if (editor) void this.download(editor); },
      onPrint: () => window.print(),
      onSave: () => { const editor = this.editor(); if (editor) void this.save(editor); },
      onOpenRecent: (path) => void this.openPath(path),
      onOpenFile: (file) => void this.importFile(file),
    }).mount();
  }

  private editor(): HTMLElement | null {
    return this.root.querySelector<HTMLElement>(".writer-editor");
  }

  private newDocument(): void {
    const editor = this.editor();
    if (!editor) return;
    this.documentId = crypto.randomUUID();
    this.createdAt = new Date().toISOString();
    this.fileName = "Untitled document";
    this.settings = { ...DEFAULT_SETTINGS };
    editor.innerHTML = EMPTY_HTML;
    this.applySettings();
    this.updateFileName();
    this.setStatus("New document");
    editor.focus();
  }

  public async importFile(file: File): Promise<void> {
    const lower = file.name.toLowerCase();

    if (lower.endsWith(".docx")) {
      try {
        const imported = await importWriterDocx(file);
        this.documentId = crypto.randomUUID();
        this.createdAt = new Date().toISOString();
        this.fileName = file.name.replace(/\.docx$/i, "") || "Untitled document";
        const editor = this.editor();
        if (!editor) this.render();
        const target = this.editor();
        if (!target) return;
        target.innerHTML = imported.html;
        this.applySettings();
        this.setStatus("Imported DOCX");
        await this.save(target);
      } catch {
        window.alert("The selected DOCX file could not be imported.");
      }
      return;
    }

    if (lower.endsWith(".oqdoc")) {
      try {
        const parsed = JSON.parse(await file.text()) as OfficeDocument<WriterContent>;
        if (!parsed?.content?.html || !parsed.metadata) throw new Error("Invalid Oqerun document");
        const path = await this.store.save({
          ...parsed,
          metadata: { ...parsed.metadata, path: `/home/user/Documents/${this.slug(parsed.metadata.name)}.oqdoc` },
        });
        this.render();
        await this.openPath(path);
      } catch {
        window.alert("The selected Oqerun document could not be opened.");
      }
      return;
    }

    if (lower.endsWith(".txt") || lower.endsWith(".html")) {
      const text = await file.text();
      const isHtml = lower.endsWith(".html");
      const name = file.name.replace(/\.(txt|html)$/i, "") || "Untitled document";
      this.documentId = crypto.randomUUID();
      this.createdAt = new Date().toISOString();
      this.fileName = name;
      const editor = this.editor();
      if (!editor) this.render();
      const target = this.editor();
      if (!target) return;
      target.innerHTML = isHtml ? text : "<p>" + this.escape(text).replace(/\r?\n/g, "</p><p>") + "</p>";
      this.applySettings();
      this.setStatus("Imported text");
      await this.save(target);
      return;
    }

    window.alert("Oqerun Writer currently supports .oqdoc, .docx, .txt and .html imports.");
  }

  private bind(): void {
    const editor = this.editor();
    if (!editor) return;

    this.root.querySelectorAll<HTMLElement>("[data-command]").forEach((control) => {
      control.addEventListener("click", () => {
        const command = control.dataset.command;
        if (!command || command === "formatBlock") return;
        document.execCommand(command, false);
        editor.focus();
        this.setStatus("Unsaved changes");
      });
    });

    this.root.querySelector<HTMLSelectElement>('select[data-command="formatBlock"]')?.addEventListener("change", (event) => {
      document.execCommand("formatBlock", false, (event.currentTarget as HTMLSelectElement).value);
      editor.focus();
      this.setStatus("Unsaved changes");
    });

    this.root.querySelector<HTMLSelectElement>('select[data-command="fontName"]')?.addEventListener("change", (event) => {
      const value = (event.currentTarget as HTMLSelectElement).value;
      document.execCommand("fontName", false, value);
      this.settings.fontFamily = value;
      editor.focus();
      this.setStatus("Unsaved changes");
    });

    this.root.querySelector('[data-action="font-size"]')?.addEventListener("change", (event) => {
      const value = Number((event.currentTarget as HTMLSelectElement).value);
      const commandSize = value <= 12 ? "2" : value <= 14 ? "3" : value <= 16 ? "4" : value <= 18 ? "5" : value <= 24 ? "6" : "7";
      document.execCommand("fontSize", false, commandSize);
      this.settings.fontSize = value;
      editor.focus();
      this.setStatus("Unsaved changes");
    });

    this.root.querySelectorAll<HTMLInputElement>("[data-color]").forEach((input) =>
      input.addEventListener("input", () => {
        const kind = input.dataset.color;
        if (kind === "text") {
          this.settings.textColor = input.value;
          document.execCommand("foreColor", false, input.value);
        } else {
          this.settings.highlightColor = input.value;
          document.execCommand("hiliteColor", false, input.value);
        }
        editor.focus();
        this.setStatus("Unsaved changes");
      }),
    );

    this.root.querySelector<HTMLSelectElement>('[data-action="line-spacing"]')?.addEventListener("change", (event) => {
      this.settings.lineSpacing = Number((event.currentTarget as HTMLSelectElement).value);
      editor.style.lineHeight = String(this.settings.lineSpacing);
      this.setStatus("Unsaved changes");
    });

    editor.addEventListener("input", () => {
      this.setStatus("Unsaved changes");
      this.updateDocumentStats(editor);
    });
    this.updateDocumentStats(editor);

    this.root.querySelector('[data-action="save"]')?.addEventListener("click", () => void this.save(editor));
    this.root.querySelector('[data-action="new"]')?.addEventListener("click", () => this.newDocument());

    this.root.querySelector('[data-action="link"]')?.addEventListener("click", () => {
      const url = window.prompt("Link URL", "https://");
      if (url) {
        document.execCommand("createLink", false, url);
        editor.focus();
        this.setStatus("Unsaved changes");
      }
    });

    this.root.querySelector('[data-action="table"]')?.addEventListener("click", () => this.insertTable(editor));
    this.root.querySelector('[data-action="equation"]')?.addEventListener("click", () => this.insertEquation(editor));

    const imageInput = this.root.querySelector<HTMLInputElement>("[data-image-input]");
    this.root.querySelector('[data-action="image"]')?.addEventListener("click", () => imageInput?.click());
    imageInput?.addEventListener("change", () => {
      const file = imageInput.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        document.execCommand("insertHTML", false, `<img src="${String(reader.result)}" alt="${this.escape(file.name)}" />`);
        editor.focus();
        this.setStatus("Unsaved changes");
      };
      reader.readAsDataURL(file);
      imageInput.value = "";
    });

    this.root.querySelector('[data-action="find"]')?.addEventListener("click", () => this.find(false));
    this.root.querySelector('[data-action="replace"]')?.addEventListener("click", () => this.find(true));

    this.root.querySelector('[data-action="layout"]')?.addEventListener("click", () => {
      const strip = this.root.querySelector<HTMLElement>("[data-settings-strip]");
      if (strip) strip.hidden = !strip.hidden;
    });
    this.root.querySelector('[data-action="close-layout"]')?.addEventListener("click", () => {
      const strip = this.root.querySelector<HTMLElement>("[data-settings-strip]");
      if (strip) strip.hidden = true;
    });

    this.root.querySelectorAll<HTMLInputElement | HTMLSelectElement>("[data-setting]").forEach((control) =>
      control.addEventListener("input", () => this.updateSetting(control)),
    );
    this.root.querySelectorAll<HTMLInputElement | HTMLSelectElement>("[data-setting]").forEach((control) =>
      control.addEventListener("change", () => this.updateSetting(control)),
    );

    editor.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        void this.save(editor);
      }
    });

    const draft = localStorage.getItem("oqerun-writer-draft");
    if (draft && this.fileName === "Untitled document") editor.innerHTML = draft;

    this.root.querySelector('[data-action="back"]')?.addEventListener("click", () =>
      this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true })),
    );
  }

  private updateSetting(control: HTMLInputElement | HTMLSelectElement): void {
    const key = control.dataset.setting as keyof WriterSettings;
    if (!key) return;
    if (key === "margin") this.settings.margin = Number((control as HTMLInputElement).value) || 72;
    else if (key === "pageSize") this.settings.pageSize = control.value as "A4" | "Letter";
    else if (key === "header" || key === "footer") this.settings[key] = control.value;
    this.applySettings();
    this.setStatus("Unsaved changes");
  }

  private applySettings(): void {
    const editor = this.editor();
    const page = this.root.querySelector<HTMLElement>("[data-page]");
    if (editor) {
      editor.style.fontFamily = this.settings.fontFamily;
      editor.style.fontSize = `${this.settings.fontSize}px`;
      editor.style.lineHeight = String(this.settings.lineSpacing);
      editor.style.color = this.settings.textColor;
    }
    if (page) {
      const width = this.settings.pageSize === "Letter" ? "816px" : "820px";
      page.style.width = `min(${width},100%)`;
      page.style.padding = `${this.settings.margin}px`;
    }
    const pageSize = this.root.querySelector<HTMLSelectElement>('[data-setting="pageSize"]');
    const margin = this.root.querySelector<HTMLInputElement>('[data-setting="margin"]');
    const header = this.root.querySelector<HTMLInputElement>('[data-setting="header"]');
    const footer = this.root.querySelector<HTMLInputElement>('[data-setting="footer"]');
    if (pageSize) pageSize.value = this.settings.pageSize;
    if (margin) margin.value = String(this.settings.margin);
    if (header) header.value = this.settings.header;
    if (footer) footer.value = this.settings.footer;
  }

  private find(replace: boolean): void {
    const editor = this.editor();
    if (!editor) return;
    const term = window.prompt(replace ? "Find text" : "Find text");
    if (!term) return;

    if (!replace) {
      const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
      let node: Node | null;
      while ((node = walker.nextNode())) {
        const index = (node.nodeValue ?? "").toLowerCase().indexOf(term.toLowerCase());
        if (index >= 0) {
          const range = document.createRange();
          range.setStart(node, index);
          range.setEnd(node, index + term.length);
          const selection = window.getSelection();
          selection?.removeAllRanges();
          selection?.addRange(range);
          editor.focus();
          return;
        }
      }
      window.alert("Text not found.");
      return;
    }

    const replacement = window.prompt("Replace with", term);
    if (replacement === null) return;
    const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    const nodes: Text[] = [];
    let node: Node | null;
    while ((node = walker.nextNode())) nodes.push(node as Text);
    let count = 0;
    for (const text of nodes) {
      const value = text.nodeValue ?? "";
      const regex = new RegExp(term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
      const next = value.replace(regex, () => { count += 1; return replacement; });
      if (next !== value) text.nodeValue = next;
    }
    this.setStatus(count ? `Replaced ${count} occurrence(s)` : "Text not found");
  }

  private async saveAs(editor: HTMLElement): Promise<void> {
    const value = window.prompt("Save As — document name", this.fileName);
    if (!value?.trim()) return;
    this.fileName = value.trim().replace(/\.oqdoc$/i, "");
    await this.save(editor);
  }

  private async download(editor: HTMLElement): Promise<void> {
    const choice = (window.prompt("Download format:\n1. Oqerun document (.oqdoc)\n2. Word document (.docx)\n3. HTML (.html)\n4. Text (.txt)", "2") ?? "2").trim();
    if (choice === "2") {
      await exportWriterDocx(this.currentContent(editor), this.fileName);
      this.setStatus("DOCX downloaded");
      return;
    }
    if (choice === "3") {
      downloadText(editor.innerHTML, this.slug(this.fileName) + ".html", "text/html");
      this.setStatus("HTML downloaded");
      return;
    }
    if (choice === "4") {
      downloadText(editor.innerText, this.slug(this.fileName) + ".txt");
      this.setStatus("Text downloaded");
      return;
    }
    const payload = JSON.stringify(this.currentDocument(editor), null, 2);
    downloadText(payload, this.slug(this.fileName) + ".oqdoc", "application/json");
    this.setStatus("Oqerun document downloaded");
  }

  private currentContent(editor: HTMLElement): WriterContent {
    return {
      html: editor.innerHTML,
      plainText: editor.innerText,
      blocks: this.blocksFromEditor(editor),
      settings: { ...this.settings },
    };
  }

  private currentDocument(editor: HTMLElement): OfficeDocument<WriterContent> {
    return {
      metadata: {
        id: this.documentId,
        name: this.fileName,
        path: `/home/user/Documents/${this.slug(this.fileName)}.oqdoc`,
        mimeType: "application/x-oqerun-document",
        createdAt: this.createdAt,
        modifiedAt: new Date().toISOString(),
      },
      content: this.currentContent(editor),
    };
  }

  private async save(editor: HTMLElement): Promise<void> {
    await this.store.save(this.currentDocument(editor));
    localStorage.setItem("oqerun-writer-draft", editor.innerHTML);
    this.updateFileName();
    this.setStatus("Saved locally");
  }

  private blocksFromEditor(editor: HTMLElement): WriterBlock[] {
    return Array.from(editor.children).map((node) => {
      const tag = node.tagName.toLowerCase();
      if (/^h[1-6]$/.test(tag)) return { type: "heading", data: { level: Number(tag.slice(1)), html: node.innerHTML, text: node.textContent ?? "" } };
      if (tag === "ul" || tag === "ol") return { type: "list", data: { ordered: tag === "ol", html: node.outerHTML } };
      if (tag === "table") return { type: "table", data: { html: node.outerHTML } };
      if (tag === "img") return { type: "image", data: { src: (node as HTMLImageElement).src, alt: (node as HTMLImageElement).alt } };
      if (node.classList.contains("writer-equation")) return { type: "equation", data: { value: node.textContent ?? "" } };
      return { type: "paragraph", data: { html: node.innerHTML, text: node.textContent ?? "" } };
    });
  }

  private insertTable(editor: HTMLElement): void {
    const rows = Math.max(1, Math.min(20, Number(window.prompt("Rows", "3")) || 3));
    const columns = Math.max(1, Math.min(12, Number(window.prompt("Columns", "3")) || 3));
    let html = "<table><tbody>";
    for (let row = 0; row < rows; row += 1) {
      html += "<tr>";
      for (let column = 0; column < columns; column += 1) html += `<td>Cell ${row + 1}.${column + 1}</td>`;
      html += "</tr>";
    }
    html += "</tbody></table><p></p>";
    document.execCommand("insertHTML", false, html);
    editor.focus();
    this.setStatus("Unsaved changes");
  }

  private insertEquation(editor: HTMLElement): void {
    const choice = window.prompt(
      "Equation template:\n1. Quadratic formula\n2. Pythagorean theorem\n3. Summation\n4. Integral\n5. Custom",
      "1",
    );
    const templates: Record<string, string> = {
      "1": "x = (−b ± √(b² − 4ac)) / 2a",
      "2": "a² + b² = c²",
      "3": "Σᵢ₌₁ⁿ i = n(n + 1) / 2",
      "4": "∫₀¹ x² dx = 1/3",
    };
    const value = choice === "5" ? window.prompt("Enter your equation", "E = mc²") ?? "" : templates[choice ?? "1"] ?? templates["1"];
    if (!value) return;
    document.execCommand("insertHTML", false, `<div class="writer-equation" contenteditable="false">${this.escape(value)}</div><p></p>`);
    editor.focus();
    this.setStatus("Equation inserted");
  }

  private updateFileName(): void {
    const element = this.root.querySelector<HTMLElement>("[data-file-name]");
    if (element) element.textContent = this.fileName;
  }

  private setStatus(text: string): void {
    const status = this.root.querySelector<HTMLElement>("[data-status]");
    if (status) status.textContent = text;
    const saveStatus = this.root.querySelector<HTMLElement>("[data-save-status]");
    if (saveStatus) saveStatus.textContent = text;
  }

  private updateDocumentStats(editor: HTMLElement): void {
    const text = editor.innerText.trim();
    const words = text ? text.split(/\s+/).filter(Boolean).length : 0;
    const wordCount = this.root.querySelector<HTMLElement>("[data-word-count]");
    if (wordCount) wordCount.textContent = `Words: ${words}`;
  }

  private slug(value: string): string {
    return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "untitled";
  }

  private escape(value: string): string {
    return value.replace(/[&<>"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[char] ?? char));
  }
}
