import type { OfficeDocument, WriterBlock, WriterContent } from "../core/DocumentModel";
import { DocumentStore } from "../core/DocumentStore";

const EMPTY_HTML = "<h1>Untitled document</h1><p>Start writing with Oqerun Writer.</p><p>Use the toolbar to format your document.</p>";

export class Writer {
  private readonly root: HTMLElement;
  private readonly store = new DocumentStore();
  private documentId = crypto.randomUUID();
  private createdAt = new Date().toISOString();
  private fileName = "Untitled document";

  constructor(root: HTMLElement) {
    this.root = root;
  }

  render(): void {
    this.root.innerHTML = `
      <main class="writer-shell">
        <header class="writer-topbar">
          <button class="writer-back" data-action="back" type="button">← Oqerun Office</button>
          <div class="writer-title">
            <span class="writer-app-icon">W</span>
            <div><strong>Oqerun Writer</strong><span data-file-name>${this.escape(this.fileName)}</span></div>
          </div>
          <span class="writer-status" data-status>Ready</span>
        </header>
        <nav class="writer-toolbar" aria-label="Writer toolbar">
          <div class="toolbar-group">
            <button data-command="undo" type="button">↶</button>
            <button data-command="redo" type="button">↷</button>
          </div>
          <div class="toolbar-group">
            <select data-command="formatBlock" aria-label="Text style">
              <option value="p">Normal text</option><option value="h1">Heading 1</option><option value="h2">Heading 2</option><option value="h3">Heading 3</option>
            </select>
          </div>
          <div class="toolbar-group">
            <button data-command="bold" type="button"><b>B</b></button>
            <button data-command="italic" type="button"><i>I</i></button>
            <button data-command="underline" type="button"><u>U</u></button>
            <button data-command="strikeThrough" type="button"><s>S</s></button>
          </div>
          <div class="toolbar-group">
            <button data-command="insertUnorderedList" type="button">• List</button>
            <button data-command="insertOrderedList" type="button">1. List</button>
          </div>
          <div class="toolbar-group">
            <button data-command="justifyLeft" type="button">Left</button>
            <button data-command="justifyCenter" type="button">Center</button>
            <button data-command="justifyRight" type="button">Right</button>
          </div>
          <div class="toolbar-group">
            <button data-action="link" type="button">Link</button>
            <button data-action="table" type="button">Table</button>
            <button data-action="equation" type="button">∑ Equation</button>
            <button data-action="image" type="button">Image</button>
          </div>
          <div class="toolbar-group toolbar-actions">
            <button data-action="new" type="button">New</button>
            <button data-action="save" class="primary-tool" type="button">Save</button>
            <input data-image-input type="file" accept="image/*" hidden>
          </div>
        </nav>
        <section class="writer-workspace">
          <div class="writer-ruler"></div>
          <article class="writer-page">
            <div class="writer-editor" contenteditable="true" spellcheck="true" role="textbox" aria-label="Document editor">${EMPTY_HTML}</div>
          </article>
        </section>
      </main>`;
    this.bind();
  }

  async openPath(path: string): Promise<void> {
    const document = await this.store.load<WriterContent>(path);
    this.documentId = document.metadata.id;
    this.createdAt = document.metadata.createdAt;
    this.fileName = document.metadata.name;

    const editor = this.root.querySelector<HTMLElement>(".writer-editor");
    if (!editor) {
      this.render();
      return this.openPath(path);
    }

    editor.innerHTML = document.content.html || EMPTY_HTML;
    this.updateFileName();
    this.setStatus("Opened");
  }

  private bind(): void {
    const editor = this.root.querySelector<HTMLElement>(".writer-editor");
    if (!editor) return;

    this.root.querySelectorAll<HTMLElement>("[data-command]").forEach((control) =>
      control.addEventListener("click", () => {
        const command = control.dataset.command;
        if (command && command !== "formatBlock") document.execCommand(command, false);
        editor.focus();
        this.setStatus("Unsaved changes");
      }),
    );

    this.root.querySelector<HTMLSelectElement>('select[data-command="formatBlock"]')?.addEventListener("change", (event) => {
      document.execCommand("formatBlock", false, (event.currentTarget as HTMLSelectElement).value);
      editor.focus();
      this.setStatus("Unsaved changes");
    });

    editor.addEventListener("input", () => this.setStatus("Unsaved changes"));

    this.root.querySelector('[data-action="save"]')?.addEventListener("click", () => void this.save(editor));

    this.root.querySelector('[data-action="new"]')?.addEventListener("click", () => {
      this.documentId = crypto.randomUUID();
      this.createdAt = new Date().toISOString();
      this.fileName = "Untitled document";
      editor.innerHTML = EMPTY_HTML;
      this.updateFileName();
      this.setStatus("New document");
      editor.focus();
    });

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
    });

    editor.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        void this.save(editor);
      }
    });

    const draft = localStorage.getItem("oqerun-writer-draft");
    if (draft) editor.innerHTML = draft;

    this.root.querySelector('[data-action="back"]')?.addEventListener("click", () =>
      this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true })),
    );
  }

  private async save(editor: HTMLElement): Promise<void> {
    const now = new Date().toISOString();
    const document: OfficeDocument<WriterContent> = {
      metadata: {
        id: this.documentId,
        name: this.fileName,
        path: `/home/user/Documents/${this.slug(this.fileName)}.oqdoc`,
        mimeType: "application/x-oqerun-document",
        createdAt: this.createdAt,
        modifiedAt: now,
      },
      content: {
        html: editor.innerHTML,
        plainText: editor.innerText,
        blocks: this.blocksFromEditor(editor),
      },
    };

    await this.store.save(document);
    localStorage.setItem("oqerun-writer-draft", editor.innerHTML);
    this.setStatus("Saved locally");
  }

  private blocksFromEditor(editor: HTMLElement): WriterBlock[] {
    return Array.from(editor.children).map((node) => {
      const tag = node.tagName.toLowerCase();
      if (/^h[1-6]$/.test(tag)) {
        return { type: "heading", data: { level: Number(tag.slice(1)), html: node.innerHTML, text: node.textContent ?? "" } };
      }
      if (tag === "ul" || tag === "ol") return { type: "list", data: { ordered: tag === "ol", html: node.outerHTML } };
      if (tag === "table") return { type: "table", data: { html: node.outerHTML } };
      if (tag === "img") return { type: "image", data: { src: (node as HTMLImageElement).src, alt: (node as HTMLImageElement).alt } };
      if (node.classList.contains("writer-equation")) return { type: "equation", data: { value: node.textContent ?? "" } };
      return { type: "paragraph", data: { html: node.innerHTML, text: node.textContent ?? "" } };
    });
  }

  private insertTable(editor: HTMLElement): void {
    const rows = Math.max(1, Math.min(10, Number(window.prompt("Rows", "3")) || 3));
    const columns = Math.max(1, Math.min(8, Number(window.prompt("Columns", "3")) || 3));
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

    const value = choice === "5"
      ? window.prompt("Enter your equation", "E = mc²") ?? ""
      : templates[choice ?? "1"] ?? templates["1"];

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
  }

  private slug(value: string): string {
    return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "untitled";
  }

  private escape(value: string): string {
    return value.replace(/[&<>"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[char] ?? char));
  }
}
