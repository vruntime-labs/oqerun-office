export class Writer {
  private readonly root: HTMLElement;
  constructor(root: HTMLElement) { this.root = root; }

  render(): void {
    this.root.innerHTML = `
      <main class="writer-shell">
        <header class="writer-topbar">
          <button class="writer-back" data-action="back" type="button">← Oqerun Office</button>
          <div class="writer-title"><span class="writer-app-icon">W</span><div><strong>Oqerun Writer</strong><span>Untitled document</span></div></div>
          <div class="writer-status" data-status>Saved locally</div>
        </header>
        <nav class="writer-toolbar" aria-label="Writer toolbar">
          <div class="toolbar-group"><button data-command="undo" type="button">↶</button><button data-command="redo" type="button">↷</button></div>
          <div class="toolbar-group"><select data-command="formatBlock" aria-label="Text style"><option value="p">Normal text</option><option value="h1">Heading 1</option><option value="h2">Heading 2</option><option value="h3">Heading 3</option></select></div>
          <div class="toolbar-group"><button data-command="bold" type="button"><b>B</b></button><button data-command="italic" type="button"><i>I</i></button><button data-command="underline" type="button"><u>U</u></button></div>
          <div class="toolbar-group"><button data-command="insertUnorderedList" type="button">• List</button><button data-command="insertOrderedList" type="button">1. List</button></div>
          <div class="toolbar-group"><button data-command="justifyLeft" type="button">≡</button><button data-command="justifyCenter" type="button">≡</button><button data-command="justifyRight" type="button">≡</button></div>
          <div class="toolbar-group toolbar-actions"><button data-action="new" type="button">New</button><button data-action="save" class="primary-tool" type="button">Save</button></div>
        </nav>
        <section class="writer-workspace"><div class="writer-ruler"></div><article class="writer-page">
          <div class="writer-editor" contenteditable="true" spellcheck="true" role="textbox" aria-label="Document editor">
            <h1>Untitled document</h1><p>Start writing with Oqerun Writer.</p><p>Use the toolbar to format your document. This editor is the first foundation of the Oqerun Office document engine.</p>
          </div>
        </article></section>
      </main>`;
    this.bind();
  }

  private bind(): void {
    const editor = this.root.querySelector<HTMLElement>(".writer-editor");
    if (!editor) return;

    this.root.querySelectorAll<HTMLElement>("[data-command]").forEach((control) => {
      control.addEventListener("click", () => {
        const command = control.dataset.command;
        if (command) document.execCommand(command, false);
        editor.focus();
        this.setStatus("Unsaved changes");
      });
    });

    this.root.querySelector<HTMLSelectElement>('select[data-command="formatBlock"]')?.addEventListener("change", (event) => {
      const target = event.currentTarget as HTMLSelectElement;
      document.execCommand("formatBlock", false, target.value);
      editor.focus();
      this.setStatus("Unsaved changes");
    });

    editor.addEventListener("input", () => this.setStatus("Unsaved changes"));

    this.root.querySelector('[data-action="save"]')?.addEventListener("click", () => {
      localStorage.setItem("oqerun-writer-draft", editor.innerHTML);
      this.setStatus("Saved locally");
    });

    this.root.querySelector('[data-action="new"]')?.addEventListener("click", () => {
      editor.innerHTML = "<h1>Untitled document</h1><p>Start writing with Oqerun Writer.</p>";
      this.setStatus("New document");
      editor.focus();
    });

    this.root.querySelector('[data-action="back"]')?.addEventListener("click", () => {
      this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true }));
    });

    const draft = localStorage.getItem("oqerun-writer-draft");
    if (draft) editor.innerHTML = draft;
  }

  private setStatus(text: string): void {
    const status = this.root.querySelector<HTMLElement>("[data-status]");
    if (status) status.textContent = text;
  }
}
