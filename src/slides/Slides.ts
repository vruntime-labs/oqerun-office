import type { OfficeDocument, PresentationContent, PresentationElement, PresentationSlide } from "../core/DocumentModel";
import { DocumentStore } from "../core/DocumentStore";
import { exportSlidesPptx, downloadText } from "../core/FormatEngine";
import { FileMenu } from "../ui/FileMenu";

export class Slides {
  private readonly root: HTMLElement;
  private readonly store = new DocumentStore();
  private readonly state: PresentationContent = {
    activeSlide: 0,
    slides: [this.newSlide("Title slide")],
  };
  private documentId: string = crypto.randomUUID();
  private createdAt = new Date().toISOString();
  private fileName = "Untitled presentation";
  private selectedElementId: string | null = null;

  constructor(root: HTMLElement) { this.root = root; }

  render(): void {
    this.root.innerHTML = `
      <main class="slides-shell">
        <header class="slides-topbar">
          <button class="writer-back" data-action="back" type="button">← Oqerun Office</button>
          <div class="slides-title"><span class="slides-app-icon">P</span><div><strong>Oqerun Slides</strong><span data-file-name>Untitled presentation</span></div></div>
          <span class="slides-status" data-status>Ready</span>
        </header>
        <div class="slides-toolbar">
          <button data-action="add" type="button">＋ Slide</button>
          <button data-action="duplicate" type="button">Duplicate</button>
          <button data-action="delete" type="button">Delete</button>
          <button data-action="up" type="button">↑</button>
          <button data-action="down" type="button">↓</button>
          <button data-action="text" type="button">Text box</button>
          <button data-action="shape" type="button">Shape</button>
          <button data-action="image" type="button">Image</button>
          <button data-action="edit-element" type="button">Edit selected</button>
          <button data-action="delete-element" type="button">Delete selected</button>
          <button data-action="present" type="button">Present</button>
          <button data-action="save" class="primary-tool" type="button">Save</button>
          <input data-image-input type="file" accept="image/*" hidden>
        </div>
        <div class="slides-options">
          <label>Layout
            <select data-option="layout"><option value="title-body">Title + body</option><option value="title">Title only</option><option value="blank">Blank</option></select>
          </label>
          <label>Theme
            <select data-option="theme"><option value="light">Light</option><option value="dark">Dark</option><option value="midnight">Midnight</option><option value="paper">Paper</option></select>
          </label>
          <label>Transition
            <select data-option="transition"><option value="none">None</option><option value="fade">Fade</option><option value="slide">Slide</option><option value="zoom">Zoom</option></select>
          </label>
          <label>Background <input data-option="background" type="color" value="#ffffff"></label>
        </div>
        <section class="slides-workspace">
          <aside class="slide-strip" data-strip></aside>
          <article class="slide-canvas">
            <div class="slide-card" data-slide-card>
              <input class="slide-heading" data-field="title" />
              <textarea class="slide-body" data-field="body"></textarea>
              <div class="slide-elements" data-elements></div>
            </div>
            <aside class="speaker-notes"><label>Speaker notes</label><textarea data-field="notes" placeholder="Add notes for this slide…"></textarea></aside>
            <aside class="slide-element-editor" data-element-editor hidden>
              <div class="slide-element-editor-heading"><strong>Selected element</strong><button data-action="close-element-editor" type="button">×</button></div>
              <label>X <input data-element-field="x" type="number" min="0" max="100" step="1"></label>
              <label>Y <input data-element-field="y" type="number" min="0" max="100" step="1"></label>
              <label>Width <input data-element-field="width" type="number" min="1" max="100" step="1"></label>
              <label>Height <input data-element-field="height" type="number" min="1" max="100" step="1"></label>
              <label>Font <input data-element-field="fontSize" type="number" min="8" max="96" step="1"></label>
              <label>Text <input data-element-field="value" type="text"></label>
            </aside>
          </article>
        </section>
      </main>`;
    this.bind();
    this.renderStrip();
    this.loadActive();
    this.mountFileMenu();
  }

  async openPath(path: string): Promise<void> {
    const document = await this.store.load<PresentationContent>(path);
    this.documentId = document.metadata.id;
    this.createdAt = document.metadata.createdAt;
    this.fileName = document.metadata.name;
    this.state.activeSlide = Math.max(0, Math.min(document.content.activeSlide, document.content.slides.length - 1));
    this.state.slides.splice(0, this.state.slides.length, ...document.content.slides.map((slide) => ({
      ...slide,
      transition: slide.transition ?? "none",
      elements: slide.elements ?? [],
    })));
    this.updateFileName();
    this.renderStrip();
    this.loadActive();
    this.setStatus("Opened");
  }

  private mountFileMenu(): void {
    new FileMenu(this.root, {
      module: "slides",
      title: "Oqerun Slides",
      acceptedTypes: ".oqslide,.pptx",
      onHome: () => this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true })),
      onNew: () => this.newPresentation(),
      onSave: () => void this.save(),
      onSaveAs: () => void this.saveAs(),
      onDownload: () => void this.download(),
      onPrint: () => window.print(),
      onOpenRecent: (path) => void this.openPath(path),
      onOpenFile: (file) => void this.importFile(file),
    }).mount();
  }

  public async importFile(file: File): Promise<void> {
    if (!file.name.toLowerCase().endsWith(".oqslide")) {
      window.alert("Oqerun Slides currently imports .oqslide files. PPTX export is available from Download.");
      return;
    }
    try {
      const parsed = JSON.parse(await file.text()) as OfficeDocument<PresentationContent>;
      if (!parsed?.content?.slides || !parsed.metadata) throw new Error("Invalid Oqerun presentation");
      const safeName = this.safeName(parsed.metadata.name);
      const path = await this.store.save({
        ...parsed,
        metadata: { ...parsed.metadata, path: `/home/user/Documents/${safeName}.oqslide` },
      });
      this.render();
      await this.openPath(path);
    } catch {
      window.alert("The selected Oqerun presentation could not be opened.");
    }
  }

  private bind(): void {
    this.root.querySelectorAll<HTMLInputElement | HTMLTextAreaElement>("[data-field]").forEach((field) =>
      field.addEventListener("input", () => {
        const slide = this.active();
        if (field.dataset.field === "title") slide.title = field.value;
        else if (field.dataset.field === "body") slide.body = field.value;
        else slide.notes = field.value;
        this.renderStrip();
        this.setStatus("Unsaved changes");
      }),
    );

    this.root.querySelectorAll<HTMLSelectElement>("[data-option]").forEach((control) =>
      control.addEventListener("change", () => {
        const slide = this.active();
        if (control.dataset.option === "layout") slide.layout = control.value as PresentationSlide["layout"];
        if (control.dataset.option === "theme") slide.theme = control.value as PresentationSlide["theme"];
        if (control.dataset.option === "transition") slide.transition = control.value as PresentationSlide["transition"];
        this.renderActiveCard();
        this.setStatus("Unsaved changes");
      }),
    );

    this.root.querySelector<HTMLInputElement>('[data-option="background"]')?.addEventListener("input", (event) => {
      this.active().background = (event.currentTarget as HTMLInputElement).value;
      this.renderActiveCard();
      this.setStatus("Unsaved changes");
    });

    this.root.querySelector('[data-action="back"]')?.addEventListener("click", () =>
      this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true })),
    );

    this.root.querySelector('[data-action="add"]')?.addEventListener("click", () => {
      this.state.slides.splice(this.state.activeSlide + 1, 0, this.newSlide("New slide"));
      this.state.activeSlide += 1;
      this.selectedElementId = null;
      this.renderStrip(); this.loadActive();
    });

    this.root.querySelector('[data-action="duplicate"]')?.addEventListener("click", () => {
      const source = this.active();
      this.state.slides.splice(this.state.activeSlide + 1, 0, {
        ...source,
        id: crypto.randomUUID(),
        title: `${source.title} copy`,
        elements: source.elements.map((element) => ({ ...element, id: crypto.randomUUID() })),
      });
      this.state.activeSlide += 1;
      this.selectedElementId = null;
      this.renderStrip(); this.loadActive();
    });

    this.root.querySelector('[data-action="delete"]')?.addEventListener("click", () => {
      if (this.state.slides.length === 1) { window.alert("A presentation needs at least one slide."); return; }
      this.state.slides.splice(this.state.activeSlide, 1);
      this.state.activeSlide = Math.max(0, this.state.activeSlide - 1);
      this.selectedElementId = null;
      this.renderStrip(); this.loadActive();
    });

    this.root.querySelector('[data-action="up"]')?.addEventListener("click", () => this.moveSlide(-1));
    this.root.querySelector('[data-action="down"]')?.addEventListener("click", () => this.moveSlide(1));
    this.root.querySelector('[data-action="text"]')?.addEventListener("click", () => this.addTextBox());
    this.root.querySelector('[data-action="shape"]')?.addEventListener("click", () => this.addShape());
    this.root.querySelector('[data-action="edit-element"]')?.addEventListener("click", () => this.editSelectedElement());
    this.root.querySelector('[data-action="delete-element"]')?.addEventListener("click", () => this.deleteSelectedElement());

    const imageInput = this.root.querySelector<HTMLInputElement>("[data-image-input]");
    this.root.querySelector('[data-action="image"]')?.addEventListener("click", () => imageInput?.click());
    imageInput?.addEventListener("change", () => {
      const file = imageInput.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        this.active().elements.push({
          id: crypto.randomUUID(), type: "image", value: String(reader.result),
          x: 15, y: 60, width: 30, height: 25,
        });
        this.selectedElementId = this.active().elements.at(-1)?.id ?? null;
        this.renderActiveCard(); this.openElementEditor(); this.setStatus("Image added");
      };
      reader.readAsDataURL(file); imageInput.value = "";
    });

    this.root.querySelector('[data-action="present"]')?.addEventListener("click", () => this.present());
    this.root.querySelector('[data-action="save"]')?.addEventListener("click", () => void this.save());
    this.root.querySelector('[data-action="close-element-editor"]')?.addEventListener("click", () => this.closeElementEditor());

    this.root.querySelectorAll<HTMLInputElement>("[data-element-field]").forEach((field) =>
      field.addEventListener("input", () => this.updateSelectedElementField(field)),
    );
  }

  private active(): PresentationSlide { return this.state.slides[this.state.activeSlide]; }

  private renderStrip(): void {
    const strip = this.root.querySelector<HTMLElement>("[data-strip]");
    if (!strip) return;
    strip.innerHTML = this.state.slides.map((slide, index) =>
      `<button class="slide-thumb ${index === this.state.activeSlide ? "active" : ""}" data-slide="${index}" type="button"><span>${index + 1}</span><strong>${this.escape(slide.title)}</strong></button>`
    ).join("");
    strip.querySelectorAll<HTMLButtonElement>("[data-slide]").forEach((button) =>
      button.addEventListener("click", () => {
        this.state.activeSlide = Number(button.dataset.slide);
        this.selectedElementId = null;
        this.renderStrip(); this.loadActive();
      }),
    );
  }

  private loadActive(): void {
    const slide = this.active();
    const title = this.root.querySelector<HTMLInputElement>('[data-field="title"]');
    const body = this.root.querySelector<HTMLTextAreaElement>('[data-field="body"]');
    const notes = this.root.querySelector<HTMLTextAreaElement>('[data-field="notes"]');
    if (title) title.value = slide.title;
    if (body) body.value = slide.body;
    if (notes) notes.value = slide.notes;
    const layout = this.root.querySelector<HTMLSelectElement>('[data-option="layout"]');
    const theme = this.root.querySelector<HTMLSelectElement>('[data-option="theme"]');
    const transition = this.root.querySelector<HTMLSelectElement>('[data-option="transition"]');
    const background = this.root.querySelector<HTMLInputElement>('[data-option="background"]');
    if (layout) layout.value = slide.layout;
    if (theme) theme.value = slide.theme;
    if (transition) transition.value = slide.transition ?? "none";
    if (background) background.value = slide.background;
    this.renderActiveCard();
    this.setStatus(`Slide ${this.state.activeSlide + 1} of ${this.state.slides.length}`);
  }

  private renderActiveCard(): void {
    const slide = this.active();
    const card = this.root.querySelector<HTMLElement>("[data-slide-card]");
    if (!card) return;
    card.dataset.theme = slide.theme;
    card.style.background = slide.background;
    const body = this.root.querySelector<HTMLTextAreaElement>('[data-field="body"]');
    if (body) body.style.display = slide.layout === "title" || slide.layout === "blank" ? "none" : "block";
    const title = this.root.querySelector<HTMLInputElement>('[data-field="title"]');
    if (title) title.style.display = slide.layout === "blank" ? "none" : "block";
    const elements = this.root.querySelector<HTMLElement>("[data-elements]");
    if (!elements) return;

    elements.innerHTML = slide.elements.map((element) => {
      const left = element.x ?? 10, top = element.y ?? 65, width = element.width ?? 25, height = element.height ?? 15;
      const selected = element.id === this.selectedElementId ? " selected" : "";
      const rotation = element.rotation ?? 0;
      const fontSize = element.fontSize ?? 18;
      const color = element.color ?? "#182033";
      if (element.type === "image") {
        return `<img class="slide-element slide-element-image${selected}" data-element="${element.id}" src="${this.escapeAttribute(element.value)}" style="left:${left}%;top:${top}%;width:${width}%;height:${height}%;transform:rotate(${rotation}deg);" alt="Slide image">`;
      }
      if (element.type === "shape") {
        return `<div class="slide-element slide-element-shape${selected}" data-element="${element.id}" style="left:${left}%;top:${top}%;width:${width}%;height:${height}%;transform:rotate(${rotation}deg);font-size:${fontSize}px;color:${this.escapeAttribute(color)};">${this.escape(element.value)}</div>`;
      }
      return `<div class="slide-element slide-element-text${selected}" data-element="${element.id}" style="left:${left}%;top:${top}%;width:${width}%;height:${height}%;transform:rotate(${rotation}deg);font-size:${fontSize}px;color:${this.escapeAttribute(color)};">${this.escape(element.value)}</div>`;
    }).join("");

    elements.querySelectorAll<HTMLElement>("[data-element]").forEach((element) =>
      element.addEventListener("click", (event) => {
        event.stopPropagation();
        this.selectedElementId = element.dataset.element ?? null;
        this.renderActiveCard();
        this.openElementEditor();
      }),
    );
    card.addEventListener("click", () => {
      this.selectedElementId = null;
      this.closeElementEditor();
      this.renderActiveCard();
    });
    if (this.selectedElementId) this.openElementEditor();
  }

  private moveSlide(direction: -1 | 1): void {
    const target = this.state.activeSlide + direction;
    if (target < 0 || target >= this.state.slides.length) return;
    const [slide] = this.state.slides.splice(this.state.activeSlide, 1);
    this.state.slides.splice(target, 0, slide);
    this.state.activeSlide = target;
    this.renderStrip(); this.loadActive();
  }

  private addTextBox(): void {
    const value = window.prompt("Text box content", "New text");
    if (!value) return;
    const element: PresentationElement = { id: crypto.randomUUID(), type: "text", value, x: 10, y: 70, width: 35, height: 12, fontSize: 18, color: "#182033" };
    this.active().elements.push(element);
    this.selectedElementId = element.id;
    this.renderActiveCard(); this.openElementEditor(); this.setStatus("Text box added");
  }

  private addShape(): void {
    const value = window.prompt("Shape label", "Shape");
    if (!value) return;
    const element: PresentationElement = { id: crypto.randomUUID(), type: "shape", value, x: 60, y: 65, width: 24, height: 16, fontSize: 18, color: "#ffffff", background: "#5d78d6" };
    this.active().elements.push(element);
    this.selectedElementId = element.id;
    this.renderActiveCard(); this.openElementEditor(); this.setStatus("Shape added");
  }

  private selectedElement(): PresentationElement | null {
    return this.active().elements.find((element) => element.id === this.selectedElementId) ?? null;
  }

  private openElementEditor(): void {
    const panel = this.root.querySelector<HTMLElement>("[data-element-editor]");
    const element = this.selectedElement();
    if (!panel || !element) return;
    panel.hidden = false;
    (["x", "y", "width", "height", "fontSize"] as const).forEach((key) => {
      const field = this.root.querySelector<HTMLInputElement>(`[data-element-field="${key}"]`);
      if (field) field.value = String(element[key] ?? (key === "fontSize" ? 18 : 10));
    });
    const value = this.root.querySelector<HTMLInputElement>('[data-element-field="value"]');
    if (value) value.value = element.value;
  }

  private closeElementEditor(): void {
    const panel = this.root.querySelector<HTMLElement>("[data-element-editor]");
    if (panel) panel.hidden = true;
  }

  private updateSelectedElementField(field: HTMLInputElement): void {
    const element = this.selectedElement();
    if (!element) return;
    const key = field.dataset.elementField;
    if (key === "value") element.value = field.value;
    else if (["x", "y", "width", "height", "fontSize"].includes(key ?? "")) (element as unknown as Record<string, unknown>)[key!] = Number(field.value);
    this.renderActiveCard();
    this.setStatus("Unsaved changes");
  }

  private editSelectedElement(): void {
    const element = this.selectedElement();
    if (!element) {
      window.alert("Select a text box, shape or image first.");
      return;
    }
    if (element.type !== "image") {
      const value = window.prompt("Element text", element.value);
      if (value !== null) element.value = value;
    }
    this.openElementEditor();
    this.renderActiveCard();
    this.setStatus("Unsaved changes");
  }

  private deleteSelectedElement(): void {
    if (!this.selectedElementId) return;
    this.active().elements = this.active().elements.filter((element) => element.id !== this.selectedElementId);
    this.selectedElementId = null;
    this.closeElementEditor();
    this.renderActiveCard();
    this.setStatus("Element deleted");
  }

  private async saveAs(): Promise<void> {
    const value = window.prompt("Save As — presentation name", this.fileName);
    if (!value?.trim()) return;
    this.fileName = value.trim().replace(/\.oqslide$/i, "");
    await this.save();
  }

  private async download(): Promise<void> {
    const choice = (window.prompt("Download format:\n1. PowerPoint (.pptx)\n2. Oqerun (.oqslide)", "1") ?? "1").trim();
    if (choice === "1") {
      await exportSlidesPptx(this.state, this.fileName);
      this.setStatus("PPTX downloaded");
      return;
    }
    const payload = JSON.stringify(this.currentDocument(), null, 2);
    downloadText(payload, this.safeName(this.fileName) + ".oqslide", "application/json");
    this.setStatus("Oqerun presentation downloaded");
  }

  private async save(): Promise<void> {
    await this.store.save(this.currentDocument());
    this.updateFileName();
    this.setStatus("Saved locally");
  }

  private currentDocument(): OfficeDocument<PresentationContent> {
    return {
      metadata: {
        id: this.documentId,
        name: this.fileName,
        path: `/home/user/Documents/${this.safeName(this.fileName)}.oqslide`,
        mimeType: "application/x-oqerun-presentation",
        createdAt: this.createdAt,
        modifiedAt: new Date().toISOString(),
      },
      content: this.state,
    };
  }

  private present(): void {
    const payload = JSON.stringify(this.state.slides).replace(/</g, "\\u003c");
    const popup = window.open("", "_blank");
    if (!popup) return;
    popup.document.write(`<!doctype html><title>Oqerun Presentation</title>
      <style>
        *{box-sizing:border-box}body{margin:0;background:#070d18;color:#fff;font-family:system-ui,sans-serif;overflow:hidden}
        #stage{height:100vh;display:grid;place-items:center;padding:5vw}.slide{position:relative;width:min(1200px,92vw);aspect-ratio:16/9;padding:8%;box-shadow:0 30px 100px #0008;overflow:hidden}
        h1{font-size:clamp(38px,6vw,80px);margin:0 0 24px;position:relative;z-index:2}p{font-size:clamp(18px,2.5vw,32px);line-height:1.5;white-space:pre-wrap;position:relative;z-index:2;margin:0}
        .element{position:absolute;z-index:3;overflow:hidden}.text{padding:8px;background:#ffffff88;border-radius:6px}.shape{display:grid;place-items:center;border-radius:12px;background:#5d78d6cc;color:#fff}
        .controls{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);display:flex;gap:8px;z-index:10}.controls button{border:1px solid #ffffff33;background:#111a;color:#fff;border-radius:8px;padding:10px 16px;cursor:pointer}
        .count{position:fixed;top:20px;right:20px;color:#ffffffaa;z-index:10}
      </style>
      <div id="stage"></div><div class="controls"><button id="prev">←</button><button id="next">→</button><button id="full">Fullscreen</button></div><div class="count" id="count"></div>
      <script>
        const slides=${payload}; let index=${this.state.activeSlide};
        const stage=document.getElementById("stage"),count=document.getElementById("count");
        function esc(v){return String(v).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]||c))}
        function draw(){const s=slides[index];let html='<main class="slide" style="background:'+esc(s.background)+';color:'+(s.theme==="light"||s.theme==="paper"?"#182033":"#f5f7ff")+'">';if(s.layout!=="blank")html+='<h1>'+esc(s.title)+'</h1>';if(s.layout==="title-body")html+='<p>'+esc(s.body)+'</p>';for(const e of s.elements||[]){const x=e.x??10,y=e.y??65,w=e.width??25,h=e.height??15,fs=e.fontSize??18,c=e.color??"#182033";if(e.type==="image")html+='<img class="element" src="'+esc(e.value)+'" style="left:'+x+'%;top:'+y+'%;width:'+w+'%;height:'+h+'%;object-fit:contain">';else if(e.type==="shape")html+='<div class="element shape" style="left:'+x+'%;top:'+y+'%;width:'+w+'%;height:'+h+'%;font-size:'+fs+'px">'+esc(e.value)+'</div>';else html+='<div class="element text" style="left:'+x+'%;top:'+y+'%;width:'+w+'%;height:'+h+'%;font-size:'+fs+'px;color:'+esc(c)+'">'+esc(e.value)+'</div>'}html+='</main>';stage.innerHTML=html;count.textContent=(index+1)+' / '+slides.length}
        document.getElementById("prev").onclick=()=>{index=Math.max(0,index-1);draw()};document.getElementById("next").onclick=()=>{index=Math.min(slides.length-1,index+1);draw()};document.getElementById("full").onclick=()=>document.documentElement.requestFullscreen?.();document.onkeydown=e=>{if(e.key==="ArrowLeft")document.getElementById("prev").click();if(e.key==="ArrowRight")document.getElementById("next").click();if(e.key==="Escape")window.close()};draw();
      <\\/script>`);
    popup.document.close();
  }

  private newPresentation(): void {
    this.documentId = crypto.randomUUID();
    this.createdAt = new Date().toISOString();
    this.fileName = "Untitled presentation";
    this.state.activeSlide = 0;
    this.state.slides.splice(0, this.state.slides.length, this.newSlide("Title slide"));
    this.selectedElementId = null;
    this.renderStrip(); this.loadActive(); this.updateFileName();
    this.setStatus("New presentation");
  }

  private newSlide(title: string): PresentationSlide {
    return {
      id: crypto.randomUUID(), title, body: "Add content here.", notes: "",
      theme: "light", background: "#ffffff", layout: "title-body", transition: "none", elements: [],
    };
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

  private escapeAttribute(value: string): string { return this.escape(value).replace(/'/g, "&#39;"); }

  private setStatus(text: string): void {
    const status = this.root.querySelector<HTMLElement>("[data-status]");
    if (status) status.textContent = text;
  }
}
