import type { OfficeDocument, PresentationContent, PresentationSlide } from "../core/DocumentModel";
import { DocumentStore } from "../core/DocumentStore";
import { FileMenu } from "../ui/FileMenu";

export class Slides {
  private readonly root: HTMLElement;
  private readonly store = new DocumentStore();
  private readonly state: PresentationContent = {
    activeSlide: 0,
    slides: [this.newSlide("Title slide")],
  };
  private documentId = crypto.randomUUID();
  private createdAt = new Date().toISOString();
  private fileName = "Untitled presentation";

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
    this.state.activeSlide = document.content.activeSlide;
    this.state.slides.splice(0, this.state.slides.length, ...document.content.slides);
    const name = this.root.querySelector<HTMLElement>("[data-file-name]");
    if (name) name.textContent = document.metadata.name;
    this.renderStrip();
    this.loadActive();
    this.setStatus("Opened");
  }

  private mountFileMenu(): void {
    new FileMenu(this.root, {
      module: "slides",
      title: "Oqerun Slides",
      acceptedTypes: ".oqslide",
      onHome: () => this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true })),
      onNew: () => {
        this.documentId = crypto.randomUUID();
        this.createdAt = new Date().toISOString();
        this.fileName = "Untitled presentation";
        this.state.activeSlide = 0;
        this.state.slides.splice(0, this.state.slides.length, this.newSlide("Title slide"));
        this.renderStrip();
        this.loadActive();
        this.setStatus("New presentation");
      },
      onSave: () => void this.save(),
      onSaveAs: () => void this.saveAs(),
      onDownload: () => this.download(),
      onPrint: () => window.print(),
      onOpenRecent: (path) => void this.openPath(path),
      onOpenFile: (file) => void this.importFile(file),
    }).mount();
  }

  private async importFile(file: File): Promise<void> {
    if (!file.name.toLowerCase().endsWith(".oqslide")) {
      window.alert("Oqerun Slides currently supports .oqslide from Open From.");
      return;
    }

    try {
      const parsed = JSON.parse(await file.text()) as OfficeDocument<PresentationContent>;
      if (!parsed?.content?.slides || !parsed.metadata) throw new Error("Invalid Oqerun presentation");
      const safeName = parsed.metadata.name.trim().replace(/[^a-z0-9._-]+/gi, "-").replace(/^-+|-+$/g, "") || "untitled-presentation";
      const path = await this.store.save({
        ...parsed,
        metadata: {
          ...parsed.metadata,
          path: `/home/user/Documents/${safeName}.oqslide`,
        },
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
      this.renderStrip(); this.loadActive();
    });

    this.root.querySelector('[data-action="duplicate"]')?.addEventListener("click", () => {
      const source = this.active();
      this.state.slides.splice(this.state.activeSlide + 1, 0, {
        ...source, id: crypto.randomUUID(), title: `${source.title} copy`,
        elements: source.elements.map((element) => ({ ...element, id: crypto.randomUUID() })),
      });
      this.state.activeSlide += 1; this.renderStrip(); this.loadActive();
    });

    this.root.querySelector('[data-action="delete"]')?.addEventListener("click", () => {
      if (this.state.slides.length === 1) { window.alert("A presentation needs at least one slide."); return; }
      this.state.slides.splice(this.state.activeSlide, 1);
      this.state.activeSlide = Math.max(0, this.state.activeSlide - 1);
      this.renderStrip(); this.loadActive();
    });

    this.root.querySelector('[data-action="up"]')?.addEventListener("click", () => this.moveSlide(-1));
    this.root.querySelector('[data-action="down"]')?.addEventListener("click", () => this.moveSlide(1));
    this.root.querySelector('[data-action="text"]')?.addEventListener("click", () => this.addTextBox());
    this.root.querySelector('[data-action="shape"]')?.addEventListener("click", () => this.addShape());

    const imageInput = this.root.querySelector<HTMLInputElement>("[data-image-input]");
    this.root.querySelector('[data-action="image"]')?.addEventListener("click", () => imageInput?.click());
    imageInput?.addEventListener("change", () => {
      const file = imageInput.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        this.active().elements.push({ id: crypto.randomUUID(), type: "image", value: String(reader.result), x: 15, y: 60, width: 30, height: 25 });
        this.renderActiveCard(); this.setStatus("Image added");
      };
      reader.readAsDataURL(file); imageInput.value = "";
    });

    this.root.querySelector('[data-action="present"]')?.addEventListener("click", () => this.present());
    this.root.querySelector('[data-action="save"]')?.addEventListener("click", () => void this.save());
  }

  private active(): PresentationSlide { return this.state.slides[this.state.activeSlide]; }

  private renderStrip(): void {
    const strip = this.root.querySelector<HTMLElement>("[data-strip]");
    if (!strip) return;
    strip.innerHTML = this.state.slides.map((slide, index) =>
      `<button class="slide-thumb ${index === this.state.activeSlide ? "active" : ""}" data-slide="${index}" type="button"><span>${index + 1}</span><strong>${this.escape(slide.title)}</strong></button>`
    ).join("");
    strip.querySelectorAll<HTMLButtonElement>("[data-slide]").forEach((button) =>
      button.addEventListener("click", () => { this.state.activeSlide = Number(button.dataset.slide); this.renderStrip(); this.loadActive(); }),
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
    const background = this.root.querySelector<HTMLInputElement>('[data-option="background"]');
    if (layout) layout.value = slide.layout;
    if (theme) theme.value = slide.theme;
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
      if (element.type === "image") return `<img class="slide-element slide-element-image" src="${this.escapeAttribute(element.value)}" style="left:${left}%;top:${top}%;width:${width}%;height:${height}%;" alt="Slide image">`;
      if (element.type === "shape") return `<div class="slide-element slide-element-shape" style="left:${left}%;top:${top}%;width:${width}%;height:${height}%;">${this.escape(element.value)}</div>`;
      return `<div class="slide-element slide-element-text" style="left:${left}%;top:${top}%;width:${width}%;height:${height}%;">${this.escape(element.value)}</div>`;
    }).join("");
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
    this.active().elements.push({ id: crypto.randomUUID(), type: "text", value, x: 10, y: 70, width: 35, height: 12 });
    this.renderActiveCard(); this.setStatus("Text box added");
  }

  private addShape(): void {
    const value = window.prompt("Shape label", "Shape");
    if (!value) return;
    this.active().elements.push({ id: crypto.randomUUID(), type: "shape", value, x: 60, y: 65, width: 24, height: 16 });
    this.renderActiveCard(); this.setStatus("Shape added");
  }

  private async saveAs(): Promise<void> {
    const value = window.prompt("Save As — presentation name", this.fileName);
    if (!value?.trim()) return;
    this.fileName = value.trim().replace(/\.oqslide$/i, "");
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
    anchor.download = this.fileName.replace(/[^a-z0-9._-]+/gi, "-") + ".oqslide";
    anchor.click();
    URL.revokeObjectURL(url);
    this.setStatus("Downloaded");
  }

  private async save(): Promise<void> {
    const now = new Date().toISOString();
    const document: OfficeDocument<PresentationContent> = {
      metadata: { id: this.documentId, name: this.fileName, path: `/home/user/Documents/${this.fileName.replace(/[^a-z0-9._-]+/gi, "-")}.oqslide`, mimeType: "application/x-oqerun-presentation", createdAt: this.createdAt, modifiedAt: now },
      content: this.state,
    };
    await this.store.save(document);
    this.setStatus("Saved locally");
  }

  private present(): void {
    const payload = JSON.stringify(this.state.slides).replace(/</g, "\\u003c");
    const popup = window.open("", "_blank", "noopener,noreferrer");
    if (!popup) return;
    popup.document.write(`<!doctype html><title>Oqerun Presentation</title>
      <style>*{box-sizing:border-box}body{margin:0;background:#070d18;color:#fff;font-family:system-ui,sans-serif;overflow:hidden}#stage{height:100vh;display:grid;place-items:center;padding:5vw}.slide{width:min(1200px,92vw);aspect-ratio:16/9;padding:8%;display:flex;flex-direction:column;justify-content:center;gap:24px;box-shadow:0 30px 100px #0008}h1{font-size:clamp(38px,6vw,80px);margin:0}p{font-size:clamp(18px,2.5vw,32px);line-height:1.5;white-space:pre-wrap}.controls{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);display:flex;gap:8px}button{border:1px solid #ffffff33;background:#111a;color:#fff;border-radius:8px;padding:10px 16px;cursor:pointer}.count{position:fixed;top:20px;right:20px;color:#ffffffaa}</style>
      <div id="stage"></div><div class="controls"><button id="prev">←</button><button id="next">→</button></div><div class="count" id="count"></div>
      <script>
        const slides=${payload}; let index=${this.state.activeSlide};
        const stage=document.getElementById("stage"),count=document.getElementById("count");
        function esc(v){return String(v).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]||c))}
        function draw(){const s=slides[index];stage.innerHTML='<main class="slide" style="background:'+esc(s.background)+';color:'+(s.theme==="light"||s.theme==="paper"?"#182033":"#f5f7ff")+'"><h1>'+esc(s.title)+'</h1><p>'+esc(s.body)+'</p></main>';count.textContent=(index+1)+' / '+slides.length}
        document.getElementById("prev").onclick=()=>{index=Math.max(0,index-1);draw()};document.getElementById("next").onclick=()=>{index=Math.min(slides.length-1,index+1);draw()};document.onkeydown=e=>{if(e.key==="ArrowLeft")document.getElementById("prev").click();if(e.key==="ArrowRight")document.getElementById("next").click()};draw();
      <\/script>`);
    popup.document.close();
  }

  private newSlide(title: string): PresentationSlide {
    return { id: crypto.randomUUID(), title, body: "Add content here.", notes: "", theme: "light", background: "#ffffff", layout: "title-body", elements: [] };
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
