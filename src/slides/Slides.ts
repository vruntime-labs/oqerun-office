import type { OfficeDocument, PresentationContent } from "../core/DocumentModel";
import { DocumentStore } from "../core/DocumentStore";

export class Slides {
  private readonly root: HTMLElement;
  private readonly store = new DocumentStore();
  private readonly state: PresentationContent = {
    activeSlide: 0,
    slides: [{
      id: crypto.randomUUID(),
      title: "Title slide",
      body: "Start your presentation with Oqerun Slides.",
      elements: [],
    }],
  };

  constructor(root: HTMLElement) {
    this.root = root;
  }

  render(): void {
    this.root.innerHTML = `
      <main class="slides-shell">
        <header class="slides-topbar">
          <button class="writer-back" data-action="back" type="button">← Oqerun Office</button>
          <div class="slides-title"><span class="slides-app-icon">P</span><div><strong>Oqerun Slides</strong><span>Untitled presentation</span></div></div>
          <span class="slides-status" data-status>Ready</span>
        </header>
        <div class="slides-toolbar">
          <button data-action="add" type="button">＋ Slide</button>
          <button data-action="duplicate" type="button">Duplicate</button>
          <button data-action="delete" type="button">Delete</button>
          <button data-action="present" type="button">Present</button>
          <button data-action="save" class="primary-tool" type="button">Save</button>
        </div>
        <section class="slides-workspace">
          <aside class="slide-strip" data-strip></aside>
          <article class="slide-canvas">
            <div class="slide-card">
              <input class="slide-heading" data-field="title" />
              <textarea class="slide-body" data-field="body"></textarea>
            </div>
          </article>
        </section>
      </main>`;
    this.bind();
    this.renderStrip();
    this.loadActive();
  }

  private bind(): void {
    this.root.querySelectorAll<HTMLInputElement | HTMLTextAreaElement>("[data-field]").forEach((field) =>
      field.addEventListener("input", () => {
        const slide = this.state.slides[this.state.activeSlide];
        if (field.dataset.field === "title") slide.title = field.value;
        else slide.body = field.value;
        this.renderStrip();
        this.setStatus("Unsaved changes");
      }),
    );

    this.root.querySelector('[data-action="back"]')?.addEventListener("click", () =>
      this.root.dispatchEvent(new CustomEvent("office:navigate", { detail: "home", bubbles: true })),
    );

    this.root.querySelector('[data-action="add"]')?.addEventListener("click", () => {
      this.state.slides.splice(this.state.activeSlide + 1, 0, {
        id: crypto.randomUUID(),
        title: "New slide",
        body: "Add content here.",
        elements: [],
      });
      this.state.activeSlide += 1;
      this.renderStrip();
      this.loadActive();
    });

    this.root.querySelector('[data-action="duplicate"]')?.addEventListener("click", () => {
      const source = this.state.slides[this.state.activeSlide];
      this.state.slides.splice(this.state.activeSlide + 1, 0, {
        ...source,
        id: crypto.randomUUID(),
        title: `${source.title} copy`,
      });
      this.state.activeSlide += 1;
      this.renderStrip();
      this.loadActive();
    });

    this.root.querySelector('[data-action="delete"]')?.addEventListener("click", () => {
      if (this.state.slides.length === 1) {
        window.alert("A presentation needs at least one slide.");
        return;
      }
      this.state.slides.splice(this.state.activeSlide, 1);
      this.state.activeSlide = Math.max(0, this.state.activeSlide - 1);
      this.renderStrip();
      this.loadActive();
    });

    this.root.querySelector('[data-action="present"]')?.addEventListener("click", () => this.present());
    this.root.querySelector('[data-action="save"]')?.addEventListener("click", () => void this.save());
  }

  private renderStrip(): void {
    const strip = this.root.querySelector<HTMLElement>("[data-strip]");
    if (!strip) return;
    strip.innerHTML = this.state.slides
      .map((slide, index) => `<button class="slide-thumb ${index === this.state.activeSlide ? "active" : ""}" data-slide="${index}" type="button"><span>${index + 1}</span><strong>${this.escape(slide.title)}</strong></button>`)
      .join("");

    strip.querySelectorAll<HTMLButtonElement>("[data-slide]").forEach((button) =>
      button.addEventListener("click", () => {
        this.state.activeSlide = Number(button.dataset.slide);
        this.renderStrip();
        this.loadActive();
      }),
    );
  }

  private loadActive(): void {
    const slide = this.state.slides[this.state.activeSlide];
    const title = this.root.querySelector<HTMLInputElement>('[data-field="title"]');
    const body = this.root.querySelector<HTMLTextAreaElement>('[data-field="body"]');
    if (title) title.value = slide.title;
    if (body) body.value = slide.body;
    this.setStatus(`Slide ${this.state.activeSlide + 1} of ${this.state.slides.length}`);
  }

  private async save(): Promise<void> {
    const now = new Date().toISOString();
    const document: OfficeDocument<PresentationContent> = {
      metadata: {
        id: "oqerun-slide-deck",
        name: "Untitled presentation",
        path: "/home/user/Documents/untitled-presentation.oqslide",
        mimeType: "application/x-oqerun-presentation",
        createdAt: now,
        modifiedAt: now,
      },
      content: this.state,
    };
    await this.store.save(document);
    this.setStatus("Saved locally");
  }

  private present(): void {
    const slide = this.state.slides[this.state.activeSlide];
    const popup = window.open("", "_blank", "noopener,noreferrer");
    if (!popup) return;
    popup.document.write(`<title>${this.escape(slide.title)}</title><style>body{margin:0;display:grid;place-items:center;min-height:100vh;background:#0b1320;color:#fff;font:24px system-ui}main{width:min(1100px,85vw);padding:80px}h1{font-size:64px;margin:0 0 24px}p{font-size:30px;line-height:1.5}</style><main><h1>${this.escape(slide.title)}</h1><p>${this.escape(slide.body).replace(/\n/g, "<br>")}</p></main>`);
    popup.document.close();
  }

  private escape(value: string): string {
    return value.replace(/[&<>"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[char] ?? char));
  }

  private setStatus(text: string): void {
    const status = this.root.querySelector<HTMLElement>("[data-status]");
    if (status) status.textContent = text;
  }
}
