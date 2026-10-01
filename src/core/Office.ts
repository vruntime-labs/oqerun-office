import { Writer } from "../writer/Writer";

export type OfficeModule = "writer" | "sheets" | "slides";

export class Office {
  private readonly root: HTMLElement | null;
  constructor(root: HTMLElement | null) { this.root = root; }

  start(): void {
    if (!this.root) throw new Error("Oqerun Office root element was not found.");
    this.renderHome();
    this.root.addEventListener("office:navigate", this.handleNavigation);
  }

  private handleNavigation = (event: Event): void => {
    if ((event as CustomEvent<string>).detail === "home") this.renderHome();
  };

  private renderHome(): void {
    this.root!.innerHTML = `
      <main class="office-shell">
        <header class="office-header"><div><div class="office-brand">Oqerun Office</div><div class="office-subtitle">Office Suite</div></div><div class="office-version">0.1.0</div></header>
        <section class="office-welcome"><p class="eyebrow">VRuntime Labs</p><h1>One suite. Three workspaces.</h1><p class="welcome-copy">Create documents, spreadsheets and presentations from one unified Oqerun application.</p>
          <div class="module-grid">${this.moduleCard("writer","Writer","Documents, rich text, tables and images.","W")}${this.moduleCard("sheets","Sheets","Spreadsheets, formulas, tables and charts.","S")}${this.moduleCard("slides","Slides","Presentations, layouts, shapes and themes.","P")}</div>
        </section>
        <footer class="office-footer"><span>Open source • Designed for Oqerun OS</span><span>VRuntime Labs</span></footer>
      </main>`;
    this.root.querySelectorAll<HTMLButtonElement>("[data-module]").forEach((button) => {
      button.addEventListener("click", () => this.openModule(button.dataset.module as OfficeModule));
    });
  }

  private openModule(module: OfficeModule): void {
    if (module === "writer") { new Writer(this.root!).render(); return; }
    this.showComingSoon(module);
  }

  private showComingSoon(module: OfficeModule): void {
    const name = module === "sheets" ? "Sheets" : "Slides";
    this.root!.innerHTML = `
      <main class="module-placeholder"><button class="writer-back" data-action="back" type="button">← Oqerun Office</button>
        <div class="placeholder-icon">${module === "sheets" ? "S" : "P"}</div><p class="eyebrow">Oqerun Office</p><h1>Oqerun ${name}</h1>
        <p>This workspace is reserved for the next module implementation. The unified application shell is already ready for it.</p>
        <button class="primary-tool placeholder-button" data-action="back" type="button">Back to Office</button>
      </main>`;
    this.root.querySelectorAll('[data-action="back"]').forEach((button) => button.addEventListener("click", () => this.renderHome()));
  }

  private moduleCard(module: OfficeModule, title: string, description: string, glyph: string): string {
    return `<button class="module-card" data-module="${module}" type="button"><span class="module-icon">${glyph}</span><span class="module-content"><strong>Oqerun ${title}</strong><span>${description}</span></span><span class="module-arrow">→</span></button>`;
  }
}
