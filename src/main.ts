import "./ui/styles.css";
import "./ui/writer.css";
import { Office } from "./core/Office";

function showStartupError(error: unknown): void {
  const root = document.getElementById("app");
  if (!root) return;

  const message = error instanceof Error ? error.message : String(error);

  root.innerHTML = `
    <main class="boot-screen">
      <section class="boot-card">
        <h1 class="boot-brand">Oqerun Office</h1>
        <p class="boot-text">The application could not start.</p>
        <div class="boot-error">${message}</div>
        <p class="boot-text">Refresh the page after the deployment finishes.</p>
      </section>
    </main>
  `;
}

try {
  const root = document.getElementById("app");

  if (!root) {
    throw new Error("Oqerun Office root element was not found.");
  }

  const office = new Office(root);
  office.start();
} catch (error) {
  showStartupError(error);
}
