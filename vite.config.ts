import { defineConfig } from "vite";
import { resolve } from "node:path";

export default defineConfig({
  base: "/oqerun-office/",
  build: {
    rollupOptions: {
      input: {
        home: resolve(__dirname, "index.html"),
        docs: resolve(__dirname, "docs/index.html"),
        sheets: resolve(__dirname, "sheets/index.html"),
        slides: resolve(__dirname, "slides/index.html"),
      },
    },
  },
});
