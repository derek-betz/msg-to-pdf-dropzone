import { resolve } from "node:path";

import { defineConfig } from "vite";

export default defineConfig({
  base: "./",
  build: {
    emptyOutDir: true,
    outDir: resolve(__dirname, "../src/msg_to_pdf_dropzone/theater_assets")
  }
});
