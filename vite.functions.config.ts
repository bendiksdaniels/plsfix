// Second build pass, run straight after the pane build by npm run build.
// Office injects functions.js into the shared runtime page as a plain script,
// so it cannot be an ES module and it cannot share a chunk with the panes:
// library mode emits one self-contained IIFE at a stable dist/functions.js,
// which is the URL manifest/spec.ts publishes as SMT.Functions.Script.Url.
// emptyOutDir stays false so this pass adds to the pane build instead of
// wiping it, and publicDir stays false so public/ is not copied twice.
import { defineConfig } from "vite";

export default defineConfig({
  publicDir: false,
  build: {
    outDir: "dist",
    emptyOutDir: false,
    lib: {
      entry: "src/functions/index.ts",
      formats: ["iife"],
      name: "SMTFunctions",
      fileName: () => "functions.js",
    },
  },
});
