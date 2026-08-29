// The custom-functions build pass, run before the pane build by npm run build
// and again by the predev/prestart hooks. Office injects functions.js into the
// shared runtime page as a plain script, so it cannot be an ES module and it
// cannot share a chunk with the panes: library mode emits one self-contained
// IIFE at a stable functions.js, which is the URL manifest/spec.ts publishes as
// SMT.Functions.Script.Url.
//
// It is written into public/, beside the functions.json that declares it,
// rather than straight into dist/. That is one file in one place serving both
// halves of the loop: `vite build` copies public/ into dist/ for production,
// and `vite` dev serves it at /functions.js, which the dev manifest points at
// and which nothing used to answer - the sideload registered the functions and
// then 404ed on their code. public/functions.js is a build artefact and is
// gitignored. emptyOutDir stays false so this pass never wipes the checked-in
// half of public/, and publicDir stays false so it does not copy public/ into
// itself.
import { defineConfig } from "vite";

export default defineConfig({
  publicDir: false,
  build: {
    outDir: "public",
    emptyOutDir: false,
    lib: {
      entry: "src/functions/index.ts",
      formats: ["iife"],
      name: "SMTFunctions",
      fileName: () => "functions.js",
    },
  },
});
