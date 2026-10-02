import tseslint from "typescript-eslint";
export default tseslint.config(
  {
    ignores: [
      "**/dist/**",
      // Build output, not source: the custom-functions IIFE the functions:js
      // pass writes beside the functions.json that declares it.
      "public/functions.js",
      "**/node_modules/**",
      "server/**",
      "docs/**",
      ".claude/**",
      ".superpowers/**",
      // Eval'd Playwright fragments for the web rig, not modules.
      "scripts/rig/snippets/**",
      // The video tool (Rust + its HTML composition and Playwright capture,
      // shared with the Planner video engine) keeps its own style and gate:
      // `cd video && cargo test`.
      "video/**",
    ],
  },
  ...tseslint.configs.strict,
  ...tseslint.configs.stylistic,
  {
    files: ["**/*.ts"],
    rules: {
      "no-console": "error",
      "@typescript-eslint/no-explicit-any": "error",
      "@typescript-eslint/no-non-null-assertion": "off",
      "@typescript-eslint/consistent-type-imports": "error",
    },
  },
);
