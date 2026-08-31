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
