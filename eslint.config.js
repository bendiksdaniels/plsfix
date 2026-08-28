import tseslint from "typescript-eslint";
export default tseslint.config(
  {
    ignores: [
      "**/dist/**",
      "**/node_modules/**",
      "server/**",
      "docs/**",
      ".claude/**",
      ".superpowers/**",
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
