import { defineConfig } from "vite";
import devCerts from "office-addin-dev-certs";

export default defineConfig(async ({ command }) => {
  const https =
    command === "serve" ? await devCerts.getHttpsServerOptions() : undefined;

  return {
    publicDir: "public",
    server: {
      host: "localhost",
      port: 3000,
      strictPort: true,
      https,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    },
    build: {
      outDir: "dist",
      emptyOutDir: true,
      rollupOptions: {
        input: "taskpane.html",
      },
    },
  };
});
