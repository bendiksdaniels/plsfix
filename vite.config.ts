import dns from "node:dns";
import { defineConfig } from "vite";
import devCerts from "office-addin-dev-certs";

// Node 17+ resolves "localhost" to ::1 first, so vite binds IPv6 only and
// anything probing 127.0.0.1 (Office tooling, some webviews) is refused.
dns.setDefaultResultOrder("ipv4first");

export default defineConfig(async ({ command }) => {
  const https =
    command === "serve" ? await devCerts.getHttpsServerOptions() : undefined;

  return {
    // Relative base so the same build serves from any suite sub-path.
    base: command === "serve" ? "/" : "./",
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
