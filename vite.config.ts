import { defineConfig } from "vite";
export default defineConfig({
  server: {
    proxy: {
      "/api": "http://127.0.0.1:7731",
      "/healthz": "http://127.0.0.1:7731",
    },
  },
  build: { chunkSizeWarningLimit: 1800 },
});
