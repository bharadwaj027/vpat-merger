import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  // Relative asset paths so the same build works at / (local) and /vpat-merger/ (GitHub Pages).
  base: "./",
  plugins: [react()],
  worker: { format: "es" },
  server: {
    port: 5173,
    proxy: { "/api": "http://127.0.0.1:8000" },
  },
});
