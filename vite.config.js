import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      includeAssets: ["favicon.svg", "apple-touch-icon.png"],
      manifest: {
        name: "Bitácora de Turnos",
        short_name: "Bitácora",
        description: "Sistema de bitácora de turnos para planta frutícola",
        theme_color: "#1e293b",
        background_color: "#f8fafc",
        display: "standalone",
        scope: "./",
        start_url: "./",
        icons: [
          { src: "icon-192.png", sizes: "192x192", type: "image/png" },
          { src: "icon-512.png", sizes: "512x512", type: "image/png", purpose: "any maskable" },
        ],
      },
      workbox: {
        globPatterns: ["**/*.{js,css,html,ico,png,svg}"],
        runtimeCaching: [
          {
            urlPattern: /^https:\/\/fonts\./,
            handler: "StaleWhileRevalidate",
          },
        ],
      },
    }),
  ],
  base: "./",
  build: { outDir: "dist", assetsDir: "assets" },
});
