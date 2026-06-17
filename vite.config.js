import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// VITE_ROLE controla qué APK se construye: "jefe" (admin) o "productor".
export default defineConfig({
  plugins: [react()],
  base: "./",
  define: {
    __APP_ROLE__: JSON.stringify(process.env.VITE_ROLE || "jefe"),
  },
  build: { outDir: "dist", emptyOutDir: true },
});
