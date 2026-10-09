import react from "@vitejs/plugin-react";
import fs from "fs";
import { defineConfig } from "vite";

export default defineConfig({
  plugins: [react()],
  base: "/tabs/home",
  build: {
    rollupOptions: {
      input: {
        main: "index.html",
        authCallback: "auth-callback.html",
      },
    },
  },
  esbuild: {
    tsconfigRaw: fs.readFileSync("./tsconfig.app.json"),
  },
});
