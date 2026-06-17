import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import basicSsl from "@vitejs/plugin-basic-ssl";

export default defineConfig({
  plugins: [react(), basicSsl()],
  server: {
    port: 3000,
    strictPort: true,
    https: {},
  },
  build: {
    rollupOptions: {
      input: {
        taskpane: "taskpane.html",
        commands: "commands.html",
      },
      output: {
        manualChunks(id) {
          if (id.includes("@fluentui/react-components") || id.includes("@fluentui/react-icons")) {
            return "fluent";
          }
          if (id.includes("node_modules/react-dom") || id.includes("node_modules/react/")) {
            return "react";
          }
        },
      },
    },
  },
});