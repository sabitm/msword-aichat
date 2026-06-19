import { defineConfig, type UserConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig(async (): Promise<UserConfig> => {
  const plugins = [react()];
  let https: NonNullable<UserConfig["server"]>["https"];

  try {
    const devCerts = await import("office-addin-dev-certs");
    https = await devCerts.default.getHttpsServerOptions();
    console.log("Using office-addin-dev-certs for HTTPS (trusted by Office on Windows/Mac).");
  } catch (error) {
    const { default: basicSsl } = await import("@vitejs/plugin-basic-ssl");
    plugins.push(basicSsl());
    https = {};
    const message = error instanceof Error ? error.message : String(error);
    console.warn(
      `office-addin-dev-certs unavailable (${message}). Using basic-ssl fallback.`,
    );
    console.warn("On Windows, run an elevated terminal: npm run certs");
  }

  return {
    plugins,
    server: {
      port: 3000,
      strictPort: true,
      https,
    },
    build: {
      rollupOptions: {
        input: {
          taskpane: "taskpane.html",
          commands: "commands.html",
        },
        output: {
          manualChunks(id) {
            if (
              id.includes("@fluentui/react-components") ||
              id.includes("@fluentui/react-icons")
            ) {
              return "fluent";
            }
            if (id.includes("node_modules/react-dom") || id.includes("node_modules/react/")) {
              return "react";
            }
          },
        },
      },
    },
  };
});