import { existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

export function getRuntimeRoot() {
  if (process.pkg) {
    return dirname(process.execPath);
  }

  var scriptDir = dirname(fileURLToPath(import.meta.url));
  if (existsSync(join(scriptDir, "package.json"))) {
    return scriptDir;
  }
  if (existsSync(join(scriptDir, "..", "package.json"))) {
    return join(scriptDir, "..");
  }
  return scriptDir;
}