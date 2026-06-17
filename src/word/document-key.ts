function sanitizeKeyPart(value: string): string {
  return value.trim().replace(/[^a-zA-Z0-9._-]+/g, "_").slice(0, 200);
}

async function hashKey(value: string): Promise<string> {
  const data = new TextEncoder().encode(value);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(digest))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("")
    .slice(0, 24);
}

const UNSAVED_SESSION_KEY = "msword-aichat:active-unsaved-key";

function getOrCreateUnsavedKey(): string {
  const existing = sessionStorage.getItem(UNSAVED_SESSION_KEY);
  if (existing) return existing;

  const created = `unsaved:${Date.now()}`;
  sessionStorage.setItem(UNSAVED_SESSION_KEY, created);
  return created;
}

export async function getDocumentKey(): Promise<string> {
  if (typeof Office === "undefined") {
    return "browser";
  }

  const url = Office.context.document.url?.trim();
  if (url) {
    return `url:${await hashKey(url)}`;
  }

  return getOrCreateUnsavedKey();
}

export function formatDocumentKeyLabel(docKey: string): string {
  if (docKey === "browser") return "Browser preview";
  if (docKey === "loading") return "Loading…";
  if (docKey.startsWith("unsaved:")) return "Unsaved document";
  if (docKey.startsWith("url:")) return "Saved document";
  return sanitizeKeyPart(docKey);
}