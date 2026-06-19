function sanitizeKeyPart(value: string): string {
  return value.trim().replace(/[^a-zA-Z0-9._-]+/g, "_").slice(0, 200);
}

function simpleHash(value: string): string {
  var hash = 0;
  for (var i = 0; i < value.length; i++) {
    hash = (hash << 5) - hash + value.charCodeAt(i);
    hash = hash & hash;
  }
  return Math.abs(hash).toString(16).slice(0, 24);
}

var UNSAVED_SESSION_KEY = "msword-aichat:active-unsaved-key";

function getOrCreateUnsavedKey(): string {
  var existing = sessionStorage.getItem(UNSAVED_SESSION_KEY);
  if (existing) {
    return existing;
  }

  var created = "unsaved:" + Date.now();
  sessionStorage.setItem(UNSAVED_SESSION_KEY, created);
  return created;
}

export async function getDocumentKey(): Promise<string> {
  if (typeof Office === "undefined") {
    return "browser";
  }

  var url = Office.context.document.url ? Office.context.document.url.trim() : "";
  if (url) {
    return "url:" + simpleHash(url);
  }

  return getOrCreateUnsavedKey();
}

export function formatDocumentKeyLabel(docKey: string): string {
  if (docKey === "browser") return "Browser preview";
  if (docKey === "loading") return "Loading…";
  if (docKey.indexOf("unsaved:") === 0) return "Unsaved document";
  if (docKey.indexOf("url:") === 0) return "Saved document";
  return sanitizeKeyPart(docKey);
}