import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./index.css";

const root = document.getElementById("root");
const OFFICE_READY_TIMEOUT_MS = 15_000;

function showBootError(title: string, body: string): void {
  if (!root) return;
  root.innerHTML = `<div class="boot-error"><h2>${title}</h2>${body}</div>`;
}

function renderApp(): void {
  if (!root) return;
  createRoot(root).render(
    <StrictMode>
      <App />
    </StrictMode>,
  );
}

function boot(): void {
  if (!root) return;

  if (typeof Office === "undefined") {
    renderApp();
    return;
  }

  let started = false;
  const timeoutId = window.setTimeout(() => {
    if (started) return;
    showBootError(
      "Office did not initialize",
      "<p>Office.onReady did not complete. Close Word, restart <code>npm run dev</code>, then run <code>npm start</code> again.</p>",
    );
  }, OFFICE_READY_TIMEOUT_MS);

  Office.onReady(() => {
    started = true;
    window.clearTimeout(timeoutId);
    try {
      renderApp();
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown startup error";
      showBootError(
        "Failed to start",
        `<p>${message}</p><p>Try a modern Office build (Microsoft 365 / Office 2019+) with WebView2.</p>`,
      );
    }
  });
}

boot();