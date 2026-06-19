import * as React from "react";
import * as ReactDOM from "react-dom";
import { AppLegacy } from "./App.legacy";

function getHostLabel(): string {
  if (typeof Office === "undefined") {
    return "Browser preview (Office.js not loaded)";
  }

  var documentMode = (document as Document & { documentMode?: number }).documentMode;
  if (documentMode) {
    return "Word desktop — IE" + documentMode;
  }

  return "Word / Office host";
}

function renderApp(): void {
  var root = document.getElementById("root");
  if (!root) {
    return;
  }

  ReactDOM.render(<AppLegacy hostLabel={getHostLabel()} />, root);
}

function showBootError(message: string): void {
  var root = document.getElementById("root");
  if (!root) {
    return;
  }

  root.innerHTML =
    '<div style="padding:16px;font-family:Segoe UI,sans-serif;color:#a4262c;">' +
    "<strong>Boot error:</strong> " +
    message +
    "</div>";
}

function boot(): void {
  if (typeof Office === "undefined") {
    renderApp();
    return;
  }

  var started = false;
  var timeoutId = window.setTimeout(function () {
    if (!started) {
      showBootError("Office.onReady timed out after 15 seconds.");
    }
  }, 15000);

  Office.onReady(function () {
    started = true;
    window.clearTimeout(timeoutId);
    try {
      renderApp();
    } catch (error) {
      var message = error instanceof Error ? error.message : "Unknown startup error";
      showBootError(message);
    }
  });
}

boot();