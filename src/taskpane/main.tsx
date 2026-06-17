import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./index.css";

const root = document.getElementById("root");

function renderApp() {
  if (!root) return;
  createRoot(root).render(
    <StrictMode>
      <App />
    </StrictMode>,
  );
}

if (typeof Office !== "undefined") {
  Office.onReady(() => {
    renderApp();
  });
} else {
  renderApp();
}