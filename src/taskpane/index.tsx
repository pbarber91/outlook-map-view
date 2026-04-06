import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";

const container = document.getElementById("container");

if (!container) {
  throw new Error("Root container #container was not found.");
}

const root = createRoot(container);

function renderApp() {
  root.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
}

const isStandalone =
  typeof window !== "undefined" &&
  new URLSearchParams(window.location.search).get("standalone") === "1";

if (isStandalone) {
  renderApp();
} else if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady(() => {
    renderApp();
  });
} else {
  renderApp();
}