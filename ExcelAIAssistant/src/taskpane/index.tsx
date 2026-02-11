import "core-js/stable";
import "regenerator-runtime/runtime";
import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { ErrorBoundary } from "./components/ErrorBoundary";

/* global document, Office, module, require, HTMLElement */

const title = "Excel AI 助理";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  try {
    root?.render(
      <ErrorBoundary>
        <App title={title} />
      </ErrorBoundary>
    );
  } catch (err) {
    console.error("Rendering failed", err);
  }
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
