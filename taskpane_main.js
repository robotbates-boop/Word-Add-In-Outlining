/* global Office */

const TASKPANE_MAIN_VERSION = "taskpane_main v1";

const MODULES = {
  dynamicNumberingToText: "modules/dynamicNumberingToText.js",
  manualNumberingToOutlineLevels: "modules/manualNumberingToOutlineLevels.js",
  applyStyleTemplateToSelected: "modules/applyStyleTemplateToSelected.js",
  automaticCrossReferencingToSelected: "modules/automaticCrossReferencingToSelected.js",
  outlineNumberingDecimal: "modules/outlineNumberingDecimal_1_1.1_1.1.1_1.1.1.1_1.1.1.1.1.js",
  outlineNumberingLegal: "modules/outlineNumberingLegal_1_1.1_a_i_A.js",
};

function setStatus(msg) {
  const el = document.getElementById("status");
  if (el) el.textContent = msg;
}

window.onerror = (m, src, line, col) =>
  setStatus(`JS ERROR (${TASKPANE_MAIN_VERSION}):\n${m}\n${src}:${line}:${col}`);

window.addEventListener("unhandledrejection", (ev) =>
  setStatus(`PROMISE ERROR (${TASKPANE_MAIN_VERSION}):\n` + String(ev.reason?.message || ev.reason))
);

function ensureRegistry() {
  // Use ONE registry name: WordToolkit
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};
}

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = `${src}?v=${Date.now()}`; // always fresh
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Failed to load: ${src}`));
    document.head.appendChild(s);
  });
}

Office.onReady(() => {
  // Prevent double-binding (old cached controllers)
  if (window.__WORDTOOLS_BOUND__) {
    setStatus(`Ready (${TASKPANE_MAIN_VERSION}) [already bound]`);
    return;
  }
  window.__WORDTOOLS_BOUND__ = true;

  ensureRegistry();

  document.getElementById("reloadBtn")?.addEventListener("click", () => location.reload());
  document.getElementById("clearCacheBtn")?.addEventListener("click", () => {
    ensureRegistry();
    window.WordToolkit.modules = {};
    setStatus(`Module cache cleared (${TASKPANE_MAIN_VERSION}).`);
  });

  document.querySelectorAll("button[data-key]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      ensureRegistry();

      const key = btn.getAttribute("data-key");
      const path = MODULES[key];
      setStatus(`CLICKED (${TASKPANE_MAIN_VERSION}): ${key}\nPath: ${path || "(missing)"}`);

      if (!key || !path) return;

      try {
        // Force reload this module every click
        delete window.WordToolkit.modules[key];
        await loadScript(path);

        const fn = window.WordToolkit.modules[key];
        if (typeof fn !== "function") {
          const keys = Object.keys(window.WordToolkit.modules);
          throw new Error(`Module did not register "${key}". Registered keys: ${keys.join(", ") || "(none)"}`);
        }

        setStatus(`RUNNING (${TASKPANE_MAIN_VERSION}): ${key}`);
        await fn({ setStatus });
        setStatus(`DONE (${TASKPANE_MAIN_VERSION}): ${key}`);
      } catch (e) {
        setStatus(`ERROR (${TASKPANE_MAIN_VERSION}):\n` + String(e?.message || e));
      }
    });
  });

  setStatus(`Ready (${TASKPANE_MAIN_VERSION}).`);
});
