/* global Office */

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

window.onerror = (m, src, line, col) => setStatus(`JS ERROR:\n${m}\n${src}:${line}:${col}`);
window.addEventListener("unhandledrejection", (ev) =>
  setStatus("PROMISE ERROR:\n" + String(ev.reason?.message || ev.reason))
);

Office.onReady(() => {
  // Prevent double-binding if an old instance is somehow present
  if (window.__WORDTOOLS_BOUND__) {
    setStatus("Ready (already bound).");
    return;
  }
  window.__WORDTOOLS_BOUND__ = true;

  window.WordToolkit = window.WordToolkit || {
    modules: {},
    register: function (key, fn) { this.modules[key] = fn; },
  };

  document.getElementById("reloadBtn")?.addEventListener("click", () => location.reload());
  document.getElementById("clearCacheBtn")?.addEventListener("click", () => {
    window.WordToolkit.modules = {};
    setStatus("Module cache cleared.");
  });

  document.querySelectorAll("button[data-key]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const key = btn.getAttribute("data-key");
      const path = MODULES[key];
      if (!key || !path) return setStatus("ERROR: Unknown module key/path.");

      setStatus(`CLICKED: ${key}\nLoading: ${path}`);

      try {
        // Always reload module fresh
        delete window.WordToolkit.modules[key];
        await loadScript(path);

        const fn = window.WordToolkit.modules[key];
        if (typeof fn !== "function") throw new Error(`Module did not register "${key}".`);

        setStatus(`RUNNING: ${key}`);
        await fn({ setStatus });
        setStatus(`DONE: ${key}`);
      } catch (e) {
        setStatus("ERROR:\n" + String(e?.message || e));
      }
    });
  });

  setStatus("Ready.");
});

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = `${src}?v=${Date.now()}`;
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Failed to load: ${src}`));
    document.head.appendChild(s);
  });
}
