/* global Office */

// Module key -> script path (single source of truth)
const MODULES = {
  dynamicNumberingToText: "modules/dynamicNumberingToText.js",
  manualNumberingToOutlineLevels: "modules/manualNumberingToOutlineLevels.js",
  applyStyleTemplateToSelected: "modules/applyStyleTemplateToSelected.js",
  automaticCrossReferencingToSelected: "modules/automaticCrossReferencingToSelected.js",
  outlineNumberingDecimal: "modules/outlineNumberingDecimal_1_1.1_1.1.1_1.1.1.1_1.1.1.1.1.js",
  outlineNumberingLegal: "modules/outlineNumberingLegal_1_1.1_a_i_A.js",
};

Office.onReady(() => {
  const statusEl = document.getElementById("status");
  const setStatus = (m) => { if (statusEl) statusEl.textContent = m; };

  // Global toolkit registry
  window.WordToolkit = window.WordToolkit || {
    modules: {},
    register: function (key, fn) { this.modules[key] = fn; },
  };

  setStatus("Ready.");

  document.querySelectorAll("button[data-key]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const key = btn.getAttribute("data-key");
      const path = MODULES[key];

      if (!key || !path) {
        setStatus("ERROR: Unknown module key.");
        return;
      }

      setStatus(`Loading: ${key}`);

      try {
        // Load module script if not already registered
        if (!window.WordToolkit.modules[key]) {
          await loadScript(path);
        }

        const fn = window.WordToolkit.modules[key];
        if (typeof fn !== "function") {
          throw new Error(`Module did not register: ${key}`);
        }

        setStatus(`Running: ${key}`);
        await fn({ setStatus });
        setStatus(`Done: ${key}`);
      } catch (e) {
        setStatus("ERROR: " + String(e?.message || e));
      }
    });
  });
});

function loadScript(src) {
  return new Promise((resolve, reject) => {
    // Cache-bust so GitHub Pages updates are fetched
    const url = `${src}?v=${Date.now()}`;

    const s = document.createElement("script");
    s.src = url;
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Failed to load script: ${src}`));
    document.head.appendChild(s);
  });
}
