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

window.onerror = (m, src, line, col) => {
  setStatus(`JS ERROR:\n${m}\n${src}:${line}:${col}`);
};

window.addEventListener("unhandledrejection", (ev) => {
  setStatus("PROMISE ERROR:\n" + String(ev.reason?.message || ev.reason));
});

Office.onReady(() => {
  setStatus("Ready (taskpane.js loaded).");

  window.WordToolkit = window.WordToolkit || {
    modules: {},
    register: function (key, fn) { this.modules[key] = fn; },
  };

  document.querySelectorAll("button[data-key]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const key = btn.getAttribute("data-key");
      const path = MODULES[key];

      // This line MUST show immediately on click.
      setStatus(`CLICKED: ${key}\nPath: ${path || "(missing)"}`);

      if (!key || !path) return;

      try {
        // Force reload registration every click (avoids stale module)
        delete window.WordToolkit.modules[key];

        await loadScript(path);

        const fn = window.WordToolkit.modules[key];
        if (typeof fn !== "function") {
          throw new Error(`Module did not register "${key}".`);
        }

        setStatus(`RUNNING: ${key}`);
        await fn({ setStatus });
        setStatus(`DONE: ${key}`);
      } catch (e) {
        setStatus("ERROR:\n" + String(e?.message || e));
      }
    });
  });
});

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = `${src}?v=${Date.now()}`; // cache-bust module file
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Failed to load: ${src}`));
    document.head.appendChild(s);
  });
}
