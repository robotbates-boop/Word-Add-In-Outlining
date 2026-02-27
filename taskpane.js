/* global Office, Word */

(function () {
  // ----------------------------
  // Simple status helper
  // ----------------------------
  function setStatus(text) {
    const el = document.getElementById("status");
    if (el) el.textContent = String(text ?? "");
  }

  // ----------------------------
  // Toolkit namespace
  // ----------------------------
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  // ----------------------------
  // Module script list
  // IMPORTANT: these filenames must exist in the SAME folder as taskpane.html
  // ----------------------------
  const MODULE_SCRIPTS = [
    "modules/dynamicNumberingToText.js",
    "modules/manualNumberingToOutlineLevels.js",
    "modules/applyStyleTemplateToSelected.js",
    "modules/automaticCrossReferencingToSelected.js",
    "modules/outlineNumberingDecimal.js",
    "modules/outlineNumberingLegal.js",
  ];

  // ----------------------------
  // Load scripts sequentially (simplest + reliable)
  // ----------------------------
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = src;
      s.onload = resolve;
      s.onerror = () => reject(new Error(`Failed to load script: ${src}`));
      document.head.appendChild(s);
    });
  }

  async function loadAllModules() {
    setStatus("Loading modules...");
    for (const src of MODULE_SCRIPTS) {
      setStatus(`Loading modules...\n${src}`);
      await loadScript(src);
    }
    setStatus("Modules loaded.");
  }

  // ----------------------------
  // Run a module by key
  // ----------------------------
  async function runModule(key) {
    const fn = window.WordToolkit?.modules?.[key];
    if (typeof fn !== "function") {
      throw new Error(
        `Module not found or not a function: ${key}\n` +
        `Loaded keys: ${Object.keys(window.WordToolkit.modules).join(", ") || "(none)"}`
      );
    }

    // IMPORTANT: pass setStatus as your module expects
    await fn({ setStatus });
  }

  // ----------------------------
  // Wire UI
  // ----------------------------
  function wireButtons() {
    const reloadBtn = document.getElementById("reloadBtn");
    if (reloadBtn) {
      reloadBtn.addEventListener("click", () => location.reload());
    }

    const clearCacheBtn = document.getElementById("clearCacheBtn");
    if (clearCacheBtn) {
      clearCacheBtn.addEventListener("click", () => {
        // Clear module functions only (simple + safe)
        window.WordToolkit.modules = {};
        setStatus("Module cache cleared.\nReload taskpane to re-load modules.");
      });
    }

    // All module buttons
    document.querySelectorAll("button[data-key]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const key = btn.getAttribute("data-key");
        if (!key) return;

        try {
          setStatus(`Running: ${key}`);
          await runModule(key);
        } catch (e) {
          setStatus(`ERROR running ${key}:\n${String(e?.message || e)}`);
          throw e;
        }
      });
    });
  }

  // ----------------------------
  // Start
  // ----------------------------
  Office.onReady(async () => {
    try {
      wireButtons();
      await loadAllModules();
      setStatus("Ready.");
    } catch (e) {
      setStatus(`Startup error:\n${String(e?.message || e)}`);
      throw e;
    }
  });
})();
