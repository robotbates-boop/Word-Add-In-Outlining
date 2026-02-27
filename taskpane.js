/* global Office */

const TASKPANE_VERSION = "taskpane.js vNC-1 2026-02-27"; // update this string whenever you want
// Never-cache build id for this taskpane load
const BUILD_ID = Date.now().toString();

(function () {
  function setStatus(text) {
    const el = document.getElementById("status");
    if (el) el.textContent = String(text ?? "");
  }

  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};
  window.WordToolkit.versions = window.WordToolkit.versions || {};

  // Paths must match your repo exactly (case-sensitive)
  const MODULE_SCRIPTS = [
    ["dynamicNumberingToText", "modules/dynamicNumberingToText.js"],
    ["manualNumberingToOutlineLevels", "modules/manualNumberingToOutlineLevels.js"],
    ["applyStyleTemplateToSelected", "modules/applyStyleTemplateToSelected.js"],
    ["automaticCrossReferencingToSelected", "modules/automaticCrossReferencingToSelected.js"],
    ["outlineNumberingDecimal", "modules/outlineNumberingDecimal.js"],
    ["outlineNumberingLegal", "modules/outlineNumberingLegal.js"],
  ];

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      // NEVER CACHED: each taskpane load gets a fresh BUILD_ID
      s.src = `${src}?v=${encodeURIComponent(BUILD_ID)}`;
      s.onload = resolve;
      s.onerror = () => reject(new Error(`Failed to load script: ${src}`));
      document.head.appendChild(s);
    });
  }

  async function loadAllModules() {
    // Reset registries so we can see what loaded THIS time
    window.WordToolkit.modules = {};
    window.WordToolkit.versions = {};

    setStatus(
      `Loading modules...\n\n` +
        `${TASKPANE_VERSION}\n` +
        `BUILD_ID: ${BUILD_ID}\n` +
        `Loaded at: ${new Date().toISOString()}`
    );

    for (const [, src] of MODULE_SCRIPTS) {
      setStatus(
        `Loading modules...\n${src}\n\n` +
          `${TASKPANE_VERSION}\nBUILD_ID: ${BUILD_ID}\n` +
          `Loaded at: ${new Date().toISOString()}`
      );
      await loadScript(src);
    }

    const keys = Object.keys(window.WordToolkit.modules);
    const versions = window.WordToolkit.versions || {};

    let vlines = "";
    for (const k of keys.sort()) {
      vlines += `${k}: ${versions[k] || "(no version reported)"}\n`;
    }

    setStatus(
      `Ready.\n\n` +
        `${TASKPANE_VERSION}\n` +
        `BUILD_ID: ${BUILD_ID}\n` +
        `Loaded at: ${new Date().toISOString()}\n\n` +
        `Loaded keys: ${keys.join(", ") || "(none)"}\n\n` +
        `Module versions:\n${vlines || "(none)"}`
    );
  }

  async function runModule(key) {
    const fn = window.WordToolkit?.modules?.[key];
    if (typeof fn !== "function") {
      throw new Error(
        `Module not found or not a function: ${key}\n` +
          `Loaded keys: ${Object.keys(window.WordToolkit.modules).join(", ") || "(none)"}`
      );
    }
    await fn({ setStatus });
  }

  function wireButtons() {
    const reloadBtn = document.getElementById("reloadBtn");
    if (reloadBtn) {
      reloadBtn.addEventListener("click", () => location.reload());
    }

    const clearCacheBtn = document.getElementById("clearCacheBtn");
    if (clearCacheBtn) {
      clearCacheBtn.addEventListener("click", async () => {
        // In never-cache mode, just reload modules again
        await loadAllModules();
      });
    }

    document.querySelectorAll("button[data-key]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const key = btn.getAttribute("data-key");
        if (!key) return;

        try {
          setStatus(
            `Running: ${key}\n\n` +
              `${TASKPANE_VERSION}\n` +
              `BUILD_ID: ${BUILD_ID}\n` +
              `Time: ${new Date().toISOString()}`
          );
          await runModule(key);
        } catch (e) {
          setStatus(`ERROR running ${key}:\n${String(e?.message || e)}`);
          throw e;
        }
      });
    });
  }

  Office.onReady(async () => {
    try {
      wireButtons();
      await loadAllModules();
    } catch (e) {
      setStatus(`Startup error:\n${String(e?.message || e)}`);
      throw e;
    }
  });
})();
