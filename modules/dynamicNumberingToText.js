/* global Word, Office */

/**
 * dynamicNumberingToText.js
 * VERSION: v1.2.0 (wrapper cleanup)
 */
const VERSION = "v1.2.0";
const WRAPPER_TAG = "WordToolkit_DNTT_WRAPPER";

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const runStamp = new Date().toISOString();
    const status = (m) =>
      setStatus(`Dynamic numbering → text\n${m}\n\n${VERSION}\nRun: ${runStamp}`);

    status("Starting…");

    let detected = 0;
    let applied = 0;

    try {
      await Word.run(async (context) => {
        const doc = context.document;

        // 0) CLEAN UP any old wrappers from previous runs (KEEP CONTENTS)
        // Old hidden wrappers are the most common cause of “one line per click”.
        const allCCs = doc.contentControls;
        allCCs.load("items, items/tag");
        await context.sync();

        const oldWrappers = allCCs.items.filter((cc) => cc.tag === WRAPPER_TAG);
        if (oldWrappers.length) {
          status(`Removing ${oldWrappers.length} old wrapper(s)…`);
          for (const cc of oldWrappers) {
            try { cc.delete(true); } catch {}
          }
          await context.sync();
        }

        // 1) Wrap CURRENT selection in a fresh wrapper to freeze scope
        const selection = doc.getSelection();
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          status("No selection detected. Select the numbered paragraphs first.");
          return;
        }

        status("Freezing selection…");
        const wrapper = selection.insertContentControl();
        wrapper.tag = WRAPPER_TAG;
        wrapper.title = `DNTT ${VERSION} ${runStamp}`;
        wrapper.appearance = "Hidden";

        const scope = wrapper.getRange();

        // 2) Load paragraphs in wrapper scope
        status("Loading paragraphs…");
        const paras = scope.paragraphs;
        paras.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        // 3) Snapshot list strings bottom-up
        const items = [];
        for (let i = 0; i < paras.items.length; i++) {
          const p = paras.items[i];
          const li = p.listItemOrNullObject;
          const ls =
            li && li.isNullObject === false && li.listString ? String(li.listString) : "";
          if (ls) items.push({ idx: i, ls });
        }
        items.sort((a, b) => b.idx - a.idx);
        detected = items.length;

        status(`Paragraphs in scope: ${paras.items.length}\nNumbered detected: ${detected}\nConverting…`);

        // 4) Apply numbering-as-text
        // Note: detach/removeNumbers are best-effort and may be ApiNotFound on your host.
        for (const it of items) {
          const p = paras.items[it.idx];

          // Insert number text + tab at paragraph start
          p.getRange().insertText(it.ls + "\t", Word.InsertLocation.start);
          applied++;

          try { p.detachFromList(); } catch {}
          try { p.getRange().listFormat.removeNumbers(); } catch {}
        }
        await context.sync();

        // 5) Remove wrapper but KEEP contents (critical)
        // This avoids leaving wrappers that break the next run.
        status("Cleaning up wrapper…");
        try { wrapper.delete(true); } catch {}
        await context.sync();

        status(
          "Complete.\n" +
          `Numbered detected: ${detected}\n` +
          `Converted: ${applied}`
        );
      });
    } catch (e) {
      const dbg = e && e.debugInfo ? JSON.stringify(e.debugInfo, null, 2) : "";
      status("ERROR:\n" + String(e?.message || e) + (dbg ? "\n\nDEBUG:\n" + dbg : ""));
      throw e;
    }
  };
})();
