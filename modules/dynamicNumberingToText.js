/* global Word, Office */

/**
 * dynamicNumberingToText.js
 * VERSION: v1.1.0
 */
const VERSION = "v1.1.0";

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const runStamp = new Date().toISOString();
    const status = (m) =>
      setStatus(`Dynamic numbering → text\n${m}\n\n${VERSION}\nRun: ${runStamp}`);

    let detected = 0;
    let converted = 0;

    status("Starting…");

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();

        // Get the paragraphs that are currently selected.
        const selParas = selection.paragraphs;
        selParas.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        if (selParas.items.length === 0) {
          status("No paragraphs detected in selection.");
          return;
        }

        // Snapshot list strings BEFORE any edits
        const items = [];
        for (let i = 0; i < selParas.items.length; i++) {
          const p = selParas.items[i];
          const li = p.listItemOrNullObject;
          const ls =
            li && li.isNullObject === false && li.listString ? String(li.listString) : "";
          if (ls) items.push({ i, ls });
        }

        // Process bottom-up so earlier paragraph references are less likely to be disturbed
        items.sort((a, b) => b.i - a.i);

        detected = items.length;

        status(`Selected paragraphs: ${selParas.items.length}\nNumbered detected: ${detected}\nConverting…`);

        // CRITICAL: sync after each paragraph so inserts do not collapse to the last line
        for (let k = 0; k < items.length; k++) {
          const { i, ls } = items[k];
          const p = selParas.items[i];

          // Insert number text + tab at the start of the paragraph
          p.getRange().insertText(ls + "\t", Word.InsertLocation.start);
          await context.sync();

          // Try to remove list formatting (best effort; may be ApiNotFound)
          try { p.detachFromList(); } catch {}
          try { p.getRange().listFormat.removeNumbers(); } catch {}
          await context.sync();

          converted++;
          status(`Converting… ${converted}/${detected}`);
        }

        status(
          "Complete.\n" +
          `Numbered detected: ${detected}\n` +
          `Converted: ${converted}\n` +
          "Note: fields are not converted in this build (numbering-only)."
        );
      });
    } catch (e) {
      const dbg = e && e.debugInfo ? JSON.stringify(e.debugInfo, null, 2) : "";
      status("ERROR:\n" + String(e?.message || e) + (dbg ? "\n\nDEBUG:\n" + dbg : ""));
      throw e;
    }
  };
})();
