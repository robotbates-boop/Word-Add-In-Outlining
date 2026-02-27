/* global Word, Office */

/**
 * dynamicNumberingToText.js
 * VERSION: v1.4.0 (tracked paragraph anchors)
 */

const VERSION = "v1.4.0";

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};
  window.WordToolkit.versions = window.WordToolkit.versions || {};
  window.WordToolkit.versions["dynamicNumberingToText"] = `dynamicNumberingToText ${VERSION}`;

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const runStamp = new Date().toISOString();
    const status = (m) =>
      setStatus(`Dynamic numbering → text\n${m}\n\n${VERSION}\nRun: ${runStamp}`);

    status("Starting…");

    let detected = 0;
    let inserted = 0;
    let detachTried = 0;
    let removeNumbersTried = 0;

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();

        // Pull paragraphs from selection
        const paras = selection.paragraphs;
        paras.load(
          "items," +
            "items/text," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        if (!paras.items.length) {
          status("No paragraphs detected in selection.");
          return;
        }

        // Snapshot targets and TRACK them before any edits
        const targets = [];
        for (let i = 0; i < paras.items.length; i++) {
          const p = paras.items[i];
          const li = p.listItemOrNullObject;
          const ls =
            li && li.isNullObject === false && li.listString ? String(li.listString) : "";

          if (ls) {
            context.trackedObjects.add(p); // critical: stabilize paragraph anchor
            targets.push({ p, ls });
          }
        }
        await context.sync();

        detected = targets.length;

        // Quick diagnostic line so you can see what it thinks it will process
        status(
          `Selection paragraphs: ${paras.items.length}\n` +
          `Numbered detected (listString): ${detected}\n` +
          `Inserting labels…`
        );

        // Bottom-up tends to reduce interference; still track objects is the main fix
        targets.reverse();

        // PASS 1: insert labels (sync per paragraph)
        for (let i = 0; i < targets.length; i++) {
          const { p, ls } = targets[i];

          // Insert at start of paragraph using Paragraph API (not range)
          p.insertText(ls + "\t", Word.InsertLocation.start);
          await context.sync();

          inserted++;
          status(`Inserted: ${inserted}/${detected}`);
        }

        // PASS 2 (optional): attempt to remove list formatting after all insertions
        // Keeping it separate reduces the chance that list operations disturb anchors.
        status("Attempting to remove list formatting (best effort)…");

        for (let i = 0; i < targets.length; i++) {
          const { p } = targets[i];

          try { p.detachFromList(); detachTried++; } catch {}
          try { p.getRange().listFormat.removeNumbers(); removeNumbersTried++; } catch {}
        }
        await context.sync();

        // Untrack objects (cleanup)
        for (const t of targets) {
          try { context.trackedObjects.remove(t.p); } catch {}
        }
        await context.sync();

        status(
          "Complete.\n" +
          `Numbered detected: ${detected}\n` +
          `Labels inserted: ${inserted}\n` +
          `detachFromList attempted: ${detachTried}\n` +
          `removeNumbers attempted: ${removeNumbersTried}`
        );
      });
    } catch (e) {
      const dbg = e && e.debugInfo ? JSON.stringify(e.debugInfo, null, 2) : "";
      status("ERROR:\n" + String(e?.message || e) + (dbg ? "\n\nDEBUG:\n" + dbg : ""));
      throw e;
    }
  };
})();
