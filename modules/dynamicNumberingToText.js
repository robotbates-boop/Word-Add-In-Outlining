/* global Word, Office */

/**
 * dynamicNumberingToText.js
 * VERSION: v1.5.0 (verify removal + report style-enforced)
 */

const VERSION = "v1.5.0";

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

    let removedOk = 0;
    let removalFailedLikelyStyle = 0;

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();

        const paras = selection.paragraphs;
        paras.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString," +
            "items/style"
        );
        await context.sync();

        if (!paras.items.length) {
          status("No paragraphs detected in selection.");
          return;
        }

        // Snapshot targets and track paragraph objects for stability
        const targets = [];
        for (let i = 0; i < paras.items.length; i++) {
          const p = paras.items[i];
          const li = p.listItemOrNullObject;
          const ls =
            li && li.isNullObject === false && li.listString ? String(li.listString) : "";

          if (ls) {
            context.trackedObjects.add(p);
            targets.push({ p, ls, style: p.style || "" });
          }
        }
        await context.sync();

        targets.reverse(); // bottom-up
        detected = targets.length;

        status(
          `Selection paragraphs: ${paras.items.length}\n` +
          `Numbered detected (listString): ${detected}\n` +
          `Step 1: inserting manual numbers…`
        );

        // Step 1: Insert manual number text (sync per paragraph)
        for (let i = 0; i < targets.length; i++) {
          const { p, ls } = targets[i];
          p.insertText(ls + "\t", Word.InsertLocation.start);
          await context.sync();
          inserted++;
          status(`Inserted: ${inserted}/${detected}`);
        }

        // Step 2: Try to remove numbering and VERIFY per paragraph
        status("Step 2: removing numbering (and verifying)…");

        for (let i = 0; i < targets.length; i++) {
          const { p } = targets[i];

          // Try removal (best effort)
          try { p.detachFromList(); } catch {}
          try { p.getRange().listFormat.removeNumbers(); } catch {}
          await context.sync();

          // Verify: if listString still exists, numbering is likely style-enforced
          try {
            p.load(
              "listItemOrNullObject," +
              "listItemOrNullObject/isNullObject," +
              "listItemOrNullObject/listString"
            );
            await context.sync();

            const li = p.listItemOrNullObject;
            const still =
              li && li.isNullObject === false && li.listString ? String(li.listString) : "";

            if (still && still.length > 0) {
              removalFailedLikelyStyle++;
            } else {
              removedOk++;
            }
          } catch {
            // If we cannot verify, assume removal failed
            removalFailedLikelyStyle++;
          }

          status(
            `Verified removal: ${removedOk}/${detected}\n` +
            `Still numbered (likely style-driven): ${removalFailedLikelyStyle}/${detected}`
          );
        }

        // Untrack
        for (const t of targets) {
          try { context.trackedObjects.remove(t.p); } catch {}
        }
        await context.sync();

        status(
          "Complete.\n" +
          `Numbered detected: ${detected}\n` +
          `Manual numbers inserted: ${inserted}\n` +
          `Original numbering removed: ${removedOk}\n` +
          `Still numbered (likely style-driven headings): ${removalFailedLikelyStyle}\n\n` +
          (removalFailedLikelyStyle
            ? "Those remaining numbers are coming from the paragraph style (e.g., Heading outline numbering). Office.js cannot reliably remove style-enforced numbering without changing the style definition or switching to an unnumbered style."
            : "All numbering removed.")
        );
      });
    } catch (e) {
      const dbg = e && e.debugInfo ? JSON.stringify(e.debugInfo, null, 2) : "";
      status("ERROR:\n" + String(e?.message || e) + (dbg ? "\n\nDEBUG:\n" + dbg : ""));
      throw e;
    }
  };
})();
