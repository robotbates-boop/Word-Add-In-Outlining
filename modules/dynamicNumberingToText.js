/* global Word, Office */

const CHUNK_SIZE = 200;

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const status = (m) => setStatus(`Dynamic numbering → text\n${m}`);

    status("Starting...");

    let fieldsConverted = 0;
    let numberedConverted = 0;
    let fieldsSkipped = false;

    let detachSucceeded = 0;
    let removeNumbersSucceeded = 0;

    try {
      await Word.run(async (context) => {
        const body = context.document.body;

        // 1) Decide target range: selection if non-empty, else whole body
        const sel = context.document.getSelection();
        sel.load("text");
        await context.sync();

        const targetRange =
          (sel.text && sel.text.trim().length > 0) ? sel : body.getRange();

        // 2) Snapshot paragraphs in the target range (safe list detection)
        status("Loading paragraphs in selection...");
        const paragraphs = targetRange.paragraphs;
        paragraphs.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        // 3) Convert fields within the same target range (best-effort)
        try {
          status(`Paragraphs in scope: ${paragraphs.items.length}\nLoading fields in scope...`);

          const fields = targetRange.fields; // may throw ApiNotFound in some hosts
          fields.load("items");
          await context.sync();

          const canUnlink =
            Office?.context?.requirements?.isSetSupported?.("WordApiDesktop", "1.4") === true;

          const fieldArray = fields.items.slice().reverse();
          status(`Fields found: ${fieldArray.length}\nConverting fields...`);

          let done = 0;
          for (const f of fieldArray) {
            try {
              try { f.updateResult(); } catch {}

              if (canUnlink) {
                f.unlink(); // desktop-only
              } else {
                const r = f.getRange();
                r.load("text");
                await context.sync();
                const t = r.text || "";
                r.insertText(t, Word.InsertLocation.replace);
                try { f.delete(); } catch {}
              }

              fieldsConverted++;
            } catch {}

            done++;
            if (done % CHUNK_SIZE === 0) {
              status(`Converting fields: ${done}/${fieldArray.length}`);
              await context.sync();
            }
          }
          await context.sync();
        } catch {
          fieldsSkipped = true;
          status("Fields step skipped (API not available).\nContinuing...");
        }

        // 4) Convert numbering bottom-up across ALL paragraphs in the target range
        status("Converting list/outline numbering (bottom-up)...");
        let done = 0;

        for (let i = paragraphs.items.length - 1; i >= 0; i--) {
          const p = paragraphs.items[i];
          const li = p.listItemOrNullObject;

          if (li && li.isNullObject === false) {
            const ls = li.listString ? String(li.listString) : "";
            if (ls) {
              // Insert the displayed label as plain text
              p.insertText(ls + "\t", Word.InsertLocation.start);
              numberedConverted++;

              // Remove list formatting (best-effort)
              try { p.detachFromList(); detachSucceeded++; } catch {}
              try { p.getRange().listFormat.removeNumbers(); removeNumbersSucceeded++; } catch {}
            }
          }

          done++;
          if (done % CHUNK_SIZE === 0) {
            status(`Processed paragraphs: ${done}/${paragraphs.items.length}`);
            await context.sync();
          }
        }

        await context.sync();

        status(
          "Complete.\n" +
            `Fields converted: ${fieldsConverted}${fieldsSkipped ? " (fields skipped)" : ""}\n` +
            `Numbered paragraphs converted: ${numberedConverted}\n` +
            `detachFromList succeeded: ${detachSucceeded}\n` +
            `removeNumbers succeeded: ${removeNumbersSucceeded}`
        );
      });
    } catch (e) {
      status("ERROR:\n" + String(e?.message || e));
      throw e;
    }
  };
})();
