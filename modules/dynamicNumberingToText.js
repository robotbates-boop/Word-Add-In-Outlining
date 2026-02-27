/* global Word, Office */

const CHUNK_SIZE = 200;

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const status = (m) => setStatus(`Dynamic numbering → text\n${m}`);

    let fieldsConverted = 0;
    let numberedConverted = 0;
    let fieldsSkipped = false;
    let detachSucceeded = 0;
    let removeNumbersSucceeded = 0;

    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const selection = context.document.getSelection();

        // Load paragraphs once
        const selParas = selection.paragraphs;
        selParas.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        if (selParas.items.length === 0) {
          status("No paragraphs in selection.");
          return;
        }

        // Build a STABLE range covering the selection from first paragraph to last paragraph.
        // This prevents “one item per click” due to the selection collapsing after edits.
        const firstParaRange = selParas.items[0].getRange();
        const lastParaRange = selParas.items[selParas.items.length - 1].getRange();
        const stableRange = firstParaRange.expandTo(lastParaRange);

        // Reload paragraphs from the stable range (now fixed)
        const paragraphs = stableRange.paragraphs;
        paragraphs.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        status(`Paragraphs in fixed scope: ${paragraphs.items.length}`);

        // Snapshot list items (bottom-up by paragraph index)
        const listItems = [];
        for (let i = 0; i < paragraphs.items.length; i++) {
          const p = paragraphs.items[i];
          const li = p.listItemOrNullObject;
          if (li && li.isNullObject === false) {
            const ls = li.listString ? String(li.listString) : "";
            if (ls) listItems.push({ index: i, listString: ls });
          }
        }
        listItems.sort((a, b) => b.index - a.index);

        // Fields in scope (best-effort)
        try {
          const fields = stableRange.fields; // may be ApiNotFound in some hosts
          fields.load("items");
          await context.sync();

          const canUnlink =
            Office?.context?.requirements?.isSetSupported?.("WordApiDesktop", "1.4") === true;

          const fieldArray = fields.items.slice().reverse();
          status(`Fields in scope: ${fieldArray.length}\nConverting fields...`);

          let done = 0;
          for (const f of fieldArray) {
            try {
              try { f.updateResult(); } catch {}

              if (canUnlink) {
                f.unlink();
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

        // Convert numbering in scope (bottom-up)
        status(`Converting numbering: 0/${listItems.length}`);
        let doneNum = 0;

        for (const it of listItems) {
          const p = paragraphs.items[it.index];

          p.insertText(it.listString + "\t", Word.InsertLocation.start);
          numberedConverted++;

          try { p.detachFromList(); detachSucceeded++; } catch {}
          try { p.getRange().listFormat.removeNumbers(); removeNumbersSucceeded++; } catch {}

          doneNum++;
          if (doneNum % CHUNK_SIZE === 0) {
            status(`Converting numbering: ${doneNum}/${listItems.length}`);
            await context.sync();
          }
        }

        await context.sync();

        const report =
          "REPORT: Dynamic numbering → text\n" +
          `Fields converted: ${fieldsConverted}${fieldsSkipped ? " (fields skipped)" : ""}\n` +
          `Numbered paragraphs processed: ${numberedConverted}\n` +
          `detachFromList succeeded: ${detachSucceeded}\n` +
          `removeNumbers succeeded: ${removeNumbersSucceeded}`;

        body.insertParagraph(report, Word.InsertLocation.end);
        await context.sync();

        status(
          "Complete.\n" +
            `Fields converted: ${fieldsConverted}${fieldsSkipped ? " (fields skipped)" : ""}\n` +
            `Numbered paragraphs processed: ${numberedConverted}\n` +
            `detachFromList succeeded: ${detachSucceeded}\n` +
            `removeNumbers succeeded: ${removeNumbersSucceeded}`
        );
      });
    } catch (e) {
      setStatus(`Dynamic numbering → text\nERROR:\n${String(e?.message || e)}`);
      throw e;
    }
  };
})();
