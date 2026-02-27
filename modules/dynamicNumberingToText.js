/* global Word, Office */

const CHUNK_SIZE = 200;

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const status = (m) => setStatus(`Dynamic numbering → text\n${m}`);

    status("Starting…");

    let fieldsConverted = 0;
    let numberedConverted = 0;
    let fieldsSkipped = false;

    try {
      await Word.run(async (context) => {
        const body = context.document.body;

        // A) Snapshot list labels
        status("Loading paragraphs…");
        const paragraphs = body.paragraphs;
        paragraphs.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

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

        // B) Fields -> plain text (best-effort; skip if API missing)
        try {
          status(`Found numbered paragraphs: ${listItems.length}\nLoading fields…`);
          const fields = body.fields; // may throw ApiNotFound
          fields.load("items");
          await context.sync();

          const canUnlink = Office?.context?.requirements?.isSetSupported?.("WordApiDesktop", "1.4") === true;
          const fieldArray = fields.items.slice().reverse();

          status(`Fields found: ${fieldArray.length}\nConverting fields…`);

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
          status("Fields step skipped (API not available).\nContinuing…");
        }

        // C) Numbering -> text (compatible mode: DO NOT detach/remove list formatting)
        // This will leave list formatting in place, but ensures no ApiNotFound.
        status(`Converting numbering: 0/${listItems.length}`);
        let done = 0;

        for (const it of listItems) {
          const p = paragraphs.items[it.index];

          // Insert the current list label as text at the paragraph start
          p.insertText(it.listString + "\t", Word.InsertLocation.start);
          numberedConverted++;

          done++;
          if (done % CHUNK_SIZE === 0) {
            status(`Converting numbering: ${done}/${listItems.length}`);
            await context.sync();
          }
        }

        await context.sync();

        const report =
          "REPORT: Dynamic numbering → text (compat mode)\n" +
          `Fields converted: ${fieldsConverted}${fieldsSkipped ? " (fields skipped by API)" : ""}\n` +
          `Numbered paragraphs prefixed: ${numberedConverted}\n` +
          "Note: list formatting was not removed (compat mode).";

        body.insertParagraph(report, Word.InsertLocation.end);
        await context.sync();

        status(
          "Complete.\n" +
            `Fields converted: ${fieldsConverted}${fieldsSkipped ? " (fields skipped)" : ""}\n` +
            `Numbered paragraphs prefixed: ${numberedConverted}\n` +
            "List formatting not removed (compat mode)."
        );
      });
    } catch (e) {
      status("ERROR:\n" + String(e?.message || e));
      throw e;
    }
  };
})();
