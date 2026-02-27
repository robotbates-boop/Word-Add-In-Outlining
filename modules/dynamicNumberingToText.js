/* global Word, Office */

// Dynamic numbering (lists/outline) -> plain text, plus fields -> plain text.
// This version registers by direct assignment to WordToolkit.modules[...] to avoid any
// register() implementation mismatch.

const CHUNK_SIZE = 200;

(function () {
  // Ensure the exact registry shape taskpane_main.js expects
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  // Key MUST match: data-key="dynamicNumberingToText"
  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const status = (m) => setStatus(`Dynamic numbering → text\n${m}`);

    status("Starting…");

    try {
      await Word.run(async (context) => {
        const body = context.document.body;

        // A) Snapshot list/outline numbers
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

        // B) Convert fields to plain text (skip if API not available -> ApiNotFound)
        let fieldCount = 0;
        let usedUnlink = false;

        try {
          status(`Found numbered paragraphs: ${listItems.length}\nLoading fields…`);

          const fields = body.fields; // may throw ApiNotFound on some builds
          fields.load("items");
          await context.sync();

          const canUnlink =
            Office.context.requirements?.isSetSupported?.("WordApiDesktop", "1.4") === true;

          const fieldArray = fields.items.slice().reverse();
          fieldCount = fieldArray.length;
          usedUnlink = canUnlink;

          status(`Fields found: ${fieldCount}\nConverting fields…`);

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
            } catch {}

            done++;
            if (done % CHUNK_SIZE === 0) {
              status(`Converting fields: ${done}/${fieldCount}`);
              await context.sync();
            }
          }
          await context.sync();
        } catch (e) {
          status("Fields step skipped (ApiNotFound / not supported).\nContinuing…");
        }

        // C) Apply numbering-as-text and remove list formatting
        status(`Converting numbering: 0/${listItems.length}`);
        let done = 0;

        for (const it of listItems) {
          const p = paragraphs.items[it.index];

          p.insertText(it.listString + "\t", Word.InsertLocation.start);

          try { p.detachFromList(); } catch {}
          try { p.getRange().listFormat.removeNumbers(); } catch {}

          done++;
          if (done % CHUNK_SIZE === 0) {
            status(`Converting numbering: ${done}/${listItems.length}`);
            await context.sync();
          }
        }

        await context.sync();

        // Also write a report into the document so you can always see it
        const report =
          "REPORT: Dynamic numbering → text\n" +
          `Fields converted: ${fieldCount} (${usedUnlink ? "unlink" : "fallback/skip"})\n` +
          `Numbered paragraphs converted: ${listItems.length}`;

        body.insertParagraph(report, Word.InsertLocation.end);

        status(
          "Complete.\n" +
            `Fields converted: ${fieldCount} (${usedUnlink ? "unlink" : "fallback/skip"})\n` +
            `Numbered paragraphs converted: ${listItems.length}`
        );

        await context.sync();
      });
    } catch (e) {
      status("ERROR:\n" + String(e?.message || e));
      throw e;
    }
  };
})();
