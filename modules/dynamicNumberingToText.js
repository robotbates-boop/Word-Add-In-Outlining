/* global Word, Office */

// Convert outline/list numbering to plain text + convert fields to plain text.
// Safe list detection via listItemOrNullObject (avoids ItemNotFound).

const CHUNK_SIZE = 200;

(function () {
  window.WordToolkit =
    window.WordToolkit || { modules: {}, register: (k, f) => (window.WordToolkit.modules[k] = f) };

  window.WordToolkit.register("dynamicNumberingToText", async ({ setStatus }) => {
    const t0 = Date.now();

    const fmtMs = (ms) => {
      const s = Math.round(ms / 1000);
      if (s < 60) return `${s}s`;
      const m = Math.floor(s / 60);
      const r = s % 60;
      return `${m}m ${r}s`;
    };

    const status = (msg) => setStatus(msg);

    status("Dynamic numbering → text\nStarting…");

    try {
      await Word.run(async (context) => {
        const body = context.document.body;

        // A) Snapshot list/outline numbers
        status("Dynamic numbering → text\nLoading paragraphs…");
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

        status(
          "Dynamic numbering → text\n" +
            `Found numbered paragraphs: ${listItems.length}\n` +
            "Loading fields…"
        );

        // B) Convert fields to plain text
        const fields = body.fields;
        fields.load("items");
        await context.sync();

        const canUnlink =
          Office.context.requirements?.isSetSupported?.("WordApiDesktop", "1.4") === true;

        const fieldArray = fields.items.slice().reverse();
        let fieldsDone = 0;

        status(
          "Dynamic numbering → text\n" +
            `Fields found: ${fieldArray.length}\n` +
            `Field unlink: ${canUnlink ? "used" : "fallback"}\n` +
            "Converting fields…"
        );

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

          fieldsDone++;
          if (fieldsDone % CHUNK_SIZE === 0) {
            status(
              "Dynamic numbering → text\n" +
                `Converting fields: ${fieldsDone}/${fieldArray.length}\n` +
                `Elapsed: ${fmtMs(Date.now() - t0)}`
            );
            await context.sync();
          }
        }
        await context.sync();

        // C) Apply numbering-as-text and remove list formatting
        let numberingDone = 0;
        status(
          "Dynamic numbering → text\n" +
            `Fields converted: ${fieldArray.length}\n` +
            `Converting numbering: 0/${listItems.length}\n` +
            `Elapsed: ${fmtMs(Date.now() - t0)}`
        );

        for (const it of listItems) {
          const p = paragraphs.items[it.index];

          // Insert list label as text at paragraph start
          p.insertText(it.listString + "\t", Word.InsertLocation.start);

          // Remove list formatting
          try { p.detachFromList(); } catch {}
          try { p.getRange().listFormat.removeNumbers(); } catch {}

          numberingDone++;
          if (numberingDone % CHUNK_SIZE === 0) {
            status(
              "Dynamic numbering → text\n" +
                `Fields converted: ${fieldArray.length}\n` +
                `Converting numbering: ${numberingDone}/${listItems.length}\n` +
                `Elapsed: ${fmtMs(Date.now() - t0)}`
            );
            await context.sync();
          }
        }

        await context.sync();

        status(
          "Dynamic numbering → text\n" +
            "Complete.\n" +
            `Fields converted: ${fieldArray.length}\n` +
            `Numbered paragraphs converted: ${listItems.length}\n` +
            `Field unlink: ${canUnlink ? "used" : "fallback"}\n` +
            `Elapsed: ${fmtMs(Date.now() - t0)}`
        );
      });
    } catch (e) {
      status("Dynamic numbering → text\nERROR:\n" + String(e?.message || e));
      throw e;
    }
  });
})();
