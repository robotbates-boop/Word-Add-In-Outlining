/* global Word, Office */

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
      return `${m}m ${s % 60}s`;
    };

    const status = (msg) => setStatus(`Dynamic numbering → text\n${msg}`);

    status("Starting…");

    try {
      let fieldsFound = 0;
      let numberedParasFound = 0;
      let canUnlink = false;

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
        numberedParasFound = listItems.length;

        // B) Convert fields to plain text
        status(`Found numbered paragraphs: ${numberedParasFound}\nLoading fields…`);
        const fields = body.fields;
        fields.load("items");
        await context.sync();

        canUnlink = Office.context.requirements?.isSetSupported?.("WordApiDesktop", "1.4") === true;

        const fieldArray = fields.items.slice().reverse();
        fieldsFound = fieldArray.length;

        status(`Fields found: ${fieldsFound}\nConverting fields…`);
        let fieldsDone = 0;

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
            status(`Converting fields: ${fieldsDone}/${fieldsFound}\nElapsed: ${fmtMs(Date.now() - t0)}`);
            await context.sync();
          }
        }
        await context.sync();

        // C) Apply numbering-as-text and remove list formatting
        status(`Fields converted: ${fieldsFound}\nConverting numbering…`);
        let numberingDone = 0;

        for (const it of listItems) {
          const p = paragraphs.items[it.index];
          p.insertText(it.listString + "\t", Word.InsertLocation.start);

          try { p.detachFromList(); } catch {}
          try { p.getRange().listFormat.removeNumbers(); } catch {}

          numberingDone++;
          if (numberingDone % CHUNK_SIZE === 0) {
            status(`Converting numbering: ${numberingDone}/${numberedParasFound}\nElapsed: ${fmtMs(Date.now() - t0)}`);
            await context.sync();
          }
        }
        await context.sync();

        const elapsed = fmtMs(Date.now() - t0);
        const summary =
          "REPORT: Dynamic numbering → text\n" +
          `Fields converted: ${fieldsFound}\n` +
          `Numbered paragraphs converted: ${numberedParasFound}\n` +
          `Field unlink: ${canUnlink ? "used" : "fallback"}\n` +
          `Elapsed: ${elapsed}`;

        // Write into the document so it is always visible
        body.insertParagraph(summary, Word.InsertLocation.end);

        status(`Complete.\nFields: ${fieldsFound}\nNumbered paras: ${numberedParasFound}\nElapsed: ${elapsed}`);
        await context.sync();
      });
    } catch (e) {
      status("ERROR:\n" + String(e?.message || e));
      throw e;
    }
  });
})();
