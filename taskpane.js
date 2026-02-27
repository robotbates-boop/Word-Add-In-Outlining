/* global Word, Office */

const CHUNK_SIZE = 200;

(function () {
  // Register into the menu system
  window.WordToolkit =
    window.WordToolkit || { modules: {}, register: (k, f) => (window.WordToolkit.modules[k] = f) };

  window.WordToolkit.register("dynamicNumberingToText", async ({ setStatus }) => {
    setStatus("Running...");

    try {
      await Word.run(async (context) => {
        const body = context.document.body;

        // A) Snapshot list/outline numbers
        setStatus("Loading paragraphs...");
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

        // B) Convert fields to plain text
        setStatus("Loading fields...");
        const fields = body.fields;
        fields.load("items");
        await context.sync();

        const canUnlink = Office.context.requirements?.isSetSupported?.("WordApiDesktop", "1.4") === true;

        setStatus(`Converting fields (${fields.items.length})...`);

        const fieldArray = fields.items.slice().reverse();
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
            setStatus(`Converting fields: ${done}/${fieldArray.length}`);
            await context.sync();
          }
        }
        await context.sync();

        // C) Apply numbering-as-text and remove list formatting
        setStatus(`Converting numbering (${listItems.length})...`);
        done = 0;

        for (const it of listItems) {
          const p = paragraphs.items[it.index];

          p.insertText(it.listString + "\t", Word.InsertLocation.start);

          try { p.detachFromList(); } catch {}
          try { p.getRange().listFormat.removeNumbers(); } catch {}

          done++;
          if (done % CHUNK_SIZE === 0) {
            setStatus(`Converting numbering: ${done}/${listItems.length}`);
            await context.sync();
          }
        }

        await context.sync();
        setStatus(
          "Done.\n" +
            `Fields converted: ${fieldArray.length}\n` +
            `Numbered paragraphs converted: ${listItems.length}\n` +
            (canUnlink ? "Field unlink: used" : "Field unlink: fallback used")
        );
      });
    } catch (e) {
      setStatus("ERROR: " + String(e?.message || e));
      throw e;
    }
  });
})();
