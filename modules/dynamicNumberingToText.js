/* global Word, Office */

const CHUNK_SIZE = 200;

(function () {
  // Ensure toolkit exists
  window.WordToolkit = window.WordToolkit || {
    modules: {},
    register: function (key, fn) { this.modules[key] = fn; },
  };

  // MUST match taskpane.js key: "dynamicNumberingToText"
  window.WordToolkit.register("dynamicNumberingToText", async ({ setStatus }) => {
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

        // B) Convert fields to plain text
        status(`Found numbered paragraphs: ${listItems.length}\nLoading fields…`);
        const fields = body.fields;
        fields.load("items");
        await context.sync();

        const canUnlink = Office.context.requirements?.isSetSupported?.("WordApiDesktop", "1.4") === true;

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
          } catch {}

          done++;
          if (done % CHUNK_SIZE === 0) {
            status(`Converting fields: ${done}/${fieldArray.length}`);
            await context.sync();
          }
        }
        await context.sync();

        // C) Apply numbering-as-text and remove list formatting
        status(`Converting numbering: 0/${listItems.length}`);
        done = 0;

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
        status(
          "Complete.\n" +
            `Fields converted: ${fieldArray.length}\n` +
            `Numbered paragraphs converted: ${listItems.length}\n` +
            (canUnlink ? "Field unlink: used" : "Field unlink: fallback used")
        );
      });
    } catch (e) {
      setStatus("ERROR:\n" + String(e?.message || e));
      throw e;
    }
  });
})();
