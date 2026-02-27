/* global Word, Office */

const CHUNK_SIZE = 200;

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

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

        // B) Convert fields to plain text (skip if not supported)
        let fieldCount = 0;
        let usedUnlink = false;
        try {
          status(`Found numbered paragraphs: ${listItems.length}\nLoading fields…`);
          const fields = body.fields; // may throw ApiNotFound
          fields.load("items");
          await context.sync();

          const canUnlink = Office.context.requirements?.isSetSupported?.("WordApiDesktop", "1.4") === true;
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
        } catch {
          status("Fields skipped (not supported). Continuing…");
        }

        // C) Apply numbering-as-text and (best-effort) remove list formatting
        status(`Converting numbering: 0/${listItems.length}`);
        let done = 0;
        let removedLists = 0;

        for (const it of listItems) {
          const p = paragraphs.items[it.index];

          // Insert the list label as text
          p.insertText(it.listString + "\t", Word.InsertLocation.start);

          // Best-effort list removal: each call might be ApiNotFound depending on host
          try { p.detachFromList(); removedLists++; } catch {}
          try { p.getRange().listFormat.removeNumbers(); } catch {}

          done++;
          if (done % CHUNK_SIZE === 0) {
            status(`Converting numbering: ${done}/${listItems.length}`);
            await context.sync();
          }
        }

        await context.sync();

        const report =
          "REPORT: Dynamic numbering → text\n" +
          `Fields converted: ${fieldCount} (${usedUnlink ? "unlink" : "fallback/skip"})\n` +
          `Numbered paragraphs processed: ${listItems.length}\n` +
          `detachFromList succeeded: ${removedLists}`;

        body.insertParagraph(report, Word.InsertLocation.end);

        status(
          "Complete.\n" +
            `Fields converted: ${fieldCount} (${usedUnlink ? "unlink" : "fallback/skip"})\n` +
            `Numbered paragraphs processed: ${listItems.length}\n` +
            `detachFromList succeeded: ${removedLists}`
        );

        await context.sync();
      });
    } catch (e) {
      // If anything still throws ApiNotFound, capture it
      status("ERROR:\n" + String(e?.message || e));
      throw e;
    }
  };
})();
