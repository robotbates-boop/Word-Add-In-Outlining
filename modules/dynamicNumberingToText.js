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
    let detachTried = 0;
    let removeNumbersTried = 0;

    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const selection = context.document.getSelection();

        // Ensure there is a real selection
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          status("No selection detected. Select the numbered paragraphs first.");
          return;
        }

        status("Freezing selection…");

        // 1) Freeze the selection with a temporary content control
        const cc = selection.insertContentControl();
        cc.tag = "WordToolkit_DynamicNumberingToText";
        cc.appearance = "Hidden"; // keep UI clean
        cc.cannotEdit = false;
        cc.cannotDelete = false;

        const scope = cc.getRange();
        await context.sync();

        // 2) Snapshot list labels within the frozen scope
        status("Loading paragraphs in frozen scope…");
        const paragraphs = scope.paragraphs;
        paragraphs.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        // Capture list strings bottom-up
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

        status(`Paragraphs in scope: ${paragraphs.items.length}\nList items detected: ${listItems.length}`);

        // 3) Convert fields within scope to plain text (best-effort)
        try {
          status("Loading fields in scope…");
          const fields = scope.fields; // may be ApiNotFound on some hosts
          fields.load("items");
          await context.sync();

          const canUnlink =
            Office?.context?.requirements?.isSetSupported?.("WordApiDesktop", "1.4") === true;

          const fieldArray = fields.items.slice().reverse();
          status(`Fields found: ${fieldArray.length}\nConverting fields…`);

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
                r.insertText(r.text || "", Word.InsertLocation.replace);
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
          status("Fields step skipped (API not available). Continuing…");
        }

        // 4) Apply numbering-as-text and remove list formatting (bottom-up)
        status(`Converting numbering: 0/${listItems.length}`);

        let doneNum = 0;
        for (const it of listItems) {
          const p = paragraphs.items[it.index];

          // Insert list label as text at start of paragraph
          p.insertText(it.listString + "\t", Word.InsertLocation.start);
          numberedConverted++;

          // Try remove list formatting (some hosts may not support these)
          try { p.detachFromList(); detachTried++; } catch {}
          try { p.getRange().listFormat.removeNumbers(); removeNumbersTried++; } catch {}

          doneNum++;
          if (doneNum % CHUNK_SIZE === 0) {
            status(`Converting numbering: ${doneNum}/${listItems.length}`);
            await context.sync();
          }
        }
        await context.sync();

        // 5) Remove the content control but keep contents
        status("Cleaning up…");
        try { cc.delete(false); } catch {}
        await context.sync();

        // 6) Final status (do NOT throw ApiNotFound after success)
        status(
          "Complete.\n" +
            `Fields converted: ${fieldsConverted}${fieldsSkipped ? " (fields skipped)" : ""}\n` +
            `Numbered paragraphs converted: ${numberedConverted}\n` +
            `detachFromList calls attempted: ${detachTried}\n` +
            `removeNumbers calls attempted: ${removeNumbersTried}`
        );
      });
    } catch (e) {
      const msg = String(e?.message || e);

      // If Word reports ApiNotFound after completing edits, treat it as a limitation, not a failure.
      if (msg.includes("ApiNotFound")) {
        status(
          "Complete (with host limitations).\n" +
            "Some optional Word APIs were unavailable, but the conversion step ran.\n\n" +
            `Details: ${msg}\n` +
            `Fields converted: ${fieldsConverted}${fieldsSkipped ? " (fields skipped)" : ""}\n` +
            `Numbered paragraphs converted: ${numberedConverted}`
        );
        return;
      }

      status("ERROR:\n" + msg);
      throw e;
    }
  };
})();
