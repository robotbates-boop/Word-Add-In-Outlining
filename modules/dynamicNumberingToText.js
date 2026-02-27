/* global Word, Office */

/**
 * dynamicNumberingToText.js
 * VERSION: v1.0.1
 */

const VERSION = "v1.0.1";
const CHUNK_SIZE = 25; // small chunk for stability

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const runStamp = new Date().toISOString();
    const status = (m) =>
      setStatus(`Dynamic numbering → text\n${m}\n\n${VERSION}\nRun: ${runStamp}`);

    let detected = 0;
    let applied = 0;
    let fieldsConverted = 0;
    let fieldsSkipped = false;

    status("Starting…");

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();

        // Ensure a real selection exists
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          status("No selection detected. Select the numbered paragraphs first.");
          return;
        }

        // 1) Freeze selection for this run
        status("Freezing selection…");
        const wrapper = selection.insertContentControl();
        wrapper.tag = "WordToolkit_DNTT_WRAPPER";
        wrapper.title = `DNTT ${VERSION} ${runStamp}`;
        wrapper.appearance = "Hidden";

        // IMPORTANT: capture the wrapper range now so we can re-select after cleanup
        const wrapperRange = wrapper.getRange();
        await context.sync();

        // Use the wrapper range as stable scope
        const scope = wrapperRange;

        // 2) Load paragraphs in scope
        status("Loading paragraphs…");
        const paras = scope.paragraphs;
        paras.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        // Snapshot list strings bottom-up
        const items = [];
        for (let i = 0; i < paras.items.length; i++) {
          const p = paras.items[i];
          const li = p.listItemOrNullObject;
          const ls =
            li && li.isNullObject === false && li.listString ? String(li.listString) : "";
          if (ls) items.push({ idx: i, ls });
        }
        items.sort((a, b) => b.idx - a.idx);
        detected = items.length;

        // 3) Convert fields in scope (best-effort)
        try {
          status(`Detected numbered paragraphs: ${detected}\nLoading fields…`);
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
                f.unlink();
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
          status(`Detected numbered paragraphs: ${detected}\nFields skipped (API not available).`);
        }

        // 4) Convert numbering (bottom-up)
        status(`Converting numbering: 0/${detected}`);
        let doneNum = 0;

        for (const it of items) {
          const p = paras.items[it.idx];

          // Insert the displayed number string as text at paragraph start
          p.getRange().insertText(it.ls + "\t", Word.InsertLocation.start);
          applied++;

          // Best-effort list removal
          try { p.detachFromList(); } catch {}
          try { p.getRange().listFormat.removeNumbers(); } catch {}

          doneNum++;
          if (doneNum % CHUNK_SIZE === 0) {
            status(`Converting numbering: ${doneNum}/${detected}`);
            await context.sync();
          }
        }

        await context.sync();

        // 5) Restore selection to what we processed (before removing wrapper)
        try {
          wrapperRange.select();
          await context.sync();
        } catch {}

        // 6) Remove wrapper but KEEP contents (critical)
        // This avoids deleting the selected paragraphs AND avoids leaving wrappers behind.
        try {
          wrapper.delete(true);
          await context.sync();
        } catch {}

        status(
          "Complete.\n" +
            `Fields converted: ${fieldsConverted}${fieldsSkipped ? " (fields skipped)" : ""}\n` +
            `Numbered paragraphs detected: ${detected}\n` +
            `Numbered paragraphs converted: ${applied}`
        );
      });
    } catch (e) {
      const dbg = e && e.debugInfo ? JSON.stringify(e.debugInfo, null, 2) : "";
      status(
        "ERROR:\n" +
          String(e?.message || e) +
          (dbg ? "\n\nDEBUG INFO:\n" + dbg : "")
      );
      throw e;
    }
  };
})();
