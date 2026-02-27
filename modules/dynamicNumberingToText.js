/* global Word, Office */

/**
 * dynamicNumberingToText.js
 * VERSION: v1.0.0+2026-02-27
 */

const VERSION = "v1.0.0+2026-02-27";
const CHUNK_SIZE = 50;

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const runStamp = new Date().toISOString();
    const status = (m) => setStatus(`Dynamic numbering → text\n${m}\n\n${VERSION}\nRun: ${runStamp}`);

    let detected = 0;
    let applied = 0;
    let fieldsConverted = 0;
    let fieldsSkipped = false;

    status("Starting…");

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          status("No selection detected. Select the numbered paragraphs first.");
          return;
        }

        // Freeze selection with a wrapper content control.
        // IMPORTANT: we do NOT delete this control at the end (to avoid deleting the selection).
        const wrapper = selection.insertContentControl();
        wrapper.tag = "WordToolkit_DNTT_WRAPPER";
        wrapper.title = `DNTT ${VERSION} ${runStamp}`;
        wrapper.appearance = "Hidden";

        const scope = wrapper.getRange();
        await context.sync();

        // Load paragraphs in scope
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

        // Fields in scope -> plain text (best-effort)
        try {
          status(`Detected numbered paragraphs: ${detected}\nLoading fields…`);
          const fields = scope.fields; // may be ApiNotFound in some hosts
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

        // Convert numbering: insert at start of paragraph range
        status(`Converting numbering: 0/${detected}`);
        let doneNum = 0;

        for (const it of items) {
          const p = paras.items[it.idx];

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

        // IMPORTANT: no wrapper.delete(...) here.
        // Leaving the wrapper avoids the selection being deleted in your host.

        status(
          "Complete.\n" +
            `Fields converted: ${fieldsConverted}${fieldsSkipped ? " (fields skipped)" : ""}\n` +
            `Numbered paragraphs detected: ${detected}\n` +
            `Numbered paragraphs converted: ${applied}\n` +
            "Note: wrapper content control left in place (hidden) to avoid deleting selection."
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
