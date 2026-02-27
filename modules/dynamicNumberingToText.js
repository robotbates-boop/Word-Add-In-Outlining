/* global Word, Office */

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const status = (m) => setStatus(`Dynamic numbering → text\n${m}`);

    let numberedDetected = 0;
    let numberedApplied = 0;
    let detachWorked = 0;
    let removeNumbersWorked = 0;

    const optionalApiNotes = new Set();

    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const selection = context.document.getSelection();

        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          status("No selection detected. Select the numbered paragraphs first.");
          return;
        }

        // Freeze the selection so edits don't collapse it
        const cc = selection.insertContentControl();
        cc.tag = "WordToolkit_DNTT";
        cc.appearance = "Hidden";
        const scope = cc.getRange();

        // Load paragraphs + list info
        const paras = scope.paragraphs;
        paras.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        // Snapshot list strings now (before edits)
        const items = [];
        for (let i = 0; i < paras.items.length; i++) {
          const p = paras.items[i];
          const li = p.listItemOrNullObject;
          const ls =
            li && li.isNullObject === false && li.listString ? String(li.listString) : "";
          if (ls) items.push({ idx: i, ls });
        }

        // Bottom-up
        items.sort((a, b) => b.idx - a.idx);
        numberedDetected = items.length;

        status(`Detected numbered paragraphs: ${numberedDetected}\nApplying...`);

        // IMPORTANT CHANGE:
        // Apply one paragraph at a time, syncing each time, so we don't lose earlier insertions.
        for (let k = 0; k < items.length; k++) {
          const it = items[k];
          const p = paras.items[it.idx];

          try {
            // Insert marker as text
            p.insertText(it.ls + "\t", Word.InsertLocation.start);
            await context.sync();
            numberedApplied++;
          } catch (e) {
            // If insertText fails, skip this paragraph
            continue;
          }

          // Optional clean-up APIs (may be missing -> ApiNotFound)
          try {
            p.detachFromList();
            await context.sync();
            detachWorked++;
          } catch (e) {
            if (String(e?.message || e).includes("ApiNotFound")) optionalApiNotes.add("detachFromList");
          }

          try {
            p.getRange().listFormat.removeNumbers();
            await context.sync();
            removeNumbersWorked++;
          } catch (e) {
            if (String(e?.message || e).includes("ApiNotFound")) optionalApiNotes.add("removeNumbers");
          }

          status(`Applied: ${numberedApplied}/${numberedDetected}`);
        }

        // Remove content control, keep contents
        try {
          cc.delete(false);
          await context.sync();
        } catch {}

        const notes = optionalApiNotes.size
          ? `\nOptional APIs unavailable: ${Array.from(optionalApiNotes).join(", ")}`
          : "";

        status(
          "Complete.\n" +
            `Detected numbered paragraphs: ${numberedDetected}\n` +
            `Numbered paragraphs converted: ${numberedApplied}\n` +
            `detachFromList succeeded: ${detachWorked}\n` +
            `removeNumbers succeeded: ${removeNumbersWorked}` +
            notes
        );
      });
    } catch (e) {
      // Do not mask real errors; but keep message readable
      status("ERROR:\n" + String(e?.message || e));
      throw e;
    }
  };
})();
