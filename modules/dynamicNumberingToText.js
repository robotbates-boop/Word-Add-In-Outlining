/* global Word, Office */

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const status = (m) => setStatus(`Dynamic numbering → text\n${m}`);

    let detected = 0;
    let applied = 0;
    let detachWorked = 0;
    let removeNumbersWorked = 0;

    const optionalApiNotes = new Set();

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          status("No selection detected. Select the numbered paragraphs first.");
          return;
        }

        // Freeze selection with a wrapper content control (so scope doesn't collapse)
        const wrapper = selection.insertContentControl();
        wrapper.tag = "WordToolkit_DNTT_WRAPPER";
        wrapper.appearance = "Hidden";

        const scope = wrapper.getRange();

        // Load paragraphs in scope + list strings
        const paras = scope.paragraphs;
        paras.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        // Snapshot list strings
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
        detected = items.length;

        status(`Detected numbered paragraphs: ${detected}\nAnchoring…`);

        // STEP 1: Create a unique anchor (content control) at the START of each target paragraph.
        // This prevents all insertions from collapsing onto the same final paragraph.
        const anchors = [];
        for (const it of items) {
          const p = paras.items[it.idx];

          // Paragraph start range
          let startRange;
          try {
            startRange = p.getRange(Word.RangeLocation.start);
          } catch {
            // Fallback if RangeLocation overload isn't available
            startRange = p.getRange();
          }

          const cc = startRange.insertContentControl();
          cc.tag = "WordToolkit_DNTT_ANCHOR";
          cc.title = it.ls; // store the listString on the control for later
          cc.appearance = "Hidden";
          anchors.push({ cc, ls: it.ls, idx: it.idx });
        }
        await context.sync();

        status(`Anchors created: ${anchors.length}\nApplying…`);

        // STEP 2: Write into each anchor and clean list formatting.
        for (let k = 0; k < anchors.length; k++) {
          const a = anchors[k];
          const p = paras.items[a.idx];

          try {
            // Insert number text at the anchor (guaranteed per-paragraph location)
            a.cc.insertText(a.ls + "\t", Word.InsertLocation.start);
            await context.sync();
            applied++;
          } catch {
            // If insertion fails, continue
          }

          // Best-effort list removal (may be ApiNotFound)
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

          // Remove the anchor control but keep its contents
          try {
            a.cc.delete(false);
            await context.sync();
          } catch {}

          status(`Applied: ${applied}/${detected}`);
        }

        // Remove wrapper control but keep contents
        try {
          wrapper.delete(false);
          await context.sync();
        } catch {}

        const notes = optionalApiNotes.size
          ? `\nOptional APIs unavailable: ${Array.from(optionalApiNotes).join(", ")}`
          : "";

        status(
          "Complete.\n" +
            `Detected numbered paragraphs: ${detected}\n` +
            `Numbered paragraphs converted: ${applied}\n` +
            `detachFromList succeeded: ${detachWorked}\n` +
            `removeNumbers succeeded: ${removeNumbersWorked}` +
            notes
        );
      });
    } catch (e) {
      status("ERROR:\n" + String(e?.message || e));
      throw e;
    }
  };
})();
