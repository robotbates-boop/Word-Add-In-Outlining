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
    let detachSucceeded = 0;
    let removeNumbersSucceeded = 0;

    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const selection = context.document.getSelection();

        // Determine if we have a non-empty selection
        selection.load("text");
        await context.sync();

        const hasSelection = !!(selection.text && selection.text.trim().length > 0);
        const scopeRange = hasSelection ? selection : body.getRange();

        status(hasSelection ? "Using selection scope." : "No selection text detected; using whole document scope.");

        // ------------------------------------------------------------
        // 1) Collect paragraphs in scope by OVERLAP test (robust)
        // ------------------------------------------------------------
        const allParas = body.paragraphs;
        allParas.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );

        // For overlap testing we need each paragraph range + compare result
        const paraCompareResults = [];
        await context.sync();

        for (let i = 0; i < allParas.items.length; i++) {
          const pr = allParas.items[i].getRange();
          // compareLocationWith returns a ClientResult<string>
          const cmp = pr.compareLocationWith(scopeRange);
          paraCompareResults.push({ index: i, cmp });
        }
        await context.sync();

        // Keep paragraphs that are inside/overlapping the selection (or whole doc)
        const scopedParaIndexes = [];
        for (const r of paraCompareResults) {
          const v = String(r.cmp.value || "");
          // Accept any non-"Before"/non-"After" relationship (covers Overlap/Inside/Contains/etc.)
          if (v && v !== "Before" && v !== "After") scopedParaIndexes.push(r.index);
        }

        status(
          `Paragraphs total: ${allParas.items.length}\n` +
          `Paragraphs in scope: ${scopedParaIndexes.length}`
        );

        // Snapshot list markers for scoped paragraphs
        const scopedListParas = [];
        for (const idx of scopedParaIndexes) {
          const p = allParas.items[idx];
          const li = p.listItemOrNullObject;
          if (li && li.isNullObject === false) {
            const ls = li.listString ? String(li.listString) : "";
            if (ls) scopedListParas.push({ index: idx, listString: ls });
          }
        }

        // Bottom-up (by index in body paragraph collection)
        scopedListParas.sort((a, b) => b.index - a.index);

        status(
          `List/outline paragraphs in scope: ${scopedListParas.length}\n` +
          `Preparing to convert fields...`
        );

        // ------------------------------------------------------------
        // 2) Convert fields in scope (robust overlap filtering)
        // ------------------------------------------------------------
        try {
          const allFields = body.fields; // may be ApiNotFound in some hosts
          allFields.load("items");
          await context.sync();

          // Determine if unlink is available (desktop)
          const canUnlink =
            Office?.context?.requirements?.isSetSupported?.("WordApiDesktop", "1.4") === true;

          // Build overlap test for each field
          const fieldCompareResults = [];
          for (let i = 0; i < allFields.items.length; i++) {
            const fr = allFields.items[i].getRange();
            const cmp = fr.compareLocationWith(scopeRange);
            fieldCompareResults.push({ index: i, cmp });
          }
          await context.sync();

          const scopedFieldIndexes = [];
          for (const r of fieldCompareResults) {
            const v = String(r.cmp.value || "");
            if (v && v !== "Before" && v !== "After") scopedFieldIndexes.push(r.index);
          }

          // Process bottom-up (reverse order)
          scopedFieldIndexes.sort((a, b) => b - a);

          status(`Fields total: ${allFields.items.length}\nFields in scope: ${scopedFieldIndexes.length}\nConverting fields...`);

          let done = 0;
          for (const idx of scopedFieldIndexes) {
            const f = allFields.items[idx];
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
              fieldsConverted++;
            } catch {}

            done++;
            if (done % CHUNK_SIZE === 0) {
              status(`Converting fields: ${done}/${scopedFieldIndexes.length}`);
              await context.sync();
            }
          }
          await context.sync();
        } catch {
          fieldsSkipped = true;
          status("Fields step skipped (API not available).\nContinuing...");
        }

        // ------------------------------------------------------------
        // 3) Convert numbering in scope bottom-up
        // ------------------------------------------------------------
        status(`Converting numbering: 0/${scopedListParas.length}`);

        let doneNum = 0;
        for (const it of scopedListParas) {
          const p = allParas.items[it.index];

          // Insert list marker as text
          p.insertText(it.listString + "\t", Word.InsertLocation.start);
          numberedConverted++;

          // Try to remove list formatting (best-effort)
          try { p.detachFromList(); detachSucceeded++; } catch {}
          try { p.getRange().listFormat.removeNumbers(); removeNumbersSucceeded++; } catch {}

          doneNum++;
          if (doneNum % CHUNK_SIZE === 0) {
            status(`Converting numbering: ${doneNum}/${scopedListParas.length}`);
            await context.sync();
          }
        }

        await context.sync();

        status(
          "Complete.\n" +
            `Fields converted: ${fieldsConverted}${fieldsSkipped ? " (fields skipped)" : ""}\n` +
            `Numbered paragraphs converted: ${numberedConverted}\n` +
            `detachFromList succeeded: ${detachSucceeded}\n` +
            `removeNumbers succeeded: ${removeNumbersSucceeded}`
        );
      });
    } catch (e) {
      status("ERROR:\n" + String(e?.message || e));
      throw e;
    }
  };
})();
