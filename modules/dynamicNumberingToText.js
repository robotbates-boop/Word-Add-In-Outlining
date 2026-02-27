/* global Word, Office */

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const status = (m) => setStatus(`Dynamic numbering → text\n${m}`);

    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const selection = context.document.getSelection();

        // Force scope to whole paragraphs from first selected paragraph to last selected paragraph
        const selParas = selection.paragraphs;
        selParas.load(
          "items," +
            "items/style," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        const countSelParas = selParas.items.length;
        status(`Selection paragraphs detected: ${countSelParas}`);

        const detected = [];
        for (let i = 0; i < selParas.items.length; i++) {
          const p = selParas.items[i];
          const li = p.listItemOrNullObject;
          const ls = (li && li.isNullObject === false && li.listString) ? String(li.listString) : "";
          if (ls) detected.push({ i, ls, style: p.style || "" });
        }

        // Build report
        let report =
          "DIAGNOSTIC REPORT\n" +
          `Selection paragraphs: ${countSelParas}\n` +
          `Paragraphs with listString: ${detected.length}\n\n`;

        for (const d of detected) {
          report += `#${d.i}  listString="${d.ls}"  style="${d.style}"\n`;
        }

        // Put report at end of doc so you can see what it found
        body.insertParagraph(report, Word.InsertLocation.end);
        await context.sync();

        // If you want conversion to still run, keep going:
        // Convert only the paragraphs Office.js actually detected (bottom-up)
        detected.sort((a, b) => b.i - a.i);
        for (const d of detected) {
          const p = selParas.items[d.i];
          p.insertText(d.ls + "\t", Word.InsertLocation.start);
          try { p.detachFromList(); } catch {}
          try { p.getRange().listFormat.removeNumbers(); } catch {}
        }
        await context.sync();

        status("Complete. Diagnostic report appended to document end.");
      });
    } catch (e) {
      status("ERROR:\n" + String(e?.message || e));
      throw e;
    }
  };
})();
