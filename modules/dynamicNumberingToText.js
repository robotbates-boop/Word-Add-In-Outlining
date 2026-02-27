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

        // Freeze scope: wrap selection in a temp content control
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          status("No selection detected. Select the numbered paragraphs first.");
          return;
        }

        const cc = selection.insertContentControl();
        cc.tag = "WordToolkit_DNTT";
        cc.appearance = "Hidden";

        const scope = cc.getRange();

        // Load paragraphs with style + list info
        const paras = scope.paragraphs;
        paras.load(
          "items," +
            "items/style," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        const total = paras.items.length;
        const convertible = [];
        const notConvertible = [];

        for (let i = 0; i < paras.items.length; i++) {
          const p = paras.items[i];
          const li = p.listItemOrNullObject;
          const ls = (li && li.isNullObject === false && li.listString) ? String(li.listString) : "";

          if (ls) {
            convertible.push({ i, ls, style: p.style || "" });
          } else {
            notConvertible.push({ i, style: p.style || "" });
          }
        }

        // Report what the API can actually see
        let report =
          "DIAGNOSTIC: What Office.js can see\n" +
          `Paragraphs in selection: ${total}\n` +
          `Convertible (has listString): ${convertible.length}\n` +
          `Not convertible (no listString): ${notConvertible.length}\n\n`;

        if (notConvertible.length) {
          report += "Not-convertible paragraph styles (first 30):\n";
          for (const x of notConvertible.slice(0, 30)) {
            report += `#${x.i}  style="${x.style}"\n`;
          }
          report += "\nIf these are Heading 1/2/3 etc., your numbering is style-driven and Office.js will not expose the number string.\n";
        }

        // Convert only what is convertible (bottom-up)
        convertible.sort((a, b) => b.i - a.i);

        for (const it of convertible) {
          const p = paras.items[it.i];
          p.insertText(it.ls + "\t", Word.InsertLocation.start);
          try { p.detachFromList(); } catch {}
          try { p.getRange().listFormat.removeNumbers(); } catch {}
        }

        await context.sync();

        // Remove content control, keep contents
        try { cc.delete(false); } catch {}
        await context.sync();

        // Put report at end of document so you can see it
        body.insertParagraph(report, Word.InsertLocation.end);
        await context.sync();

        status(
          "Complete.\n" +
          `Paragraphs in selection: ${total}\n` +
          `Converted (listString): ${convertible.length}\n` +
          `Unconverted (no listString): ${notConvertible.length}\n` +
          "A diagnostic report was appended to the end of the document."
        );
      });
    } catch (e) {
      status("ERROR:\n" + String(e?.message || e));
      throw e;
    }
  };
})();
