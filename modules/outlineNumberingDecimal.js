/* global Word */

// Placeholder module. Will be implemented later.
(function () {
  window.WordToolkit = window.WordToolkit || { modules: {}, register: (k, f) => (window.WordToolkit.modules[k] = f) };

  window.WordToolkit.register("outlineNumberingDecimal", async ({ setStatus }) => {
    setStatus("Not implemented yet (outline numbering decimal).");

    await Word.run(async (context) => {
      context.document.body.insertParagraph(
        "outlineNumberingDecimal: placeholder ran",
        Word.InsertLocation.end
      );
      await context.sync();
    });
  });
})();
