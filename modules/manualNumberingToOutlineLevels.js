/* global Word */

// Placeholder module. Will be implemented later.
(function () {
  window.WordToolkit = window.WordToolkit || { modules: {}, register: (k, f) => (window.WordToolkit.modules[k] = f) };

  window.WordToolkit.register("manualNumberingToOutlineLevels", async ({ setStatus }) => {
    setStatus("Not implemented yet (manual numbering -> outline levels).");

    await Word.run(async (context) => {
      context.document.body.insertParagraph(
        "manualNumberingToOutlineLevels: placeholder ran",
        Word.InsertLocation.end
      );
      await context.sync();
    });
  });
})();
