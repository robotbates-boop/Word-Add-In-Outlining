/* global Word */

// Placeholder module. Will be implemented later.
(function () {
  window.WordToolkit = window.WordToolkit || { modules: {}, register: (k, f) => (window.WordToolkit.modules[k] = f) };

  window.WordToolkit.register("dynamicNumberingToText", async ({ setStatus }) => {
    setStatus("Not implemented yet (dynamic numbering to text).");

    await Word.run(async (context) => {
      context.document.body.insertParagraph(
        "dynamicNumberingToText: placeholder ran",
        Word.InsertLocation.end
      );
      await context.sync();
    });
  });
})();
