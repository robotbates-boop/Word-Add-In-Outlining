/* global Word */

// Placeholder module. Will be implemented later.
(function () {
  window.WordToolkit = window.WordToolkit || { modules: {}, register: (k, f) => (window.WordToolkit.modules[k] = f) };

  window.WordToolkit.register("applyStyleTemplateToSelected", async ({ setStatus }) => {
    setStatus("Not implemented yet (apply style template to selected).");

    await Word.run(async (context) => {
      context.document.body.insertParagraph(
        "applyStyleTemplateToSelected: placeholder ran",
        Word.InsertLocation.end
      );
      await context.sync();
    });
  });
})();
