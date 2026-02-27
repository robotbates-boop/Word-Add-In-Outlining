/* global Word */

// Placeholder module. Will be implemented later.
(function () {
  window.WordToolkit = window.WordToolkit || { modules: {}, register: (k, f) => (window.WordToolkit.modules[k] = f) };

  window.WordToolkit.register("automaticCrossReferencingToSelected", async ({ setStatus }) => {
    setStatus("Not implemented yet (automatic cross referencing to selected).");

    await Word.run(async (context) => {
      context.document.body.insertParagraph(
        "automaticCrossReferencingToSelected: placeholder ran",
        Word.InsertLocation.end
      );
      await context.sync();
    });
  });
})();
