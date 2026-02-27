/* global Word */

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};

  // Must match button data-key:
  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    setStatus("dynamicNumberingToText\nModule registered and running…");

    await Word.run(async (context) => {
      context.document.body.insertParagraph("OK: dynamicNumberingToText minimal module ran", Word.InsertLocation.end);
      await context.sync();
    });

    setStatus("dynamicNumberingToText\nComplete.");
  };
})();
