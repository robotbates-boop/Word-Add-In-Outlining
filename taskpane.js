Office.onReady(() => {
  const btn = document.getElementById("boldAll");
  const status = document.getElementById("status");

  btn.onclick = async () => {
    try {
      await Word.run(async (context) => {
        context.document.body.font.bold = true;
        await context.sync();
      });
      status.textContent = "Done.";
    } catch (e) {
      console.error(e);
      status.textContent = "Error: " + (e?.message || e);
    }
  };
});
