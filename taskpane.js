Office.onReady(() => {
  document.querySelectorAll("#controls button").forEach((btn) => {
    btn.addEventListener("click", async () => runCommand(btn.dataset.cmd));
  });
});

async function runCommand(cmd) {
  const status = document.getElementById("status");
  status.textContent = `Running: ${cmd}...`;

  try {
    switch (cmd) {
      case "boldAll":
        await setAllBold(true);
        break;
      case "italicAll":
        await setAllItalic(true);
        break;
      case "smallerAll":
        await changeAllFontSize(-1); // -1 pt
        break;
      case "largerAll":
        await changeAllFontSize(+1); // +1 pt
        break;
      default:
        throw new Error(`Unknown command: ${cmd}`);
    }
    status.textContent = "Done.";
  } catch (e) {
    console.error(e);
    status.textContent = `Error: ${e?.message || e}`;
  }
}

async function setAllBold(isBold) {
  await Word.run(async (context) => {
    context.document.body.font.bold = isBold;
    await context.sync();
  });
}

async function setAllItalic(isItalic) {
  await Word.run(async (context) => {
    context.document.body.font.italic = isItalic;
    await context.sync();
  });
}

async function changeAllFontSize(deltaPoints) {
  await Word.run(async (context) => {
    const font = context.document.body.font;
    font.load("size");
    await context.sync();

    const current = font.size;
    const next = Math.max(1, current + deltaPoints); // prevent 0/negative
    font.size = next;

    await context.sync();
  });
}
