/* global Word, Office */

/**
 * dynamicNumberingToText.js
 * VERSION: v1.3.0 (stable per-paragraph anchors)
 *
 * Method:
 * - Freeze selection with a wrapper CC.
 * - Snapshot listString for numbered paragraphs.
 * - Create a CONTENT CONTROL over each target paragraph (stable anchor).
 * - Insert listString text at start of each anchored paragraph and remove list formatting (best effort).
 * - Delete each per-paragraph CC with keepContent=true.
 * - Delete wrapper CC with keepContent=true.
 */

const VERSION = "v1.3.0";
const WRAPPER_TAG = "WordToolkit_DNTT_WRAPPER";
const ITEM_TAG = "WordToolkit_DNTT_ITEM";

(function () {
  window.WordToolkit = window.WordToolkit || {};
  window.WordToolkit.modules = window.WordToolkit.modules || {};
  window.WordToolkit.versions = window.WordToolkit.versions || {};
  window.WordToolkit.versions["dynamicNumberingToText"] = `dynamicNumberingToText ${VERSION}`;

  window.WordToolkit.modules["dynamicNumberingToText"] = async ({ setStatus }) => {
    const runStamp = new Date().toISOString();
    const status = (m) =>
      setStatus(`Dynamic numbering → text\n${m}\n\n${VERSION}\nRun: ${runStamp}`);

    let detected = 0;
    let anchored = 0;
    let converted = 0;
    let detachAttempts = 0;
    let removeNumbersAttempts = 0;

    status("Starting…");

    try {
      await Word.run(async (context) => {
        const doc = context.document;
        const selection = doc.getSelection();

        // --- 0) Clean up any leftover item controls from previous runs (keep contents)
        const allCCs = doc.contentControls;
        allCCs.load("items, items/tag");
        await context.sync();

        const leftovers = allCCs.items.filter((cc) => cc.tag === ITEM_TAG || cc.tag === WRAPPER_TAG);
        if (leftovers.length) {
          status(`Removing ${leftovers.length} leftover control(s)…`);
          for (const cc of leftovers) {
            try { cc.delete(true); } catch {}
          }
          await context.sync();
        }

        // --- 1) Ensure selection exists
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          status("No selection detected. Select the numbered paragraphs first.");
          return;
        }

        // --- 2) Freeze selection for this run
        status("Freezing selection…");
        const wrapper = selection.insertContentControl();
        wrapper.tag = WRAPPER_TAG;
        wrapper.appearance = "Hidden";

        const scope = wrapper.getRange();

        // --- 3) Snapshot numbered paragraphs in scope
        status("Loading paragraphs…");
        const paras = scope.paragraphs;
        paras.load(
          "items," +
            "items/listItemOrNullObject," +
            "items/listItemOrNullObject/isNullObject," +
            "items/listItemOrNullObject/listString"
        );
        await context.sync();

        const targets = [];
        for (let i = 0; i < paras.items.length; i++) {
          const p = paras.items[i];
          const li = p.listItemOrNullObject;
          const ls =
            li && li.isNullObject === false && li.listString ? String(li.listString) : "";
          if (ls) targets.push({ idx: i, ls });
        }

        // Bottom-up to reduce interference if Word is touchy
        targets.sort((a, b) => b.idx - a.idx);

        detected = targets.length;
        status(`Numbered detected: ${detected}\nCreating anchors…`);

        // --- 4) Create a stable CONTENT CONTROL over each target paragraph
        // Important: this anchors the paragraph location so edits don't drift.
        const itemControls = [];
        for (const t of targets) {
          const p = paras.items[t.idx];
          const pr = p.getRange(); // full paragraph range
          const cc = pr.insertContentControl();
          cc.tag = ITEM_TAG;
          cc.title = t.ls;          // stash listString
          cc.appearance = "Hidden";
          itemControls.push(cc);
        }
        anchored = itemControls.length;
        await context.sync();

        status(`Anchors created: ${anchored}\nConverting…`);

        // --- 5) Convert each anchored paragraph
        // We do NOT use the old Paragraph objects anymore; we use the CC ranges (stable).
        for (let i = 0; i < itemControls.length; i++) {
          const cc = itemControls[i];
          const ls = cc.title || "";

          const r = cc.getRange();
          // Insert the marker as plain text at the start of the anchored paragraph
          r.insertText(ls + "\t", Word.InsertLocation.start);

          // Best-effort list removal: operate on first paragraph inside the control
          try {
            const p = r.paragraphs.getFirst();
            p.detachFromList();
            detachAttempts++;
          } catch {}

          try {
            const p = r.paragraphs.getFirst();
            p.getRange().listFormat.removeNumbers();
            removeNumbersAttempts++;
          } catch {}

          // Delete control but KEEP contents (critical)
          try { cc.delete(true); } catch {}

          // Sync occasionally (or every time if you prefer)
          if ((i + 1) % 10 === 0) {
            await context.sync();
            status(`Converted: ${i + 1}/${anchored}`);
          }
          converted = i + 1;
        }
        await context.sync();

        // --- 6) Remove wrapper but KEEP contents (critical)
        try { wrapper.delete(true); } catch {}
        await context.sync();

        status(
          "Complete.\n" +
            `Numbered detected: ${detected}\n` +
            `Anchored: ${anchored}\n` +
            `Converted: ${converted}\n` +
            `detachFromList attempts: ${detachAttempts}\n` +
            `removeNumbers attempts: ${removeNumbersAttempts}`
        );
      });
    } catch (e) {
      const dbg = e && e.debugInfo ? JSON.stringify(e.debugInfo, null, 2) : "";
      status("ERROR:\n" + String(e?.message || e) + (dbg ? "\n\nDEBUG:\n" + dbg : ""));
      throw e;
    }
  };
})();
