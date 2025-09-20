/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import {
  generateReviewComments,
  generateTrimmedText,
  generateParaphrasedText,
  loadApiSettings,
  saveApiSettings,
  type ReviewComment,
} from "./openai";

let lastInput: string | null = null;
let lastFocuses: string[] = [];
let lastCustom = "";
let lastComments: ReviewComment[] = [];
let lastSummary: string | undefined;

// Simple DOM helpers inline to avoid unused warnings

function getFocuses(): string[] {
  return Array.from(document.querySelectorAll("input.focus:checked")).map((i: any) => i.value);
}

function getCustom(): string {
  const t = document.getElementById("custom-instructions") as any;
  return (t?.value || "").trim();
}

function setStatus(msg: string) {
  const el = document.getElementById("status");
  if (el) el.textContent = msg;
}

function setBusy(isBusy: boolean) {
  const spinner = document.getElementById("spinner");
  if (spinner) {
    if (isBusy) {
      spinner.classList.add("active");
      spinner.setAttribute("aria-hidden", "false");
    } else {
      spinner.classList.remove("active");
      spinner.setAttribute("aria-hidden", "true");
    }
  }
  const buttons = Array.from(
    document.querySelectorAll(
      "#btn-review-selection,#btn-review-document,#btn-review-again,#btn-generate-summary,#btn-save-settings,#btn-toggle-key,#btn-trim-selection,#btn-paraphrase-selection"
    )
  ) as any[];
  buttons.forEach((b) => {
    if (isBusy) b.setAttribute("disabled", "true");
    else b.removeAttribute("disabled");
  });
}

async function loadSettingsToUI() {
  const s = await loadApiSettings();
  const base = document.getElementById("api-base") as any;
  const modelSel = document.getElementById("api-model") as any;
  const modelCustom = document.getElementById("api-model-custom") as any;
  if (base) base.value = s.baseUrl || "https://api.openai.com/v1";
  const savedModel = s.model || "gpt-4o-mini";
  if (modelSel) {
    const options = Array.from(modelSel.options || []).map((o: any) => o.value);
    if (options.indexOf(savedModel) >= 0) {
      modelSel.value = savedModel;
      if (modelCustom) modelCustom.classList.add("app-hidden");
    } else {
      modelSel.value = "custom";
      if (modelCustom) {
        modelCustom.value = savedModel;
        modelCustom.classList.remove("app-hidden");
      }
    }
  }
  // Do not auto-populate API key for safety; user can paste again.
}

async function readSettingsFromUI() {
  const baseUrl = (document.getElementById("api-base") as any)?.value?.trim();
  const selVal = (document.getElementById("api-model") as any)?.value?.trim();
  let model = selVal;
  if (selVal === "custom") {
    const custom = (document.getElementById("api-model-custom") as any)?.value?.trim();
    if (custom) model = custom;
  }
  const apiKey = (document.getElementById("api-key") as any)?.value?.trim();
  await saveApiSettings({ baseUrl, model, apiKey: apiKey || undefined });
}

async function insertCommentsForRange(context: Word.RequestContext, rangeLike: any, comments: ReviewComment[]) {
  for (const c of comments) {
    if (c.quote && c.quote.trim().length >= 4) {
      const results = rangeLike.search(c.quote.trim(), { matchCase: false, matchWholeWord: false });
      results.load("items");
      // eslint-disable-next-line office-addins/no-context-sync-in-loop
      await context.sync();
      if (results.items.length > 0) {
        results.items[0].insertComment(c.comment);
        continue;
      }
    }
    rangeLike.insertComment(c.comment);
  }
}

async function insertSummaryAtEnd(context: Word.RequestContext, summaryText: string, comments: ReviewComment[]) {
  const body = context.document.body;
  body.insertParagraph("", Word.InsertLocation.end);
  body.insertParagraph("AI Review Summary", Word.InsertLocation.end);
  if (summaryText) {
    body.insertParagraph(summaryText, Word.InsertLocation.end);
  }
  if (comments.length) {
    for (const c of comments) {
      const line = `- ${c.comment}${c.severity ? ` [${c.severity}]` : ""}${c.quote ? ` (Ref: "${c.quote.slice(0, 80)}")` : ""}`;
      body.insertParagraph(line, Word.InsertLocation.end);
    }
  }
  await context.sync();
}

export async function reviewSelection() {
  setStatus("Reviewing selection...");
  setBusy(true);
  lastComments = [];
  lastSummary = undefined;
  lastFocuses = getFocuses();
  lastCustom = getCustom();
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load(["text"]);
      await context.sync();

      const text = (range.text || "").trim();
      if (!text) {
        setStatus("No selection found.");
        return;
      }

      lastInput = text;

      const { comments, summary } = await generateReviewComments(text, lastFocuses, lastCustom);
      if (!comments.length) {
        setStatus("No comments generated.");
        return;
      }

      await insertCommentsForRange(context, range, comments);

      lastComments = comments;
      lastSummary = summary;
      const sumEl = document.getElementById("last-summary");
      if (sumEl) sumEl.textContent = summary || "";
      setStatus(`Inserted ${comments.length} comment(s) on selection.`);
    });
  } catch (err: any) {
    setStatus(`Error: ${err?.message || err}`);
  } finally {
    setBusy(false);
  }
}

export async function reviewDocument() {
  setStatus("Reviewing document...");
  setBusy(true);
  lastComments = [];
  lastSummary = undefined;
  lastFocuses = getFocuses();
  lastCustom = getCustom();
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();

      const fullText = (body.text || "").trim();
      if (!fullText) {
        setStatus("Document is empty.");
        return;
      }

      const MAX = 7000;
      const input = fullText.length > MAX ? fullText.slice(0, MAX) : fullText;
      lastInput = input;

      const { comments, summary } = await generateReviewComments(input, lastFocuses, lastCustom);
      if (!comments.length) {
        setStatus("No comments generated.");
        return;
      }

      await insertCommentsForRange(context, body, comments);

      lastComments = comments;
      lastSummary = summary;
      const sumEl = document.getElementById("last-summary");
      if (sumEl) sumEl.textContent = summary || "";
      setStatus(`Inserted ${comments.length} comment(s) across document.`);
    });
  } catch (err: any) {
    setStatus(`Error: ${err?.message || err}`);
  } finally {
    setBusy(false);
  }
}

export async function reviewAgain() {
  if (!lastInput) {
    setStatus("Nothing to review again. Run a review first.");
    return;
  }
  setStatus("Reviewing again with same settings...");
  setBusy(true);
  try {
    await Word.run(async (context) => {
      const target = context.document.body;
      await context.sync();

      const { comments, summary } = await generateReviewComments(lastInput!, lastFocuses, lastCustom);
      if (!comments.length) {
        setStatus("No new comments generated.");
        return;
      }

      await insertCommentsForRange(context, target, comments);

      lastComments = comments;
      lastSummary = summary;
      const sumEl = document.getElementById("last-summary");
      if (sumEl) sumEl.textContent = summary || "";
      setStatus(`Inserted ${comments.length} additional comment(s).`);
    });
  } catch (err: any) {
    setStatus(`Error: ${err?.message || err}`);
  } finally {
    setBusy(false);
  }
}

export async function trimSelection() {
  setStatus("Trimming selection...");
  setBusy(true);
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load(["text"]);
      await context.sync();
      const text = (range.text || "").trim();
      if (!text) {
        setStatus("No selection found.");
        return;
      }
      const slider = document.getElementById("trim-length") as any;
      const toneSel = document.getElementById("tone-style") as any;
      let maxChars = 0;
      if (slider) {
        const pct = parseInt(slider.value, 10) || 0;
        if (pct > 0 && pct <= 100) {
          maxChars = Math.max(Math.round((text.length * pct) / 100), 20);
        }
      }
      const style = (toneSel && toneSel.value) || "default";
      const trimmed = await generateTrimmedText(text, maxChars, style);
      if (!trimmed) {
        setStatus("No trimmed result returned.");
        return;
      }
      range.insertText(trimmed, Word.InsertLocation.replace);
      await context.sync();
      setStatus("Selection trimmed.");
    });
  } catch (err: any) {
    setStatus(`Error: ${err?.message || err}`);
  } finally {
    setBusy(false);
  }
}

export async function paraphraseSelection() {
  setStatus("Paraphrasing selection...");
  setBusy(true);
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load(["text"]);
      await context.sync();
      const text = (range.text || "").trim();
      if (!text) {
        setStatus("No selection found.");
        return;
      }
      const toneSel = document.getElementById("tone-style") as any;
      const style = (toneSel && toneSel.value) || "default";
      const variants = await generateParaphrasedText(text, 1, style);
      const first = variants[0];
      if (!first) {
        setStatus("No paraphrase returned.");
        return;
      }
      range.insertText(first, Word.InsertLocation.replace);
      await context.sync();
      setStatus("Paraphrased selection.");
    });
  } catch (err: any) {
    setStatus(`Error: ${err?.message || err}`);
  } finally {
    setBusy(false);
  }
}

export async function insertSummaryReport() {
  await Word.run(async (context) => {
    if (!lastComments.length && !lastSummary) {
      setStatus("No recent review to summarize. Run a review first.");
      return;
    }
    await insertSummaryAtEnd(context, lastSummary || "", lastComments);
    setStatus("Summary report inserted at end of document.");
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const sideload = document.getElementById("sideload-msg");
    const body = document.getElementById("app-body");
    const boot = document.getElementById("boot-overlay");
    if (sideload) sideload.style.display = "none";
    if (body) body.style.display = "block";
    loadSettingsToUI().catch(() => {});
    if (boot) boot.classList.remove("active");

    const saveBtn = document.getElementById("btn-save-settings");
    if (saveBtn) {
      saveBtn.onclick = async () => {
        setBusy(true);
        try {
          await readSettingsFromUI();
          setStatus("Settings saved.");
        } finally {
          setBusy(false);
        }
      };
    }

    const toggleKeyBtn = document.getElementById("btn-toggle-key");
    if (toggleKeyBtn) {
      toggleKeyBtn.onclick = () => {
        const input = document.getElementById("api-key") as any;
        if (!input) return;
        const showing = input.type === "text";
        input.type = showing ? "password" : "text";
        const label = toggleKeyBtn.querySelector(".ms-Button-label") as any;
        if (label) label.textContent = showing ? "Show" : "Hide";
      };
    }

    const preset = document.getElementById("preset-select") as any;
    if (preset) {
      preset.onchange = () => {
        const v = preset.value;
        const boxes = Array.from(document.querySelectorAll("input.focus")) as any[];
        const set = (name: string, val: boolean) => {
          const b = boxes.find((x) => x.value === name);
          if (b) b.checked = val;
        };
        if (v === "grammar-clarity") {
          boxes.forEach((b) => (b.checked = false));
          set("grammar", true);
          set("clarity", true);
        } else if (v === "structure-argument") {
          boxes.forEach((b) => (b.checked = false));
          set("structure", true);
          set("argument strength", true);
        } else if (v === "citations") {
          boxes.forEach((b) => (b.checked = false));
          set("citation check", true);
        } else if (v === "comprehensive") {
          boxes.forEach((b) => (b.checked = true));
        }
      };
    }

    const modelSel = document.getElementById("api-model") as any;
    if (modelSel) {
      modelSel.onchange = () => {
        const isCustom = modelSel.value === "custom";
        const modelCustom = document.getElementById("api-model-custom") as any;
        if (modelCustom) {
          if (isCustom) modelCustom.classList.remove("app-hidden");
          else modelCustom.classList.add("app-hidden");
        }
      };
    }

    const selBtn = document.getElementById("btn-review-selection");
    if (selBtn) selBtn.onclick = reviewSelection;
    const docBtn = document.getElementById("btn-review-document");
    if (docBtn) docBtn.onclick = reviewDocument;
    const againBtn = document.getElementById("btn-review-again");
    if (againBtn) againBtn.onclick = reviewAgain;
    const sumBtn = document.getElementById("btn-generate-summary");
    if (sumBtn) sumBtn.onclick = insertSummaryReport;
    const trimBtn = document.getElementById("btn-trim-selection");
    if (trimBtn) trimBtn.onclick = trimSelection;
    const paraBtn = document.getElementById("btn-paraphrase-selection");
    if (paraBtn) paraBtn.onclick = paraphraseSelection;
    const trimSlider = document.getElementById("trim-length") as any;
    const trimValue = document.getElementById("trim-length-value");
    if (trimSlider && trimValue) {
      const update = () => {
        const v = parseInt(trimSlider.value, 10) || 0;
        trimValue.textContent = v > 0 ? `${v}%` : "auto";
      };
      trimSlider.oninput = update;
      update();
    }

    // Accordion toggle for Connection section
    const acc = document.getElementById("connection-accordion");
    const accT = document.getElementById("connection-toggle");
    if (acc && accT) {
      accT.onclick = () => {
        const isCollapsed = acc.classList.contains("collapsed");
        if (isCollapsed) {
          acc.classList.remove("collapsed");
          accT.setAttribute("aria-expanded", "true");
        } else {
          acc.classList.add("collapsed");
          accT.setAttribute("aria-expanded", "false");
        }
      };
    }

    setStatus("Ready.");
  }
});
