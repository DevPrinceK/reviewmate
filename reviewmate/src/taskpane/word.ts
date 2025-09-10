/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { generateReviewComments, loadApiSettings, saveApiSettings, type ReviewComment } from "./openai";

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

async function loadSettingsToUI() {
  const s = await loadApiSettings();
  const base = document.getElementById("api-base") as any;
  const model = document.getElementById("api-model") as any;
  if (base) base.value = s.baseUrl || "https://api.openai.com/v1";
  if (model) model.value = s.model || "gpt-4o-mini";
}

async function readSettingsFromUI() {
  const baseUrl = (document.getElementById("api-base") as any)?.value?.trim();
  const model = (document.getElementById("api-model") as any)?.value?.trim();
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
  lastComments = [];
  lastSummary = undefined;
  lastFocuses = getFocuses();
  lastCustom = getCustom();

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
    setStatus(`Inserted ${comments.length} comment(s) on selection.`);
  });
}

export async function reviewDocument() {
  setStatus("Reviewing document...");
  lastComments = [];
  lastSummary = undefined;
  lastFocuses = getFocuses();
  lastCustom = getCustom();

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
    setStatus(`Inserted ${comments.length} comment(s) across document.`);
  });
}

export async function reviewAgain() {
  if (!lastInput) {
    setStatus("Nothing to review again. Run a review first.");
    return;
  }
  setStatus("Reviewing again with same settings...");
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
    setStatus(`Inserted ${comments.length} additional comment(s).`);
  });
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
    if (sideload) sideload.style.display = "none";
    if (body) body.style.display = "block";
    loadSettingsToUI().catch(() => {});
  const saveBtn = document.getElementById("btn-save-settings");
  if (saveBtn) saveBtn.onclick = () => readSettingsFromUI().then(() => setStatus("Settings saved."));
  const selBtn = document.getElementById("btn-review-selection");
  if (selBtn) selBtn.onclick = reviewSelection;
  const docBtn = document.getElementById("btn-review-document");
  if (docBtn) docBtn.onclick = reviewDocument;
  const againBtn = document.getElementById("btn-review-again");
  if (againBtn) againBtn.onclick = reviewAgain;
  const sumBtn = document.getElementById("btn-generate-summary");
  if (sumBtn) sumBtn.onclick = insertSummaryReport;
    setStatus("Ready.");
  }
});
