/* global Word console */

import { analyze as apiAnalyze, type AnalyzeResponse } from "./api";

/** Highlight colors: yellow first, then 9 others. Cycles for each new selection. */
const HIGHLIGHT_COLORS: string[] = [
  "Yellow",
  "BrightGreen",
  "Turquoise",
  "Pink",
  "Blue",
  "Red",
  "DarkBlue",
  "Teal",
  "Green",
  "Violet",
];

let nextColorIndex = 0;

/** Get currently selected text from the document. */
export async function getSelectedText(): Promise<string> {
  let text = "";
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    text = range.text?.trim() ?? "";
  });
  return text;
}

/** Apply the next highlight color to the current selection. */
async function highlightCurrentSelection(): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    const color = HIGHLIGHT_COLORS[nextColorIndex % HIGHLIGHT_COLORS.length];
    nextColorIndex += 1;
    range.font.highlightColor = color;
    await context.sync();
  });
}

export interface AnalyzeSelectionResult {
  text: string;
  citation: AnalyzeResponse | null;
  error?: string;
}

/**
 * Gets the current selection, sends it to the backend POST /analyze,
 * then highlights the selection and returns the citation result.
 */
export async function analyzeSelection(options?: {
  documentId?: string;
  userId?: string;
}): Promise<AnalyzeSelectionResult> {
  const text = await getSelectedText();
  if (!text) {
    console.log("Analyze Selection: no text selected.");
    return { text: "", citation: null };
  }

  const { ok, status, data, error: apiError } = await apiAnalyze({
    text,
    document_id: options?.documentId,
    user_id: options?.userId,
  });

  if (!ok) {
    const errorMsg = apiError ?? `HTTP ${status}`;
    console.log("Analyze API error:", errorMsg);
    return { text, citation: null, error: errorMsg };
  }

  const citation = data ?? null;
  if (!citation || !citation.citation_text || citation.citation_text.trim() === "" || citation.confidence < 0.50) {
    const lowConfidenceMsg = "Unable to generate citation. Try refining your selection.";
    console.log(lowConfidenceMsg, citation);
    return { text, citation: null, error: lowConfidenceMsg };
  }

  console.log("Analyze Selection:", text, "->", citation);

  await highlightCurrentSelection();

  return { text, citation };
}

/**
 * Searches the document for the given text and selects the first occurrence,
 * so Word jumps back to that passage (e.g. for Re-select in the citation list).
 */
export async function reselectText(text: string): Promise<void> {
  if (!text.trim()) return;
  try {
    await Word.run(async (context) => {
      const searchResults = context.document.body.search(text.trim());
      const firstRange = searchResults.getFirstOrNullObject();
      firstRange.load("isNullObject");
      await context.sync();
      if (!firstRange.isNullObject) {
        firstRange.select();
        await context.sync();
      }
    });
  } catch (err) {
    console.log("reselectText error:", err);
  }
}

/**
 * Removes highlight and associated comments for the first matching occurrence of the given text.
 * Used when a citation is removed from the UI list.
 */
export async function removeCitation(text: string): Promise<void> {
  const needle = text.trim();
  if (!needle) return;

  try {
    await Word.run(async (context) => {
      const searchResults = context.document.body.search(needle);
      const firstRange = searchResults.getFirstOrNullObject();
      firstRange.load("isNullObject");
      await context.sync();
      if (firstRange.isNullObject) return;

      // Remove highlight. The runtime accepts `null` to clear highlight.
      (firstRange.font as any).highlightColor = null;

      // Delete comments if supported (WordApi 1.4+).
      try {
        const comments = firstRange.getComments();
        comments.load("items");
        await context.sync();
        comments.items.forEach((c) => c.delete());
      } catch (commentErr) {
        // Ignore if comments APIs aren't available or deletion fails.
        console.log("removeCitation comment cleanup error:", commentErr);
      }

      await context.sync();
    });
  } catch (err) {
    console.log("removeCitation error:", err);
  }
}

/**
 * Returns true if the given text exists in the document body (at least one match), false otherwise.
 * Used to detect when cited text has been deleted so the citation list can be synced.
 */
export async function checkCitationExists(text: string): Promise<boolean> {
  const needle = text.trim();
  if (!needle) return false;
  try {
    let found = false;
    await Word.run(async (context) => {
      const searchResults = context.document.body.search(needle);
      const firstRange = searchResults.getFirstOrNullObject();
      firstRange.load("isNullObject");
      await context.sync();
      found = !firstRange.isNullObject;
    });
    return found;
  } catch {
    return false;
  }
}
