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

  console.log("Analyze Selection:", text, "->", data);

  await highlightCurrentSelection();

  return { text, citation: data ?? null };
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
