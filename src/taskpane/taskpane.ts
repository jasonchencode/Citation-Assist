/* global Word console */

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

/**
 * Analyzes the current selection: retrieves its text (logged to console)
 * and applies the next highlight color in the cycle (yellow, then 9 others).
 */
export async function analyzeSelection(): Promise<void> {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();

      const text = range.text?.trim() ?? "";
      if (!text) {
        console.log("Analyze Selection: no text selected.");
        return;
      }

      console.log("Analyze Selection:", text);

      const color = HIGHLIGHT_COLORS[nextColorIndex % HIGHLIGHT_COLORS.length];
      nextColorIndex += 1;

      range.font.highlightColor = color;
      await context.sync();
    });
  } catch (error) {
    console.log("Error in analyzeSelection: " + error);
  }
}
