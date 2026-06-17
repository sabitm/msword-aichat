import { isWordApiAvailable } from "./context";

export class WordOperationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "WordOperationError";
  }
}

function assertWordAvailable(): void {
  if (!isWordApiAvailable()) {
    throw new WordOperationError("Word APIs are unavailable outside Word.");
  }
}

export async function readSelectionPlain(): Promise<string> {
  assertWordAvailable();
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();
    return selection.text ?? "";
  });
}

export async function readBodyTextChunk(
  start = 0,
  maxChars = 4000,
): Promise<{ text: string; totalLength: number; start: number; hasMore: boolean }> {
  assertWordAvailable();
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const fullText = body.text ?? "";
    const slice = fullText.slice(start, start + maxChars);
    const end = start + slice.length;

    return {
      text: slice,
      totalLength: fullText.length,
      start,
      hasMore: end < fullText.length,
    };
  });
}

export async function insertText(
  text: string,
  location: "selection" | "end" = "selection",
): Promise<void> {
  assertWordAvailable();
  await Word.run(async (context) => {
    if (location === "end") {
      const body = context.document.body;
      body.insertText(text, Word.InsertLocation.end);
    } else {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
    }
    await context.sync();
  });
}

export async function replaceSelectionText(text: string): Promise<void> {
  assertWordAvailable();
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(text, Word.InsertLocation.replace);
    await context.sync();
  });
}