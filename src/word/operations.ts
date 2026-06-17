import type { DocumentStyleName } from "../types/agent";
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

function wrapWordError(error: unknown, fallback: string): never {
  if (error instanceof WordOperationError) throw error;
  const message = error instanceof Error ? error.message : fallback;
  if (/busy|locked|coauthor|edit/i.test(message)) {
    throw new WordOperationError(
      "The document is locked or being co-edited. Try again when the section is available.",
    );
  }
  throw new WordOperationError(message);
}

export interface SearchMatch {
  index: number;
  text: string;
  start: number;
  end: number;
}

export interface FormatOptions {
  bold?: boolean;
  italic?: boolean;
  font_size?: number;
}

export async function readSelectionPlain(): Promise<string> {
  assertWordAvailable();
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return selection.text ?? "";
    });
  } catch (error) {
    wrapWordError(error, "Failed to read selection.");
  }
}

export async function readSelectionStyle(): Promise<string> {
  assertWordAvailable();
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      paragraphs.load("items");
      await context.sync();
      if (paragraphs.items.length === 0) return "Unknown";
      paragraphs.items[0].load("styleBuiltIn");
      await context.sync();
      return String(paragraphs.items[0].styleBuiltIn);
    });
  } catch (error) {
    wrapWordError(error, "Failed to read selection style.");
  }
}

export async function readBodyTextChunk(
  start = 0,
  maxChars = 4000,
): Promise<{ text: string; totalLength: number; start: number; hasMore: boolean }> {
  assertWordAvailable();
  try {
    return await Word.run(async (context) => {
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
  } catch (error) {
    wrapWordError(error, "Failed to read document text.");
  }
}

export async function searchDocument(
  query: string,
  maxResults = 20,
): Promise<SearchMatch[]> {
  assertWordAvailable();
  const trimmed = query.trim();
  if (!trimmed) return [];

  try {
    return await Word.run(async (context) => {
      const results = context.document.body.search(trimmed, {
        matchCase: false,
        matchWholeWord: false,
      });
      results.load("items");
      await context.sync();

      const matches: SearchMatch[] = [];
      const limit = Math.min(results.items.length, maxResults);

      for (let index = 0; index < limit; index += 1) {
        const range = results.items[index];
        range.load(["text", "start", "end"]);
      }
      await context.sync();

      for (let index = 0; index < limit; index += 1) {
        const range = results.items[index];
        matches.push({
          index,
          text: range.text ?? "",
          start: range.start,
          end: range.end,
        });
      }

      return matches;
    });
  } catch (error) {
    wrapWordError(error, "Failed to search document.");
  }
}

export async function insertText(
  text: string,
  location: "selection" | "end" = "selection",
): Promise<void> {
  assertWordAvailable();
  try {
    await Word.run(async (context) => {
      if (location === "end") {
        context.document.body.insertText(text, Word.InsertLocation.end);
      } else {
        const selection = context.document.getSelection();
        selection.insertText(text, Word.InsertLocation.replace);
      }
      await context.sync();
    });
  } catch (error) {
    wrapWordError(error, "Failed to insert text.");
  }
}

export async function replaceSelectionText(text: string): Promise<void> {
  assertWordAvailable();
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    wrapWordError(error, "Failed to replace selection.");
  }
}

export async function applyStyleAtBookmark(
  bookmark: string,
  style: DocumentStyleName,
): Promise<void> {
  assertWordAvailable();
  try {
    await Word.run(async (context) => {
      const range = context.document.getBookmarkRange(bookmark);
      const paragraphs = range.paragraphs;
      paragraphs.load("items");
      await context.sync();

      for (const paragraph of paragraphs.items) {
        paragraph.styleBuiltIn = style as Word.BuiltInStyleName;
      }
      await context.sync();
    });
  } catch (error) {
    wrapWordError(error, "Failed to apply style.");
  }
}

export async function formatAtBookmark(
  bookmark: string,
  format: FormatOptions,
): Promise<void> {
  assertWordAvailable();
  try {
    await Word.run(async (context) => {
      const range = context.document.getBookmarkRange(bookmark);
      range.load("font");
      await context.sync();

      if (typeof format.bold === "boolean") {
        range.font.bold = format.bold;
      }
      if (typeof format.italic === "boolean") {
        range.font.italic = format.italic;
      }
      if (typeof format.font_size === "number") {
        range.font.size = format.font_size;
      }

      await context.sync();
    });
  } catch (error) {
    wrapWordError(error, "Failed to format range.");
  }
}

export async function insertTableAtBookmark(
  bookmark: string,
  rows: number,
  columns: number,
  cellValues?: string[][],
): Promise<void> {
  assertWordAvailable();
  try {
    await Word.run(async (context) => {
      const range = context.document.getBookmarkRange(bookmark);
      const table = range.insertTable(rows, columns, Word.InsertLocation.after);
      table.load("rows");
      await context.sync();

      if (cellValues?.length) {
        for (let rowIndex = 0; rowIndex < Math.min(rows, cellValues.length); rowIndex += 1) {
          const row = table.rows.items[rowIndex];
          row.load("cells");
          await context.sync();
          const rowValues = cellValues[rowIndex] ?? [];
          for (
            let colIndex = 0;
            colIndex < Math.min(columns, rowValues.length);
            colIndex += 1
          ) {
            const cell = row.cells.items[colIndex];
            cell.body.insertText(rowValues[colIndex] ?? "", Word.InsertLocation.replace);
          }
        }
        await context.sync();
      }
    });
  } catch (error) {
    wrapWordError(error, "Failed to insert table.");
  }
}