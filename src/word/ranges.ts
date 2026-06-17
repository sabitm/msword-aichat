import { isWordApiAvailable } from "./context";
import { WordOperationError } from "./operations";

export function bookmarkNameForEdit(editId: string): string {
  return `msword_aichat_${editId.replace(/[^a-zA-Z0-9]/g, "_")}`;
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

export async function captureSelectionBookmark(
  editId: string,
  options: { allowEmpty?: boolean } = {},
): Promise<{ bookmark: string; text: string }> {
  assertWordAvailable();
  const bookmark = bookmarkNameForEdit(editId);

  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const text = selection.text ?? "";
      if (!text.trim() && !options.allowEmpty) {
        throw new WordOperationError("No text selected.");
      }

      selection.insertBookmark(bookmark);
      await context.sync();

      return { bookmark, text };
    });
  } catch (error) {
    wrapWordError(error, "Failed to capture selection range.");
  }
}

export async function captureEndBookmark(
  editId: string,
): Promise<{ bookmark: string }> {
  assertWordAvailable();
  const bookmark = bookmarkNameForEdit(editId);

  try {
    await Word.run(async (context) => {
      const endRange = context.document.body.getRange(Word.RangeLocation.end);
      endRange.insertBookmark(bookmark);
      await context.sync();
    });
    return { bookmark };
  } catch (error) {
    wrapWordError(error, "Failed to capture end-of-document position.");
  }
}

export async function readBookmarkText(bookmark: string): Promise<string> {
  assertWordAvailable();
  try {
    return await Word.run(async (context) => {
      const range = context.document.getBookmarkRange(bookmark);
      range.load("text");
      await context.sync();
      return range.text ?? "";
    });
  } catch (error) {
    wrapWordError(error, "Failed to read bookmark range.");
  }
}

export async function replaceBookmarkText(bookmark: string, text: string): Promise<void> {
  assertWordAvailable();
  try {
    await Word.run(async (context) => {
      const range = context.document.getBookmarkRange(bookmark);
      range.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    wrapWordError(error, "Failed to replace text at bookmark.");
  }
}

export async function insertAtBookmark(
  bookmark: string,
  text: string,
  location: "before" | "after" | "replace" = "replace",
): Promise<void> {
  assertWordAvailable();
  const insertLocation =
    location === "before"
      ? Word.InsertLocation.before
      : location === "after"
        ? Word.InsertLocation.after
        : Word.InsertLocation.replace;

  try {
    await Word.run(async (context) => {
      const range = context.document.getBookmarkRange(bookmark);
      range.insertText(text, insertLocation);
      await context.sync();
    });
  } catch (error) {
    wrapWordError(error, "Failed to insert text at bookmark.");
  }
}

export async function deleteBookmarkRange(bookmark: string): Promise<void> {
  assertWordAvailable();
  try {
    await Word.run(async (context) => {
      const range = context.document.getBookmarkRange(bookmark);
      range.delete();
      await context.sync();
    });
  } catch (error) {
    wrapWordError(error, "Failed to delete bookmark range.");
  }
}

export async function deleteBookmark(_bookmark: string): Promise<void> {
  // Word JS API does not expose bookmark deletion. Orphan bookmarks use the msword_aichat_ prefix.
}