import type { FindReplaceMatchSnapshot } from "../types/agent";
import { isWordApiAvailable } from "./context";
import { USER_SELECTION_BOOKMARK, WordOperationError } from "./operations";

export { USER_SELECTION_BOOKMARK };

export function bookmarkNameForEdit(editId: string): string {
  return `msword_aichat_${editId.replace(/[^a-zA-Z0-9]/g, "_")}`;
}

export function bookmarkNameForSearchMatch(editId: string, searchIndex: number): string {
  return bookmarkNameForEdit(editId) + "_m" + String(searchIndex);
}

export interface FindReplaceStageOptions {
  matchCase?: boolean;
  maxReplacements?: number;
  matchIndex?: number;
  skipTables?: boolean;
}

export interface FindReplaceStageResult {
  find: string;
  replace: string;
  totalMatches: number;
  skippedInTables: number;
  staged: FindReplaceMatchSnapshot[];
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

export async function upsertUserSelectionBookmark(): Promise<{ bookmark: string; text: string }> {
  assertWordAvailable();

  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const text = selection.text ?? "";
      selection.insertBookmark(USER_SELECTION_BOOKMARK);
      await context.sync();

      return { bookmark: USER_SELECTION_BOOKMARK, text };
    });
  } catch (error) {
    wrapWordError(error, "Failed to pin the current selection.");
  }
}

export async function isUserSelectionBookmarkPresent(): Promise<boolean> {
  assertWordAvailable();

  try {
    return await Word.run(async (context) => {
      const range = context.document.getBookmarkRangeOrNullObject(USER_SELECTION_BOOKMARK);
      range.load("isNullObject");
      await context.sync();
      return !range.isNullObject;
    });
  } catch (_error) {
    return false;
  }
}

export async function readUserSelectionBookmarkText(): Promise<string | null> {
  assertWordAvailable();

  try {
    return await Word.run(async (context) => {
      const range = context.document.getBookmarkRangeOrNullObject(USER_SELECTION_BOOKMARK);
      range.load(["isNullObject", "text"]);
      await context.sync();
      if (range.isNullObject) {
        return null;
      }
      return range.text ?? "";
    });
  } catch (_error) {
    return null;
  }
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

export async function deleteBookmarks(bookmarks: string[]): Promise<void> {
  for (var i = 0; i < bookmarks.length; i += 1) {
    await deleteBookmark(bookmarks[i]);
  }
}

export async function stageFindReplacements(
  editId: string,
  find: string,
  replace: string,
  options: FindReplaceStageOptions,
): Promise<FindReplaceStageResult> {
  assertWordAvailable();
  var trimmedFind = find.trim();
  if (!trimmedFind) {
    throw new WordOperationError("find text is required.");
  }

  var matchCase = Boolean(options.matchCase);
  var skipTables = options.skipTables !== false;
  var maxReplacements = Math.min(100, Math.max(1, Math.floor(options.maxReplacements ?? 50)));
  var matchIndex =
    options.matchIndex === undefined ? undefined : Math.max(0, Math.floor(options.matchIndex));

  try {
    return await Word.run(async function (context) {
      var results = context.document.body.search(trimmedFind, {
        matchCase: matchCase,
        matchWholeWord: false,
      });
      results.load("items");
      await context.sync();

      var totalMatches = results.items.length;
      var skippedInTables = 0;
      var staged: FindReplaceMatchSnapshot[] = [];

      if (totalMatches === 0) {
        return {
          find: trimmedFind,
          replace: replace,
          totalMatches: 0,
          skippedInTables: 0,
          staged: [],
        };
      }

      if (matchIndex !== undefined && matchIndex >= totalMatches) {
        var lastIndex = Math.max(0, totalMatches - 1);
        throw new WordOperationError(
          "match_index " +
            matchIndex +
            " is out of range. Document has " +
            totalMatches +
            " match(es); valid indices are 0" +
            (totalMatches > 1 ? " through " + lastIndex : "") +
            " (0-based — first match is 0, not 1).",
        );
      }

      for (var index = 0; index < totalMatches; index += 1) {
        results.items[index].load("text");
        results.items[index].parentTableOrNullObject.load("isNullObject");
      }
      await context.sync();

      for (var scanIndex = 0; scanIndex < totalMatches; scanIndex += 1) {
        if (matchIndex !== undefined && scanIndex !== matchIndex) {
          continue;
        }

        var range = results.items[scanIndex];
        var previousText = range.text ?? "";
        var inTable = !range.parentTableOrNullObject.isNullObject;

        if (inTable && skipTables) {
          skippedInTables += 1;
          if (matchIndex !== undefined) {
            throw new WordOperationError(
              "Match " +
                matchIndex +
                " is inside a table. Use list_tables and update_table for table cell text, or set skip_tables to false.",
            );
          }
          continue;
        }

        var bookmark = bookmarkNameForSearchMatch(editId, scanIndex);
        range.insertBookmark(bookmark);
        staged.push({
          searchIndex: scanIndex,
          bookmark: bookmark,
          previousText: previousText,
        });

        if (matchIndex !== undefined) {
          break;
        }

        if (staged.length >= maxReplacements) {
          break;
        }
      }

      await context.sync();

      return {
        find: trimmedFind,
        replace: replace,
        totalMatches: totalMatches,
        skippedInTables: skippedInTables,
        staged: staged,
      };
    });
  } catch (error) {
    wrapWordError(error, "Failed to stage find and replace.");
  }
}

export async function applyStagedFindReplacements(
  snapshots: FindReplaceMatchSnapshot[],
  replaceText: string,
): Promise<void> {
  assertWordAvailable();
  if (!snapshots.length) {
    return;
  }

  var sorted = snapshots.slice().sort(function (a, b) {
    return b.searchIndex - a.searchIndex;
  });

  for (var i = 0; i < sorted.length; i += 1) {
    await replaceBookmarkText(sorted[i].bookmark, replaceText);
  }
}

export async function restoreFindReplaceSnapshots(
  snapshots: FindReplaceMatchSnapshot[],
): Promise<void> {
  assertWordAvailable();
  var sorted = snapshots.slice().sort(function (a, b) {
    return a.searchIndex - b.searchIndex;
  });

  for (var i = 0; i < sorted.length; i += 1) {
    await replaceBookmarkText(sorted[i].bookmark, sorted[i].previousText);
  }
}