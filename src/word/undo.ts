import type { UndoSnapshot } from "../types/agent";
import { deleteBookmark, deleteBookmarks, replaceBookmarkText, restoreFindReplaceSnapshots } from "./ranges";
import {
  applyStyleAtBookmark,
  deleteTableRowsAtIndex,
  formatAtBookmark,
  readTableAtIndex,
  restoreTableAtIndex,
  WordOperationError,
} from "./operations";

export async function revertUndoSnapshot(snapshot: UndoSnapshot): Promise<void> {
  if (!snapshot.bookmark) {
    throw new WordOperationError("Cannot undo: range bookmark is missing.");
  }

  switch (snapshot.kind) {
    case "replace_text":
    case "delete_range":
    case "insert_text":
      await replaceBookmarkText(snapshot.bookmark, snapshot.previousText);
      break;
    case "apply_style":
      await replaceBookmarkText(snapshot.bookmark, snapshot.previousText);
      if (snapshot.previousStyle) {
        await applyStyleAtBookmark(snapshot.bookmark, snapshot.previousStyle);
      }
      break;
    case "format_range":
      await replaceBookmarkText(snapshot.bookmark, snapshot.previousText);
      if (snapshot.previousFormat) {
        await formatAtBookmark(snapshot.bookmark, snapshot.previousFormat);
      }
      break;
    case "insert_table":
      await replaceBookmarkText(snapshot.bookmark, snapshot.previousText);
      break;
    case "update_table":
      if (snapshot.tableIndex === undefined || !snapshot.previousTableValues) {
        throw new WordOperationError("Cannot undo table update: snapshot is incomplete.");
      }
      if (snapshot.insertedRowCount && snapshot.insertedAtRow !== undefined) {
        await deleteTableRowsAtIndex(
          snapshot.tableIndex,
          snapshot.insertedAtRow,
          snapshot.insertedRowCount,
        );
      } else {
        const current = await readTableAtIndex(snapshot.tableIndex);
        const previousRows = snapshot.previousTableValues.length;
        if (current.rows > previousRows) {
          await deleteTableRowsAtIndex(snapshot.tableIndex, previousRows, current.rows - previousRows);
        }
      }
      await restoreTableAtIndex(snapshot.tableIndex, snapshot.previousTableValues);
      break;
    case "find_and_replace":
    case "replace_at_match":
      if (!snapshot.matchSnapshots?.length) {
        throw new WordOperationError("Cannot undo find/replace: snapshot is incomplete.");
      }
      await restoreFindReplaceSnapshots(snapshot.matchSnapshots);
      await deleteBookmarks(
        snapshot.matchSnapshots.map(function (match) {
          return match.bookmark;
        }),
      );
      return;
    default: {
      const exhaustive: never = snapshot.kind;
      throw new WordOperationError(`Unsupported undo kind: ${exhaustive}`);
    }
  }

  await deleteBookmark(snapshot.bookmark);
}