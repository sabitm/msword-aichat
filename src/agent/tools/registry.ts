import type {
  DocumentStyleName,
  PendingEdit,
  ToolDefinition,
  ToolExecutionResult,
  UndoSnapshot,
} from "../../types/agent";
import { DOCUMENT_STYLES } from "../../types/agent";
import type { FormatOptions } from "../../word/operations";
import {
  applyStyleAtBookmark,
  formatAtBookmark,
  insertCommentOnSelection,
  insertTableAtBookmark,
  readBodyTextChunk,
  readSelectionPlain,
  readSelectionStyle,
  searchDocument,
  WordOperationError,
} from "../../word/operations";
import {
  captureEndBookmark,
  captureSelectionBookmark,
  deleteBookmark,
  deleteBookmarkRange,
  insertAtBookmark,
  replaceBookmarkText,
} from "../../word/ranges";
import { revertUndoSnapshot } from "../../word/undo";

function createId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function failure(toolName: string, error: string): ToolExecutionResult {
  return {
    success: false,
    output: { tool: toolName, error },
    error,
  };
}

function parseArgs<T extends Record<string, unknown>>(raw: string): T {
  try {
    return JSON.parse(raw) as T;
  } catch {
    return {} as T;
  }
}

function buildUndo(
  kind: UndoSnapshot["kind"],
  bookmark: string,
  previousText: string,
): UndoSnapshot {
  return { kind, bookmark, previousText };
}

async function applyMutationNow(
  edit: PendingEdit,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult | null> {
  if (!autoApplyEdits) return null;
  await applyPendingEdit(edit);
  return {
    success: true,
    output: { applied: true, tool: edit.toolName },
  };
}

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "get_selection",
    description: "Read the currently selected text in the Word document.",
    parameters: { type: "object", properties: {}, additionalProperties: false },
  },
  {
    name: "get_document_text",
    description:
      "Read a chunk of the document body text. Use start and max_chars to page through long documents.",
    parameters: {
      type: "object",
      properties: {
        start: { type: "number", description: "Character offset. Defaults to 0." },
        max_chars: { type: "number", description: "Max characters. Defaults to 4000." },
      },
      additionalProperties: false,
    },
  },
  {
    name: "search_document",
    description: "Search the document body for matching text and return up to 20 matches.",
    parameters: {
      type: "object",
      properties: {
        query: { type: "string", description: "Text to search for." },
        max_results: { type: "number", description: "Maximum matches to return. Defaults to 20." },
      },
      required: ["query"],
      additionalProperties: false,
    },
  },
  {
    name: "insert_text",
    description: "Insert text at the current selection or at the end of the document.",
    parameters: {
      type: "object",
      properties: {
        text: { type: "string", description: "Text to insert." },
        location: {
          type: "string",
          enum: ["selection", "end"],
          description: "Where to insert. Defaults to selection.",
        },
      },
      required: ["text"],
      additionalProperties: false,
    },
  },
  {
    name: "replace_text",
    description:
      "Replace the currently selected text with new text. Captures the range so Apply works even if selection changes.",
    parameters: {
      type: "object",
      properties: {
        new_text: { type: "string", description: "Replacement text for the current selection." },
      },
      required: ["new_text"],
      additionalProperties: false,
    },
  },
  {
    name: "delete_range",
    description: "Delete the currently selected text. Requires an active selection.",
    parameters: { type: "object", properties: {}, additionalProperties: false },
  },
  {
    name: "apply_style",
    description: "Apply a built-in paragraph style to the current selection.",
    parameters: {
      type: "object",
      properties: {
        style: {
          type: "string",
          enum: [...DOCUMENT_STYLES],
          description: "Built-in style name.",
        },
      },
      required: ["style"],
      additionalProperties: false,
    },
  },
  {
    name: "format_range",
    description: "Apply bold, italic, or font size to the current selection.",
    parameters: {
      type: "object",
      properties: {
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        font_size: { type: "number", description: "Font size in points." },
      },
      additionalProperties: false,
    },
  },
  {
    name: "insert_comment",
    description:
      "Add a Word review comment on the current selection. Does not change body text. Requires an active selection.",
    parameters: {
      type: "object",
      properties: {
        comment: { type: "string", description: "Comment text to attach to the selection." },
      },
      required: ["comment"],
      additionalProperties: false,
    },
  },
  {
    name: "insert_table",
    description: "Insert a table after the current selection (or document end if empty).",
    parameters: {
      type: "object",
      properties: {
        rows: { type: "number", description: "Number of rows (1-20)." },
        columns: { type: "number", description: "Number of columns (1-10)." },
        cells: {
          type: "array",
          description: "Optional 2D array of cell text values.",
          items: {
            type: "array",
            items: { type: "string" },
          },
        },
      },
      required: ["rows", "columns"],
      additionalProperties: false,
    },
  },
];

export async function executeTool(
  name: string,
  argsJson: string,
  options: { autoApplyEdits: boolean },
): Promise<ToolExecutionResult> {
  try {
    switch (name) {
      case "get_selection":
        return executeGetSelection();
      case "get_document_text":
        return executeGetDocumentText(argsJson);
      case "search_document":
        return executeSearchDocument(argsJson);
      case "insert_text":
        return executeInsertText(argsJson, options.autoApplyEdits);
      case "replace_text":
        return executeReplaceText(argsJson, options.autoApplyEdits);
      case "delete_range":
        return executeDeleteRange(options.autoApplyEdits);
      case "apply_style":
        return executeApplyStyle(argsJson, options.autoApplyEdits);
      case "format_range":
        return executeFormatRange(argsJson, options.autoApplyEdits);
      case "insert_table":
        return executeInsertTable(argsJson, options.autoApplyEdits);
      case "insert_comment":
        return executeInsertComment(argsJson);
      default:
        return failure(name, `Unknown tool: ${name}`);
    }
  } catch (error) {
    const message =
      error instanceof WordOperationError || error instanceof Error
        ? error.message
        : "Tool execution failed";
    return failure(name, message);
  }
}

async function executeGetSelection(): Promise<ToolExecutionResult> {
  const text = await readSelectionPlain();
  return {
    success: true,
    output: { text, empty: !text.trim(), length: text.length },
  };
}

async function executeGetDocumentText(argsJson: string): Promise<ToolExecutionResult> {
  const args = parseArgs<{ start?: number; max_chars?: number }>(argsJson);
  const start = Math.max(0, Math.floor(args.start ?? 0));
  const maxChars = Math.min(12_000, Math.max(1, Math.floor(args.max_chars ?? 4000)));
  const chunk = await readBodyTextChunk(start, maxChars);
  return {
    success: true,
    output: {
      text: chunk.text,
      start: chunk.start,
      returnedLength: chunk.text.length,
      totalLength: chunk.totalLength,
      hasMore: chunk.hasMore,
    },
  };
}

async function executeSearchDocument(argsJson: string): Promise<ToolExecutionResult> {
  const args = parseArgs<{ query?: string; max_results?: number }>(argsJson);
  const query = args.query?.trim() ?? "";
  if (!query) return failure("search_document", "query is required");

  const maxResults = Math.min(20, Math.max(1, Math.floor(args.max_results ?? 20)));
  const matches = await searchDocument(query, maxResults);
  return {
    success: true,
    output: {
      query,
      matchCount: matches.length,
      matches,
    },
  };
}

async function executeInsertText(
  argsJson: string,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult> {
  const args = parseArgs<{ text?: string; location?: "selection" | "end" }>(argsJson);
  const text = args.text?.trim() ?? "";
  const location = args.location ?? "selection";
  if (!text) return failure("insert_text", "text is required");

  const editId = createId();
  const bookmark =
    location === "end"
      ? (await captureEndBookmark(editId)).bookmark
      : (await captureSelectionBookmark(editId, { allowEmpty: true })).bookmark;

  const before =
    location === "end" ? "(end of document)" : (await readSelectionPlain()) || "(cursor)";

  const pendingEdit: PendingEdit = {
    id: editId,
    toolName: "insert_text",
    description: `Insert text at ${location}`,
    before,
    after: text,
    status: "pending",
    bookmark,
    undo: buildUndo("insert_text", bookmark, before),
    payload: { location },
  };

  const applied = await applyMutationNow(pendingEdit, autoApplyEdits);
  if (applied) return applied;

  return {
    success: true,
    output: {
      pendingApproval: true,
      location,
      preview: text.slice(0, 200),
      message: "Insert staged for user approval.",
    },
    pendingEdit,
  };
}

async function executeReplaceText(
  argsJson: string,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult> {
  const args = parseArgs<{ new_text?: string }>(argsJson);
  const newText = args.new_text ?? "";
  if (!newText.trim()) return failure("replace_text", "new_text is required");

  const editId = createId();
  const { bookmark, text: before } = await captureSelectionBookmark(editId);

  const pendingEdit: PendingEdit = {
    id: editId,
    toolName: "replace_text",
    description: "Replace selected text",
    before,
    after: newText,
    status: "pending",
    bookmark,
    undo: buildUndo("replace_text", bookmark, before),
  };

  const applied = await applyMutationNow(pendingEdit, autoApplyEdits);
  if (applied) return applied;

  return {
    success: true,
    output: {
      pendingApproval: true,
      beforeLength: before.length,
      afterLength: newText.length,
      message: "Replacement staged for user approval. Range captured via bookmark.",
    },
    pendingEdit,
  };
}

async function executeDeleteRange(autoApplyEdits: boolean): Promise<ToolExecutionResult> {
  const editId = createId();
  const { bookmark, text: before } = await captureSelectionBookmark(editId);

  const pendingEdit: PendingEdit = {
    id: editId,
    toolName: "delete_range",
    description: "Delete selected text",
    before,
    after: "(deleted)",
    status: "pending",
    bookmark,
    undo: buildUndo("delete_range", bookmark, before),
  };

  const applied = await applyMutationNow(pendingEdit, autoApplyEdits);
  if (applied) return applied;

  return {
    success: true,
    output: {
      pendingApproval: true,
      message: "Deletion staged for user approval.",
    },
    pendingEdit,
  };
}

async function executeApplyStyle(
  argsJson: string,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult> {
  const args = parseArgs<{ style?: DocumentStyleName }>(argsJson);
  const style = args.style;
  if (!style || !DOCUMENT_STYLES.includes(style)) {
    return failure("apply_style", `style must be one of: ${DOCUMENT_STYLES.join(", ")}`);
  }

  const editId = createId();
  const { bookmark, text } = await captureSelectionBookmark(editId);
  const beforeStyle = await readSelectionStyle();

  const pendingEdit: PendingEdit = {
    id: editId,
    toolName: "apply_style",
    description: `Apply style ${style}`,
    before: `${beforeStyle} — ${text.slice(0, 120)}`,
    after: `${style} — ${text.slice(0, 120)}`,
    status: "pending",
    bookmark,
    undo: {
      kind: "apply_style",
      bookmark,
      previousText: text,
      previousStyle: beforeStyle as DocumentStyleName,
    },
    payload: { style },
  };

  const applied = await applyMutationNow(pendingEdit, autoApplyEdits);
  if (applied) return applied;

  return {
    success: true,
    output: { pendingApproval: true, style, message: "Style change staged for approval." },
    pendingEdit,
  };
}

async function executeFormatRange(
  argsJson: string,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult> {
  const args = parseArgs<Record<string, unknown>>(argsJson) as FormatOptions;
  if (
    typeof args.bold !== "boolean" &&
    typeof args.italic !== "boolean" &&
    typeof args.font_size !== "number"
  ) {
    return failure("format_range", "Provide at least one of: bold, italic, font_size");
  }

  const editId = createId();
  const { bookmark, text } = await captureSelectionBookmark(editId);

  const pendingEdit: PendingEdit = {
    id: editId,
    toolName: "format_range",
    description: "Format selected text",
    before: text.slice(0, 200),
    after: JSON.stringify(args),
    status: "pending",
    bookmark,
    undo: buildUndo("format_range", bookmark, text),
    payload: { format: args },
  };

  const applied = await applyMutationNow(pendingEdit, autoApplyEdits);
  if (applied) return applied;

  return {
    success: true,
    output: { pendingApproval: true, format: args, message: "Formatting staged for approval." },
    pendingEdit,
  };
}

async function executeInsertComment(argsJson: string): Promise<ToolExecutionResult> {
  const args = parseArgs<{ comment?: string }>(argsJson);
  const comment = args.comment?.trim() ?? "";
  if (!comment) return failure("insert_comment", "comment is required");

  await insertCommentOnSelection(comment);
  return {
    success: true,
    output: {
      applied: true,
      preview: comment.slice(0, 200),
      message: "Comment added to the current selection.",
    },
  };
}

async function executeInsertTable(
  argsJson: string,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult> {
  const args = parseArgs<{
    rows?: number;
    columns?: number;
    cells?: string[][];
  }>(argsJson);
  const rows = Math.min(20, Math.max(1, Math.floor(args.rows ?? 0)));
  const columns = Math.min(10, Math.max(1, Math.floor(args.columns ?? 0)));
  if (!rows || !columns) {
    return failure("insert_table", "rows and columns are required (rows 1-20, columns 1-10)");
  }

  const editId = createId();
  let bookmark: string;
  let before: string;

  try {
    const captured = await captureSelectionBookmark(editId);
    bookmark = captured.bookmark;
    before = captured.text;
  } catch {
    const captured = await captureEndBookmark(editId);
    bookmark = captured.bookmark;
    before = "(end of document)";
  }

  const preview = `${rows}x${columns} table`;
  const pendingEdit: PendingEdit = {
    id: editId,
    toolName: "insert_table",
    description: `Insert ${rows}x${columns} table`,
    before,
    after: preview,
    status: "pending",
    bookmark,
    undo: buildUndo("insert_table", bookmark, before),
    payload: { rows, columns, cells: args.cells },
  };

  const applied = await applyMutationNow(pendingEdit, autoApplyEdits);
  if (applied) return applied;

  return {
    success: true,
    output: { pendingApproval: true, rows, columns, message: "Table insertion staged for approval." },
    pendingEdit,
  };
}

export async function applyPendingEdit(edit: PendingEdit): Promise<void> {
  if (!edit.bookmark) {
    throw new WordOperationError("Cannot apply edit: range bookmark is missing.");
  }

  switch (edit.toolName) {
    case "replace_text":
      await replaceBookmarkText(edit.bookmark, edit.after);
      break;
    case "insert_text": {
      const location = (edit.payload?.location as "selection" | "end" | undefined) ?? "selection";
      const insertLocation = location === "end" ? "after" : "before";
      await insertAtBookmark(edit.bookmark, edit.after, insertLocation);
      break;
    }
    case "delete_range":
      await deleteBookmarkRange(edit.bookmark);
      break;
    case "apply_style":
      await applyStyleAtBookmark(
        edit.bookmark,
        edit.payload?.style as DocumentStyleName,
      );
      break;
    case "format_range":
      await formatAtBookmark(edit.bookmark, edit.payload?.format as FormatOptions);
      break;
    case "insert_table":
      await insertTableAtBookmark(
        edit.bookmark,
        edit.payload?.rows as number,
        edit.payload?.columns as number,
        edit.payload?.cells as string[][] | undefined,
      );
      break;
    default:
      throw new WordOperationError(`Cannot apply edit for tool: ${edit.toolName}`);
  }
}

export async function rejectPendingEdit(edit: PendingEdit): Promise<void> {
  if (edit.status === "pending" && edit.bookmark) {
    await deleteBookmark(edit.bookmark);
  }
}

export async function undoPendingEdit(edit: PendingEdit): Promise<void> {
  if (!edit.undo) {
    throw new WordOperationError("No undo snapshot available for this edit.");
  }

  if (edit.toolName === "insert_table") {
    await undoInsertTable();
    if (edit.bookmark) await deleteBookmark(edit.bookmark);
    return;
  }

  await revertUndoSnapshot(edit.undo);
  if (edit.bookmark) await deleteBookmark(edit.bookmark);
}

async function undoInsertTable(): Promise<void> {
  if (typeof Word === "undefined") {
    throw new WordOperationError("Word APIs are unavailable outside Word.");
  }

  await Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    if (tables.items.length > 0) {
      tables.items[tables.items.length - 1].delete();
      await context.sync();
    }
  });
}