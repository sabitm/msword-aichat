import type {
  DocumentStyleName,
  FindReplaceMatchSnapshot,
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
  listDocumentTables,
  readBodyTextChunk,
  readSelectionPlain,
  readSelectionStyle,
  selectionContainsTable,
  readTableAtIndex,
  readTableSelectionContext,
  resolveTableIndex,
  searchDocument,
  updateTableAtIndex,
  WordOperationError,
} from "../../word/operations";
import {
  applyStagedFindReplacements,
  captureEndBookmark,
  captureSelectionBookmark,
  deleteBookmark,
  deleteBookmarkRange,
  deleteBookmarks,
  insertAtBookmark,
  replaceBookmarkText,
  stageFindReplacements,
  type FindReplaceStageResult,
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

function buildFindReplaceUndo(
  kind: "find_and_replace" | "replace_at_match",
  snapshots: FindReplaceMatchSnapshot[],
): UndoSnapshot {
  return {
    kind: kind,
    bookmark: snapshots[0] ? snapshots[0].bookmark : "",
    previousText: "",
    matchSnapshots: snapshots,
  };
}

function buildFindReplaceSamples(
  snapshots: FindReplaceMatchSnapshot[],
  limit: number,
): Array<{ index: number; before: string }> {
  var samples: Array<{ index: number; before: string }> = [];
  var max = Math.min(limit, snapshots.length);
  for (var i = 0; i < max; i += 1) {
    samples.push({
      index: snapshots[i].searchIndex,
      before: snapshots[i].previousText,
    });
  }
  return samples;
}

function buildFindReplaceOutput(
  staged: FindReplaceStageResult,
  count: number,
  pendingApproval: boolean,
): Record<string, unknown> {
  return {
    find: staged.find,
    replace: staged.replace,
    totalMatches: staged.totalMatches,
    matchCount: count,
    replacedCount: count,
    skippedInTables: staged.skippedInTables,
    samples: buildFindReplaceSamples(staged.staged, 3),
    pendingApproval: pendingApproval,
    message: pendingApproval
      ? "Replacement staged for user approval."
      : "Replacement applied.",
  };
}

async function createFindReplacePendingEdit(
  toolName: "find_and_replace" | "replace_at_match",
  editId: string,
  staged: FindReplaceStageResult,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult> {
  var snapshots = staged.staged;
  var count = snapshots.length;
  var find = staged.find;
  var replace = staged.replace;

  var pendingEdit: PendingEdit = {
    id: editId,
    toolName: toolName,
    description:
      toolName === "find_and_replace"
        ? 'Replace "' + find + '" → "' + replace + '" (' + count + "x)"
        : 'Replace match ' + snapshots[0].searchIndex + ': "' + find + '" → "' + replace + '"',
    before:
      toolName === "find_and_replace"
        ? 'Find: "' + find + '" (' + count + " match" + (count === 1 ? "" : "es") + ")"
        : snapshots[0].previousText,
    after: toolName === "find_and_replace" ? 'Replace with: "' + replace + '"' : replace,
    status: "pending",
    bookmark: snapshots[0].bookmark,
    undo: buildFindReplaceUndo(toolName, snapshots),
    payload: {
      find: find,
      replace: replace,
      matchSnapshots: snapshots,
    },
  };

  var applied = await applyMutationNow(pendingEdit, autoApplyEdits);
  if (applied) {
    return applied;
  }

  return {
    success: true,
    output: buildFindReplaceOutput(staged, count, true),
    pendingEdit: pendingEdit,
  };
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
    description:
      "Read the currently selected text. When the selection is inside a table, also returns table_index, row_index, column_index, and the full row values.",
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
      "Replace the currently selected plain text with new text. Do not use on tables — use update_table instead.",
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
    name: "find_and_replace",
    description:
      "Find all occurrences of text in the document body and replace them. Skips table cells by default. Use for bulk changes such as dates or repeated phrases.",
    parameters: {
      type: "object",
      properties: {
        find: { type: "string", description: "Text to search for." },
        replace: { type: "string", description: "Replacement text (empty string deletes matches)." },
        max_replacements: {
          type: "number",
          description: "Maximum occurrences to replace. Defaults to 50 (hard max 100).",
        },
        match_case: { type: "boolean", description: "Case-sensitive search. Defaults to false." },
        skip_tables: {
          type: "boolean",
          description: "Skip matches inside tables. Defaults to true.",
        },
      },
      required: ["find", "replace"],
      additionalProperties: false,
    },
  },
  {
    name: "replace_at_match",
    description:
      "Replace one specific search match by 0-based match_index. Skips table cells by default. Use search_document first to inspect matches.",
    parameters: {
      type: "object",
      properties: {
        find: { type: "string", description: "Text to search for." },
        replace: { type: "string", description: "Replacement text." },
        match_index: {
          type: "number",
          description: "0-based index of the match to replace (same order as search_document).",
        },
        match_case: { type: "boolean", description: "Case-sensitive search. Defaults to false." },
        skip_tables: {
          type: "boolean",
          description: "Fail if the match is inside a table when true. Defaults to true.",
        },
      },
      required: ["find", "replace", "match_index"],
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
      "Add a Word review comment on the current selection. On Word 2016 (no WordApi 1.4), highlights the selection and returns the suggestion text instead. Requires an active selection.",
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
  {
    name: "list_tables",
    description:
      "List tables in the document with index, dimensions, cell values, and a short preview. Use before update_table when multiple tables exist.",
    parameters: {
      type: "object",
      properties: {
        max_tables: {
          type: "number",
          description: "Maximum tables to return. Defaults to 10.",
        },
      },
      additionalProperties: false,
    },
  },
  {
    name: "update_table",
    description:
      "Replace cell contents of an existing table in place. Pass a full cells grid, or pass fewer rows with start_row (defaults to the selected row from get_selection). Use list_tables for the full table snapshot.",
    parameters: {
      type: "object",
      properties: {
        table_index: {
          type: "number",
          description:
            "0-based table index in the document. Omit to update the table containing the selection, or the first table.",
        },
        start_row: {
          type: "number",
          description:
            "0-based row to start writing cells when not passing a full grid. Defaults to the current table selection row.",
        },
        rows: {
          type: "number",
          description: "Optional row count hint (actual table size is used).",
        },
        columns: {
          type: "number",
          description: "Optional column count hint (actual table size is used).",
        },
        cells: {
          type: "array",
          description: "2D array of cell text values (full table or a row patch from start_row).",
          items: {
            type: "array",
            items: { type: "string" },
          },
        },
      },
      required: ["cells"],
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
      case "find_and_replace":
        return executeFindAndReplace(argsJson, options.autoApplyEdits);
      case "replace_at_match":
        return executeReplaceAtMatch(argsJson, options.autoApplyEdits);
      case "delete_range":
        return executeDeleteRange(options.autoApplyEdits);
      case "apply_style":
        return executeApplyStyle(argsJson, options.autoApplyEdits);
      case "format_range":
        return executeFormatRange(argsJson, options.autoApplyEdits);
      case "insert_table":
        return executeInsertTable(argsJson, options.autoApplyEdits);
      case "list_tables":
        return executeListTables(argsJson);
      case "update_table":
        return executeUpdateTable(argsJson, options.autoApplyEdits);
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
  const tableSelection = await readTableSelectionContext();
  const inTable = tableSelection !== null;
  return {
    success: true,
    output: {
      text,
      empty: !text.trim(),
      length: text.length,
      inTable,
      ...(tableSelection
        ? {
            table_index: tableSelection.tableIndex,
            row_index: tableSelection.rowIndex,
            column_index: tableSelection.columnIndex,
            table_rows: tableSelection.rows,
            table_columns: tableSelection.columns,
            is_uniform: tableSelection.isUniform,
            row_values: tableSelection.rowValues,
            hint:
              "Selection is inside a table. Call list_tables for the full grid, then update_table with start_row = row_index to patch rows from the selection (or pass a full cells grid). Do not use replace_text.",
          }
        : {}),
    },
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

async function executeFindAndReplace(
  argsJson: string,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult> {
  var args = parseArgs<{
    find?: string;
    replace?: string;
    max_replacements?: number;
    match_case?: boolean;
    skip_tables?: boolean;
  }>(argsJson);
  var find = args.find?.trim() ?? "";
  if (!find) return failure("find_and_replace", "find is required");
  if (args.replace === undefined || args.replace === null) {
    return failure("find_and_replace", "replace is required");
  }

  var editId = createId();
  var staged = await stageFindReplacements(editId, find, args.replace, {
    matchCase: Boolean(args.match_case),
    maxReplacements: args.max_replacements,
    skipTables: args.skip_tables !== false,
  });

  if (staged.staged.length === 0) {
    return {
      success: true,
      output: {
        find: staged.find,
        replace: staged.replace,
        totalMatches: staged.totalMatches,
        matchCount: 0,
        replacedCount: 0,
        skippedInTables: staged.skippedInTables,
        message:
          staged.totalMatches === 0
            ? "No matches found."
            : "No replaceable matches (all matches were inside tables).",
      },
    };
  }

  return createFindReplacePendingEdit("find_and_replace", editId, staged, autoApplyEdits);
}

async function executeReplaceAtMatch(
  argsJson: string,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult> {
  var args = parseArgs<{
    find?: string;
    replace?: string;
    match_index?: number;
    match_case?: boolean;
    skip_tables?: boolean;
  }>(argsJson);
  var find = args.find?.trim() ?? "";
  if (!find) return failure("replace_at_match", "find is required");
  if (args.replace === undefined || args.replace === null) {
    return failure("replace_at_match", "replace is required");
  }
  if (args.match_index === undefined || args.match_index === null) {
    return failure("replace_at_match", "match_index is required");
  }

  var editId = createId();
  var staged = await stageFindReplacements(editId, find, args.replace, {
    matchCase: Boolean(args.match_case),
    matchIndex: Math.floor(args.match_index),
    skipTables: args.skip_tables !== false,
  });

  if (staged.staged.length === 0) {
    return failure("replace_at_match", "No match staged for replacement.");
  }

  return createFindReplacePendingEdit("replace_at_match", editId, staged, autoApplyEdits);
}

async function executeReplaceText(
  argsJson: string,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult> {
  const args = parseArgs<{ new_text?: string }>(argsJson);
  const newText = args.new_text ?? "";
  if (!newText.trim()) return failure("replace_text", "new_text is required");

  if (await selectionContainsTable()) {
    return failure(
      "replace_text",
      "Selection is inside a table. Call list_tables, then update_table with the full cells grid (same rows/columns).",
    );
  }

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

  const result = await insertCommentOnSelection(comment);
  if (result.mode === "highlight_fallback") {
    return {
      success: true,
      output: {
        applied: true,
        preview: comment.slice(0, 200),
        fallback: "highlight",
        message:
          "Word 2016 cannot add review comments via the API (requires WordApi 1.4). The selection was highlighted yellow; tell the user the suggestion text shown here.",
        comment,
      },
    };
  }

  return {
    success: true,
    output: {
      applied: true,
      preview: comment.slice(0, 200),
      message: "Comment added to the current selection.",
    },
  };
}

async function executeListTables(argsJson: string): Promise<ToolExecutionResult> {
  const args = parseArgs<{ max_tables?: number }>(argsJson);
  const maxTables = Math.min(20, Math.max(1, Math.floor(args.max_tables ?? 10)));
  const tables = await listDocumentTables(maxTables);
  return {
    success: true,
    output: {
      tableCount: tables.length,
      tables: tables.map((table) => ({
        index: table.index,
        rows: table.rows,
        columns: table.columns,
        isUniform: table.isUniform,
        preview: table.preview,
        values: table.values,
      })),
    },
  };
}

function formatTableEditPreview(values: string[][]): string {
  if (!values.length) return "(empty table)";
  const maxRows = 6;
  const lines = values.slice(0, maxRows).map((row) => row.join(" | "));
  if (values.length > maxRows) lines.push("...");
  return lines.join("\n");
}

function normalizeUpdateTableCells(
  cellValues: string[][] | undefined,
  rows: number,
  columns: number,
): string[][] | null {
  if (!cellValues?.length || rows < 1 || columns < 1) return null;
  const grid: string[][] = [];
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const source = cellValues[rowIndex] ?? [];
    const line: string[] = [];
    for (let colIndex = 0; colIndex < columns; colIndex += 1) {
      line.push(source[colIndex] ?? "");
    }
    grid.push(line);
  }
  return grid;
}

function mergeTableRowPatch(
  currentValues: string[][],
  patch: string[][],
  startRow: number,
  columns: number,
): string[][] | null {
  if (!patch.length || startRow < 0) return null;
  const merged = currentValues.map(function (row) {
    return row.slice();
  });
  for (let patchIndex = 0; patchIndex < patch.length; patchIndex += 1) {
    const targetRow = startRow + patchIndex;
    if (targetRow >= merged.length) break;
    const source = patch[patchIndex] ?? [];
    for (let colIndex = 0; colIndex < columns; colIndex += 1) {
      merged[targetRow][colIndex] = source[colIndex] ?? "";
    }
  }
  return merged;
}

async function executeUpdateTable(
  argsJson: string,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult> {
  const args = parseArgs<{
    table_index?: number;
    start_row?: number;
    rows?: number;
    columns?: number;
    cells?: string[][];
  }>(argsJson);
  if (!args.cells?.length) {
    return failure("update_table", "cells 2D array is required");
  }

  let tableIndex: number;
  try {
    const requested =
      args.table_index !== undefined && args.table_index !== null
        ? Math.floor(args.table_index)
        : undefined;
    tableIndex = await resolveTableIndex(requested);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Could not find a table to update.";
    return failure("update_table", message);
  }

  let current;
  try {
    current = await readTableAtIndex(tableIndex);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Could not read the target table.";
    return failure("update_table", message);
  }

  const rows = current.rows;
  const columns = current.columns;
  let normalizedCells: string[][] | null = null;

  if (args.cells.length >= rows) {
    normalizedCells = normalizeUpdateTableCells(args.cells, rows, columns);
  } else {
    let startRow =
      args.start_row !== undefined && args.start_row !== null
        ? Math.floor(args.start_row)
        : null;
    if (startRow === null) {
      const tableSelection = await readTableSelectionContext();
      if (tableSelection && tableSelection.tableIndex === tableIndex) {
        startRow = tableSelection.rowIndex;
      }
    }
    if (startRow === null || startRow < 0) {
      return failure(
        "update_table",
        "Partial row update needs start_row or a selection inside the target table. Call get_selection first.",
      );
    }
    if (startRow + args.cells.length > rows) {
      return failure(
        "update_table",
        `Patch rows ${startRow}-${startRow + args.cells.length - 1} exceed table row count (${rows}).`,
      );
    }
    const patch = normalizeUpdateTableCells(args.cells, args.cells.length, columns);
    if (!patch) {
      return failure("update_table", "cells 2D array is required");
    }
    normalizedCells = mergeTableRowPatch(current.values, patch, startRow, columns);
  }

  if (!normalizedCells) {
    return failure("update_table", "cells 2D array is required");
  }

  const editId = createId();
  const captured = await captureEndBookmark(editId);

  const pendingEdit: PendingEdit = {
    id: editId,
    toolName: "update_table",
    description: `Update table ${tableIndex} (${rows}x${columns})`,
    before: formatTableEditPreview(current.values),
    after: formatTableEditPreview(normalizedCells),
    status: "pending",
    bookmark: captured.bookmark,
    undo: {
      kind: "update_table",
      bookmark: captured.bookmark,
      previousText: "",
      tableIndex,
      previousTableValues: current.values,
    },
    payload: { tableIndex, rows, columns, cells: normalizedCells },
  };

  const applied = await applyMutationNow(pendingEdit, autoApplyEdits);
  if (applied) {
    return {
      success: true,
      output: {
        applied: true,
        tableIndex,
        rows,
        columns,
        message: `Table ${tableIndex} updated in place.`,
      },
    };
  }

  return {
    success: true,
    output: {
      pendingApproval: true,
      tableIndex,
      rows,
      columns,
      message: `Table ${tableIndex} update staged for approval.`,
    },
    pendingEdit,
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
    case "find_and_replace":
    case "replace_at_match": {
      var matchSnapshots = edit.payload?.matchSnapshots as FindReplaceMatchSnapshot[] | undefined;
      var replaceValue = edit.payload?.replace as string | undefined;
      if (!matchSnapshots?.length || replaceValue === undefined) {
        throw new WordOperationError("Cannot apply find/replace edit: payload is incomplete.");
      }
      await applyStagedFindReplacements(matchSnapshots, replaceValue);
      break;
    }
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
    case "update_table":
      await updateTableAtIndex(
        edit.payload?.tableIndex as number,
        edit.payload?.rows as number,
        edit.payload?.columns as number,
        edit.payload?.cells as string[][],
      );
      break;
    default:
      throw new WordOperationError(`Cannot apply edit for tool: ${edit.toolName}`);
  }
}

export async function rejectPendingEdit(edit: PendingEdit): Promise<void> {
  if (edit.status !== "pending") {
    return;
  }

  var matchSnapshots = edit.payload?.matchSnapshots as FindReplaceMatchSnapshot[] | undefined;
  if (matchSnapshots?.length) {
    await deleteBookmarks(
      matchSnapshots.map(function (snapshot) {
        return snapshot.bookmark;
      }),
    );
    return;
  }

  if (edit.bookmark) {
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

  if (edit.toolName === "update_table") {
    await revertUndoSnapshot(edit.undo);
    if (edit.bookmark) await deleteBookmark(edit.bookmark);
    return;
  }

  if (edit.toolName === "find_and_replace" || edit.toolName === "replace_at_match") {
    await revertUndoSnapshot(edit.undo);
    var snapshots = edit.payload?.matchSnapshots as FindReplaceMatchSnapshot[] | undefined;
    if (snapshots?.length) {
      await deleteBookmarks(
        snapshots.map(function (snapshot) {
          return snapshot.bookmark;
        }),
      );
    }
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