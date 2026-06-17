import type {
  PendingEdit,
  ToolDefinition,
  ToolExecutionResult,
} from "../../types/agent";
import {
  readBodyTextChunk,
  readSelectionPlain,
  replaceSelectionText,
  insertText,
  WordOperationError,
} from "../../word/operations";

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

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "get_selection",
    description: "Read the currently selected text in the Word document.",
    parameters: {
      type: "object",
      properties: {},
      additionalProperties: false,
    },
  },
  {
    name: "get_document_text",
    description:
      "Read a chunk of the document body text. Use start and max_chars to page through long documents.",
    parameters: {
      type: "object",
      properties: {
        start: {
          type: "number",
          description: "Character offset to start reading from. Defaults to 0.",
        },
        max_chars: {
          type: "number",
          description: "Maximum characters to return. Defaults to 4000.",
        },
      },
      additionalProperties: false,
    },
  },
  {
    name: "insert_text",
    description:
      "Insert text into the document at the current selection or at the end of the document.",
    parameters: {
      type: "object",
      properties: {
        text: {
          type: "string",
          description: "Text to insert.",
        },
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
      "Replace the currently selected text with new text. Requires an active selection.",
    parameters: {
      type: "object",
      properties: {
        new_text: {
          type: "string",
          description: "Replacement text for the current selection.",
        },
      },
      required: ["new_text"],
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
      case "insert_text":
        return executeInsertText(argsJson, options.autoApplyEdits);
      case "replace_text":
        return executeReplaceText(argsJson, options.autoApplyEdits);
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
    output: {
      text,
      empty: !text.trim(),
      length: text.length,
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

async function executeInsertText(
  argsJson: string,
  autoApplyEdits: boolean,
): Promise<ToolExecutionResult> {
  const args = parseArgs<{ text?: string; location?: "selection" | "end" }>(argsJson);
  const text = args.text?.trim() ?? "";
  const location = args.location ?? "selection";

  if (!text) {
    return failure("insert_text", "text is required");
  }

  if (autoApplyEdits) {
    await insertText(text, location);
    return {
      success: true,
      output: {
        applied: true,
        location,
        insertedLength: text.length,
        preview: text.slice(0, 200),
      },
    };
  }

  const before = location === "end" ? "(end of document)" : await readSelectionPlain();
  const pendingEdit: PendingEdit = {
    id: createId(),
    toolName: "insert_text",
    description: `Insert text at ${location}`,
    before,
    after: text,
    status: "pending",
  };

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

  if (!newText.trim()) {
    return failure("replace_text", "new_text is required");
  }

  const before = await readSelectionPlain();
  if (!before.trim()) {
    return failure("replace_text", "No text selected. Ask the user to select text first.");
  }

  if (autoApplyEdits) {
    await replaceSelectionText(newText);
    return {
      success: true,
      output: {
        applied: true,
        replacedLength: before.length,
        preview: newText.slice(0, 200),
      },
    };
  }

  const pendingEdit: PendingEdit = {
    id: createId(),
    toolName: "replace_text",
    description: "Replace selected text",
    before,
    after: newText,
    status: "pending",
  };

  return {
    success: true,
    output: {
      pendingApproval: true,
      beforeLength: before.length,
      afterLength: newText.length,
      message: "Replacement staged for user approval.",
    },
    pendingEdit,
  };
}

export async function applyPendingEdit(edit: PendingEdit): Promise<void> {
  if (edit.toolName === "replace_text") {
    await replaceSelectionText(edit.after);
    return;
  }

  if (edit.toolName === "insert_text") {
    const location = edit.before === "(end of document)" ? "end" : "selection";
    await insertText(edit.after, location);
    return;
  }

  throw new WordOperationError(`Cannot apply edit for tool: ${edit.toolName}`);
}