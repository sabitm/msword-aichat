import type { ContextMode, DocumentContext, TableSelectionContext } from "../types/context";
import { MAX_CONTEXT_CHARS } from "../types/context";

const HEADING_STYLES = new Set([
  "Heading1",
  "Heading2",
  "Heading3",
  "Heading4",
  "Heading5",
  "Heading6",
  "Title",
  "Subtitle",
]);

const HEADING_INDENT: Record<string, string> = {
  Title: "",
  Subtitle: "",
  Heading1: "",
  Heading2: "  ",
  Heading3: "    ",
  Heading4: "      ",
  Heading5: "        ",
  Heading6: "          ",
};

export function estimateTokens(text: string): number {
  if (!text.trim()) return 0;
  return Math.ceil(text.length / 4);
}

function truncateText(text: string): { text: string; truncated: boolean } {
  if (text.length <= MAX_CONTEXT_CHARS) {
    return { text, truncated: false };
  }
  return {
    text: `${text.slice(0, MAX_CONTEXT_CHARS)}\n\n[Context truncated]`,
    truncated: true,
  };
}

function emptyContext(mode: ContextMode, error?: string): DocumentContext {
  return {
    mode,
    text: "",
    tokenEstimate: 0,
    truncated: false,
    empty: true,
    error,
  };
}

function finalizeContext(mode: ContextMode, rawText: string): DocumentContext {
  const trimmed = rawText.trim();
  if (!trimmed) {
    return emptyContext(mode);
  }

  const { text, truncated } = truncateText(trimmed);
  return {
    mode,
    text,
    tokenEstimate: estimateTokens(text),
    truncated,
    empty: false,
  };
}

export function isWordApiAvailable(): boolean {
  return typeof Word !== "undefined";
}

function formatTableSelectionPromptHint(tableSelection: TableSelectionContext): string {
  const columnLabel =
    tableSelection.columnIndex === null ? "?" : String(tableSelection.columnIndex);
  const pinnedNote =
    tableSelection.selectionSource === "pinned" && tableSelection.bookmark
      ? ` Pinned bookmark=${tableSelection.bookmark} — prefer update_table with this pin active.`
      : "";
  const rowLabel =
    tableSelection.rowIndexEnd !== undefined &&
    tableSelection.rowIndexEnd !== tableSelection.rowIndex
      ? "row_index=" +
        tableSelection.rowIndex +
        "-" +
        tableSelection.rowIndexEnd
      : "row_index=" + tableSelection.rowIndex;
  const columnEndLabel =
    tableSelection.columnIndexEnd !== undefined &&
    tableSelection.columnIndex !== null &&
    tableSelection.columnIndexEnd !== tableSelection.columnIndex
      ? columnLabel + "-" + tableSelection.columnIndexEnd
      : columnLabel;
  return (
    "Table position (0-based): table_index=" +
    tableSelection.tableIndex +
    ", " +
    rowLabel +
    ", column_index=" +
    columnEndLabel +
    ", size=" +
    tableSelection.rows +
    "x" +
    tableSelection.columns +
    (tableSelection.isUniform ? "" : ", merged headers") +
    ". Table edits: call get_selection for table_values, then update_table — do not use get_document_text for table layout." +
    pinnedNote
  );
}

export interface GetDocumentContextOptions {
  preferPinned?: boolean;
}

export async function getSelectionText(
  options?: GetDocumentContextOptions,
): Promise<DocumentContext> {
  if (!isWordApiAvailable()) {
    return emptyContext("selection", "Word APIs are unavailable outside Word.");
  }

  try {
    const preferPinned = options?.preferPinned === true;
    const { readTableSelectionContext, USER_SELECTION_BOOKMARK } = await import("./operations");
    const { readUserSelectionBookmarkText } = await import("./ranges");
    const tableSelection = await readTableSelectionContext(
      preferPinned ? "pinned_or_live" : "live",
    );

    let rawText = "";
    let selectionPinned = false;
    let selectionBookmark: string | undefined;

    if (preferPinned) {
      const pinnedText = await readUserSelectionBookmarkText();
      if (pinnedText !== null) {
        rawText = pinnedText;
        selectionPinned = true;
        selectionBookmark = tableSelection?.bookmark ?? USER_SELECTION_BOOKMARK;
      }
    }

    if (!selectionPinned) {
      rawText = await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        return selection.text ?? "";
      });
    }

    const trimmed = rawText.trim();
    if (!trimmed && !tableSelection) {
      return emptyContext("selection");
    }

    const { text, truncated } = truncateText(trimmed);
    return {
      mode: "selection",
      text,
      tokenEstimate: estimateTokens(trimmed),
      truncated,
      empty: false,
      tableSelection,
      selectionPinned,
      selectionBookmark,
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to read selection";
    return emptyContext("selection", message);
  }
}

export async function getDocumentOutline(): Promise<DocumentContext> {
  if (!isWordApiAvailable()) {
    return emptyContext("outline", "Word APIs are unavailable outside Word.");
  }

  try {
    const rawText = await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      for (const paragraph of paragraphs.items) {
        paragraph.load(["text", "styleBuiltIn"]);
      }
      await context.sync();

      const lines: string[] = [];
      for (const paragraph of paragraphs.items) {
        const style = String(paragraph.styleBuiltIn);
        if (!HEADING_STYLES.has(style)) continue;

        const title = paragraph.text.trim();
        if (!title) continue;

        const indent = HEADING_INDENT[style] ?? "";
        lines.push(`${indent}- ${title}`);
      }

      return lines.join("\n");
    });

    return finalizeContext("outline", rawText);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to read outline";
    return emptyContext("outline", message);
  }
}

export async function getDocumentContext(
  mode: ContextMode,
  options?: GetDocumentContextOptions,
): Promise<DocumentContext> {
  switch (mode) {
    case "selection":
      return getSelectionText(options);
    case "outline":
      return getDocumentOutline();
    case "none":
      return emptyContext("none");
    default: {
      const exhaustive: never = mode;
      return emptyContext(exhaustive);
    }
  }
}

export function buildContextPrompt(context: DocumentContext): string | null {
  if (context.mode === "none" || context.empty) {
    return null;
  }

  const label = context.mode === "selection" ? "Selected text" : "Document outline";
  const truncatedNote = context.truncated ? "\nNote: context was truncated to fit token limits." : "";

  var parts: string[] = [];
  if (context.text) {
    parts.push(context.text);
  }
  if (context.tableSelection) {
    parts.push(formatTableSelectionPromptHint(context.tableSelection));
  }
  if (!parts.length) {
    return null;
  }

  return `${label} from the user's Word document:\n---\n${parts.join("\n\n")}\n---${truncatedNote}`;
}