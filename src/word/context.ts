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

function formatTableSelectionBlock(tableSelection: TableSelectionContext): string {
  const columnLabel =
    tableSelection.columnIndex === null ? "unknown" : String(tableSelection.columnIndex);
  return (
    "Table selection (0-based indices):\n" +
    `- table_index: ${tableSelection.tableIndex}\n` +
    `- row_index: ${tableSelection.rowIndex}\n` +
    `- column_index: ${columnLabel}\n` +
    `- table size: ${tableSelection.rows}x${tableSelection.columns}\n` +
    `- isUniform: ${tableSelection.isUniform}\n` +
    `- row values: ${JSON.stringify(tableSelection.rowValues)}\n` +
    "Use update_table with start_row = row_index to patch rows from the selection, or pass a full cells grid."
  );
}

export async function getSelectionText(): Promise<DocumentContext> {
  if (!isWordApiAvailable()) {
    return emptyContext("selection", "Word APIs are unavailable outside Word.");
  }

  try {
    const { readTableSelectionContext } = await import("./operations");
    const tableSelection = await readTableSelectionContext();
    const rawText = await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return selection.text ?? "";
    });
    const contextText = tableSelection
      ? `${rawText}\n\n${formatTableSelectionBlock(tableSelection)}`
      : rawText;
    const result = finalizeContext("selection", contextText);
    result.tableSelection = tableSelection;
    return result;
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

export async function getDocumentContext(mode: ContextMode): Promise<DocumentContext> {
  switch (mode) {
    case "selection":
      return getSelectionText();
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
  if (context.mode === "none") {
    return null;
  }
  if (context.empty && !context.tableSelection) {
    return null;
  }
  if (!context.text && !context.tableSelection) {
    return null;
  }

  const label = context.mode === "selection" ? "Selected text" : "Document outline";
  const truncatedNote = context.truncated ? "\nNote: context was truncated to fit token limits." : "";
  const body = context.text || "(empty selection text)";

  return `${label} from the user's Word document:\n---\n${body}\n---${truncatedNote}`;
}