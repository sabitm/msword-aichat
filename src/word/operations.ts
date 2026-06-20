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
  if (/itemnotfound/i.test(message)) {
    throw new WordOperationError(
      "Word could not find the table cell to update. Try the request again or insert the table at the document end.",
    );
  }
  if (/invalidargument/i.test(message)) {
    throw new WordOperationError(
      "Word rejected the table update. Tables with merged header cells cannot be replaced in one shot — update only the data rows that changed.",
    );
  }
  if (/generalexception/i.test(message)) {
    throw new WordOperationError(fallback);
  }
  throw new WordOperationError(message);
}

function isBulkTableWriteError(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error);
  return /invalidargument|generalexception/i.test(message);
}

async function writeTableValuesPerCell(
  table: Word.Table,
  grid: string[][],
  previous: string[][],
  context: Word.RequestContext,
): Promise<void> {
  const rows = grid.length;
  const columns = rows > 0 ? (grid[0]?.length ?? 0) : 0;
  const updates: { cell: Word.TableCell; value: string }[] = [];

  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    for (let colIndex = 0; colIndex < columns; colIndex += 1) {
      const next = grid[rowIndex][colIndex] ?? "";
      const prev = previous[rowIndex]?.[colIndex] ?? "";
      if (next === prev) continue;

      const cell = table.getCellOrNullObject(rowIndex, colIndex);
      updates.push({ cell, value: next });
    }
  }

  if (!updates.length) {
    return;
  }

  for (let index = 0; index < updates.length; index += 1) {
    updates[index].cell.load("isNullObject");
  }
  await context.sync();

  let wrote = 0;
  for (let index = 0; index < updates.length; index += 1) {
    if (!updates[index].cell.isNullObject) {
      updates[index].cell.body.insertText(updates[index].value, Word.InsertLocation.replace);
      wrote += 1;
    }
  }

  if (!wrote) {
    throw new WordOperationError(
      "Could not update any table cells. The table may use merged cells that block bulk edits.",
    );
  }

  await context.sync();
}

async function writeTableValues(
  table: Word.Table,
  grid: string[][],
  previous: string[][],
  context: Word.RequestContext,
): Promise<void> {
  table.load("isUniform");
  await context.sync();

  if (table.isUniform !== false) {
    try {
      table.values = grid;
      await context.sync();
      return;
    } catch (error) {
      if (!isBulkTableWriteError(error)) {
        throw error;
      }
    }
  }

  await writeTableValuesPerCell(table, grid, previous, context);
}

export function isWordCommentInsertSupported(): boolean {
  if (typeof Office === "undefined") return false;
  return Office.context.requirements.isSetSupported("WordApi", "1.4");
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

export async function selectionContainsTable(): Promise<boolean> {
  assertWordAvailable();
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const parentTable = selection.parentTableOrNullObject;
      parentTable.load("isNullObject");
      await context.sync();
      return !parentTable.isNullObject;
    });
  } catch (error) {
    wrapWordError(error, "Failed to inspect selection.");
  }
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

      const body = context.document.body;
      body.load("text");
      for (let index = 0; index < limit; index += 1) {
        results.items[index].load("text");
      }
      await context.sync();

      const fullText = body.text ?? "";
      let scanAt = 0;

      for (let index = 0; index < limit; index += 1) {
        const text = results.items[index].text ?? "";
        let start = -1;
        if (text) {
          start = fullText.indexOf(text, scanAt);
          if (start < 0) start = fullText.indexOf(text);
          if (start >= 0) scanAt = start + Math.max(text.length, 1);
        }
        const end = start >= 0 ? start + text.length : -1;
        matches.push({ index, text, start, end });
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

export type CommentInsertResult =
  | { mode: "comment" }
  | { mode: "highlight_fallback"; comment: string };

export async function insertCommentOnSelection(comment: string): Promise<CommentInsertResult> {
  assertWordAvailable();
  const trimmed = comment.trim();
  if (!trimmed) {
    throw new WordOperationError("Comment text is required.");
  }

  if (!isWordCommentInsertSupported()) {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (!selection.text?.trim()) {
          throw new WordOperationError("Select text in the document before adding a comment.");
        }

        // Word 2016 lacks Range.insertComment (WordApi 1.4). Highlight instead.
        selection.font.highlightColor = "#FFFF00";
        await context.sync();
      });
      return { mode: "highlight_fallback", comment: trimmed };
    } catch (error) {
      wrapWordError(
        error,
        "Failed to mark the selection for review (Word 2016 comment API unavailable).",
      );
    }
  }

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (!selection.text?.trim()) {
        throw new WordOperationError("Select text in the document before adding a comment.");
      }

      selection.insertComment(trimmed);
      await context.sync();
    });
    return { mode: "comment" };
  } catch (error) {
    wrapWordError(error, "Failed to insert comment on the selection.");
  }
}

export interface DocumentTableInfo {
  index: number;
  rows: number;
  columns: number;
  values: string[][];
  preview: string;
  isUniform?: boolean;
}

function cloneCellGrid(values: string[][]): string[][] {
  return values.map((row) => row.slice());
}

/** Dimensions Word accepts for table.values writes (merged cells may return ragged rows on read). */
function getTableWriteDimensions(
  values: string[][],
  rowCount?: number,
): { rows: number; columns: number } {
  const rows = rowCount !== undefined && rowCount > 0 ? rowCount : values.length;
  const columns = rows > 0 ? (values[0]?.length ?? 0) : 0;
  return { rows, columns };
}

function normalizeTableValues(values: string[][], rows: number, columns: number): string[][] {
  const grid: string[][] = [];
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const source = values[rowIndex] ?? [];
    const line: string[] = [];
    for (let colIndex = 0; colIndex < columns; colIndex += 1) {
      line.push(source[colIndex] ?? "");
    }
    grid.push(line);
  }
  return grid;
}

function formatTablePreview(values: string[][], maxRows = 4): string {
  if (!values.length) return "(empty table)";
  const lines = values.slice(0, maxRows).map((row) => row.join(" | "));
  if (values.length > maxRows) lines.push("...");
  return lines.join("\n");
}

function normalizeCellGrid(
  cellValues: string[][] | undefined,
  rows: number,
  columns: number,
): string[][] | null {
  if (!cellValues?.length) return null;
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

export async function insertTableAtBookmark(
  bookmark: string,
  rows: number,
  columns: number,
  cellValues?: string[][],
): Promise<void> {
  assertWordAvailable();
  const grid = normalizeCellGrid(cellValues, rows, columns);

  try {
    await Word.run(async (context) => {
      const range = context.document.getBookmarkRange(bookmark);
      const table = range.insertTable(rows, columns, Word.InsertLocation.after);
      // Materialize the table before writing cells (Word 2016 needs this).
      await context.sync();

      if (grid) {
        // Bulk assign — avoids per-cell getCell/rows.items ItemNotFound on IE WebView.
        table.values = grid;
        await context.sync();
      }
    });
  } catch (error) {
    wrapWordError(error, "Failed to insert table.");
  }
}

export async function listDocumentTables(maxTables = 10): Promise<DocumentTableInfo[]> {
  assertWordAvailable();
  const limit = Math.min(20, Math.max(1, Math.floor(maxTables)));

  try {
    return await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      const count = Math.min(tables.items.length, limit);
      for (let index = 0; index < count; index += 1) {
        tables.items[index].load(["values", "rowCount", "isUniform"]);
      }
      await context.sync();

      const infos: DocumentTableInfo[] = [];
      for (let index = 0; index < count; index += 1) {
        const rawValues = (tables.items[index].values as string[][]) ?? [];
        const dimensions = getTableWriteDimensions(rawValues, tables.items[index].rowCount);
        const values = normalizeTableValues(rawValues, dimensions.rows, dimensions.columns);
        infos.push({
          index,
          rows: dimensions.rows,
          columns: dimensions.columns,
          values,
          preview: formatTablePreview(values),
          isUniform: tables.items[index].isUniform,
        });
      }
      return infos;
    });
  } catch (error) {
    wrapWordError(error, "Failed to list tables.");
  }
}

export async function resolveTableIndex(requested?: number): Promise<number> {
  assertWordAvailable();

  try {
    return await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      const selection = context.document.getSelection();
      const parentTable = selection.parentTableOrNullObject;
      parentTable.load("isNullObject");
      await context.sync();

      if (tables.items.length === 0) {
        throw new WordOperationError("No tables found in the document.");
      }

      if (requested !== undefined) {
        const index = Math.floor(requested);
        if (index < 0 || index >= tables.items.length) {
          throw new WordOperationError(
            `Table index ${index} is out of range (document has ${tables.items.length} table(s)).`,
          );
        }
        return index;
      }

      if (!parentTable.isNullObject) {
        for (let index = 0; index < tables.items.length; index += 1) {
          if (tables.items[index] === parentTable) return index;
        }
      }

      return 0;
    });
  } catch (error) {
    wrapWordError(error, "Failed to resolve table.");
  }
}

export async function readTableAtIndex(tableIndex: number): Promise<DocumentTableInfo> {
  assertWordAvailable();

  try {
    return await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tableIndex < 0 || tableIndex >= tables.items.length) {
        throw new WordOperationError(
          `Table index ${tableIndex} is out of range (document has ${tables.items.length} table(s)).`,
        );
      }

      const table = tables.items[tableIndex];
      table.load(["values", "rowCount", "isUniform"]);
      await context.sync();

      const rawValues = (table.values as string[][]) ?? [];
      const dimensions = getTableWriteDimensions(rawValues, table.rowCount);
      const values = normalizeTableValues(rawValues, dimensions.rows, dimensions.columns);

      return {
        index: tableIndex,
        rows: dimensions.rows,
        columns: dimensions.columns,
        values,
        preview: formatTablePreview(values),
        isUniform: table.isUniform,
      };
    });
  } catch (error) {
    wrapWordError(error, "Failed to read table.");
  }
}

export async function updateTableAtIndex(
  tableIndex: number,
  rows: number,
  columns: number,
  cellValues: string[][],
): Promise<void> {
  assertWordAvailable();
  const grid = normalizeCellGrid(cellValues, rows, columns);
  if (!grid) {
    throw new WordOperationError("cells 2D array is required for update_table.");
  }

  try {
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tableIndex < 0 || tableIndex >= tables.items.length) {
        throw new WordOperationError(
          `Table index ${tableIndex} is out of range (document has ${tables.items.length} table(s)).`,
        );
      }

      const table = tables.items[tableIndex];
      table.load(["values", "rowCount"]);
      await context.sync();

      const current = (table.values as string[][]) ?? [];
      const dimensions = getTableWriteDimensions(current, table.rowCount);

      if (dimensions.rows !== rows || dimensions.columns !== columns) {
        throw new WordOperationError(
          `Table ${tableIndex} is ${dimensions.rows}x${dimensions.columns}. Pass matching rows and columns.`,
        );
      }

      const previous = normalizeTableValues(current, dimensions.rows, dimensions.columns);
      await writeTableValues(table, grid, previous, context);
    });
  } catch (error) {
    wrapWordError(error, "Failed to update table.");
  }
}

export async function restoreTableAtIndex(tableIndex: number, values: string[][]): Promise<void> {
  assertWordAvailable();

  try {
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tableIndex < 0 || tableIndex >= tables.items.length) {
        throw new WordOperationError(`Cannot undo: table index ${tableIndex} no longer exists.`);
      }

      const table = tables.items[tableIndex];
      table.load(["values", "rowCount"]);
      await context.sync();

      const current = (table.values as string[][]) ?? [];
      const dimensions = getTableWriteDimensions(current, table.rowCount);
      const previous = normalizeTableValues(current, dimensions.rows, dimensions.columns);
      const grid = normalizeTableValues(cloneCellGrid(values), dimensions.rows, dimensions.columns);
      await writeTableValues(table, grid, previous, context);
    });
  } catch (error) {
    wrapWordError(error, "Failed to restore table.");
  }
}