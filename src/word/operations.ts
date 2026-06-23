import type { DocumentStyleName } from "../types/agent";
import type {
  TableIndexResolution,
  TableSelectionContext,
  TableSelectionSource,
} from "../types/context";
import { isWordApiAvailable } from "./context";

export const USER_SELECTION_BOOKMARK = "msword_aichat_user_select";

export type TableSelectionReadSource = "live" | "pinned" | "pinned_or_live";

export interface TableSelectionReadOptions {
  source?: TableSelectionReadSource;
  /** When true, reads every cell (get_selection). UI refresh keeps false. */
  includeTableValues?: boolean;
}

function normalizeTableSelectionReadOptions(
  sourceOrOptions: TableSelectionReadSource | TableSelectionReadOptions = "pinned_or_live",
): Required<TableSelectionReadOptions> {
  if (typeof sourceOrOptions === "string") {
    return { source: sourceOrOptions, includeTableValues: false };
  }
  return {
    source: sourceOrOptions.source ?? "pinned_or_live",
    includeTableValues: sourceOrOptions.includeTableValues === true,
  };
}

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

function inferTableColumnIndex(
  rowValues: string[],
  selectionText: string,
  cellText: string,
): number | null {
  const needle = (cellText || selectionText).trim();
  if (!needle) return null;
  for (let columnIndex = 0; columnIndex < rowValues.length; columnIndex += 1) {
    const cell = (rowValues[columnIndex] ?? "").trim();
    if (!cell) continue;
    if (cell === needle || cell.indexOf(needle) >= 0 || needle.indexOf(cell) >= 0) {
      return columnIndex;
    }
  }
  return null;
}

function findTableRowIndexBySelection(values: string[][], selectionText: string): number | null {
  const needle = selectionText.trim();
  if (!needle) return null;
  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex] ?? [];
    const rowText = row.join("\t");
    if (rowText.indexOf(needle) >= 0) {
      return rowIndex;
    }
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      const cell = (row[columnIndex] ?? "").trim();
      if (cell && (needle.indexOf(cell) >= 0 || cell.indexOf(needle) >= 0)) {
        return rowIndex;
      }
    }
  }
  return null;
}

function findTableRowIndexByCellText(values: string[][], cellText: string): number | null {
  const needle = cellText.trim();
  if (!needle) return null;
  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex] ?? [];
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if ((row[columnIndex] ?? "").trim() === needle) {
        return rowIndex;
      }
    }
  }
  return null;
}

function tableRowValuesMatch(row: string[] | undefined, expected: string[]): boolean {
  if (!row?.length || !expected.length) return false;
  let compared = 0;
  for (let columnIndex = 0; columnIndex < expected.length; columnIndex += 1) {
    const expectedCell = (expected[columnIndex] ?? "").trim();
    if (!expectedCell) continue;
    compared += 1;
    if ((row[columnIndex] ?? "").trim() !== expectedCell) {
      return false;
    }
  }
  return compared > 0;
}

function tableValuesEqual(left: string[][], right: string[][]): boolean {
  if (left.length !== right.length) return false;
  for (let rowIndex = 0; rowIndex < left.length; rowIndex += 1) {
    const leftRow = left[rowIndex] ?? [];
    const rightRow = right[rowIndex] ?? [];
    if (leftRow.length !== rightRow.length) return false;
    for (let columnIndex = 0; columnIndex < leftRow.length; columnIndex += 1) {
      if ((leftRow[columnIndex] ?? "") !== (rightRow[columnIndex] ?? "")) {
        return false;
      }
    }
  }
  return true;
}

function rowPatchCompatible(existing: string[], patch: string[]): boolean {
  const width = Math.max(existing.length, patch.length);
  for (let columnIndex = 0; columnIndex < width; columnIndex += 1) {
    const next = (patch[columnIndex] ?? "").trim();
    const prev = (existing[columnIndex] ?? "").trim();
    if (!next || !prev) continue;
    if (next !== prev) return false;
  }
  return true;
}

function resolveSelectionRowIndex(
  rawRowIndex: number,
  values: string[][],
  selectionText: string,
  cellText: string,
): { rowIndex: number; adjusted: boolean } {
  const rows = values.length;
  if (rows > 0 && rawRowIndex >= 0 && rawRowIndex < rows) {
    return { rowIndex: rawRowIndex, adjusted: false };
  }

  const bySelection = findTableRowIndexBySelection(values, selectionText);
  if (bySelection !== null) {
    return { rowIndex: bySelection, adjusted: true };
  }

  const byCell = findTableRowIndexByCellText(values, cellText);
  if (byCell !== null) {
    return { rowIndex: byCell, adjusted: true };
  }

  if (rows > 0) {
    const clamped = Math.max(0, Math.min(rawRowIndex, rows - 1));
    return { rowIndex: clamped, adjusted: rawRowIndex !== clamped };
  }

  return { rowIndex: 0, adjusted: rawRowIndex !== 0 };
}

interface ParentTableIndexResolution {
  index: number;
  method: TableIndexResolution;
}

async function resolveParentTableIndex(
  context: Word.RequestContext,
  tables: Word.TableCollection,
  parentTable: Word.Table,
  parentValues: string[][],
  rows: number,
  columns: number,
  rowIndex: number,
  rowValues: string[],
): Promise<ParentTableIndexResolution> {
  for (let index = 0; index < tables.items.length; index += 1) {
    if (tables.items[index] === parentTable) {
      return { index, method: "reference" };
    }
  }

  for (let index = 0; index < tables.items.length; index += 1) {
    tables.items[index].load(["values", "rowCount"]);
  }
  await context.sync();

  for (let index = 0; index < tables.items.length; index += 1) {
    const rawValues = (tables.items[index].values as string[][]) ?? [];
    const dimensions = getTableWriteDimensions(rawValues, tables.items[index].rowCount);
    if (dimensions.rows !== rows || dimensions.columns !== columns) continue;
    const values = normalizeTableValues(rawValues, dimensions.rows, dimensions.columns);
    if (tableValuesEqual(values, parentValues)) {
      return { index, method: "values_match" };
    }
  }

  const rowMatchCandidates: number[] = [];
  for (let index = 0; index < tables.items.length; index += 1) {
    const rawValues = (tables.items[index].values as string[][]) ?? [];
    const dimensions = getTableWriteDimensions(rawValues, tables.items[index].rowCount);
    if (rowIndex < 0 || rowIndex >= dimensions.rows) continue;
    const values = normalizeTableValues(rawValues, dimensions.rows, dimensions.columns);
    if (tableRowValuesMatch(values[rowIndex], rowValues)) {
      rowMatchCandidates.push(index);
    }
  }
  if (rowMatchCandidates.length === 1) {
    return { index: rowMatchCandidates[0], method: "row_match" };
  }

  const dimensionCandidates: number[] = [];
  for (let index = 0; index < tables.items.length; index += 1) {
    const rawValues = (tables.items[index].values as string[][]) ?? [];
    const dimensions = getTableWriteDimensions(rawValues, tables.items[index].rowCount);
    if (dimensions.rows === rows && dimensions.columns === columns) {
      dimensionCandidates.push(index);
    }
  }
  if (dimensionCandidates.length === 1) {
    return { index: dimensionCandidates[0], method: "dimensions_match" };
  }

  let bestFit = -1;
  for (let index = 0; index < tables.items.length; index += 1) {
    const rawValues = (tables.items[index].values as string[][]) ?? [];
    const dimensions = getTableWriteDimensions(rawValues, tables.items[index].rowCount);
    if (rowIndex >= 0 && rowIndex < dimensions.rows) {
      bestFit = index;
      break;
    }
  }
  if (bestFit >= 0) {
    return { index: bestFit, method: "row_match" };
  }

  return { index: 0, method: "dimensions_match" };
}

interface TableRowSpanResolution {
  startRow: number;
  endRow: number;
  startColumn: number | null;
  endColumn: number | null;
  cellText: string;
  adjusted: boolean;
}

async function loadTableCellColumnIndex(
  context: Word.RequestContext,
  cell: Word.TableCell,
  rowValues: string[],
  selectionText: string,
  cellText: string,
): Promise<number | null> {
  try {
    cell.load("cellIndex");
    await context.sync();
    return cell.cellIndex;
  } catch (_cellIndexError) {
    return inferTableColumnIndex(rowValues, selectionText, cellText);
  }
}

async function resolveTableRowSpanFromRange(
  context: Word.RequestContext,
  range: Word.Range,
  values: string[][],
  selectionText: string,
): Promise<TableRowSpanResolution> {
  const rangeStart = range.getRange(Word.RangeLocation.start);
  const rangeEnd = range.getRange(Word.RangeLocation.end);
  const startCell = rangeStart.parentTableCellOrNullObject;
  const endCell = rangeEnd.parentTableCellOrNullObject;
  startCell.load(["isNullObject", "rowIndex", "value"]);
  endCell.load(["isNullObject", "rowIndex", "value"]);
  await context.sync();

  if (!startCell.isNullObject && !endCell.isNullObject) {
    let startRow = startCell.rowIndex;
    let endRow = endCell.rowIndex;
    const startCellText = startCell.value ?? "";
    const endCellText = endCell.value ?? "";

    if (startRow > endRow) {
      const swapRow = startRow;
      startRow = endRow;
      endRow = swapRow;
    }

    const startResolution = resolveSelectionRowIndex(
      startRow,
      values,
      selectionText,
      startCellText,
    );
    const endResolution = resolveSelectionRowIndex(endRow, values, selectionText, endCellText);

    let resolvedStartRow = startResolution.rowIndex;
    let resolvedEndRow = endResolution.rowIndex;
    if (resolvedStartRow > resolvedEndRow) {
      const swapResolved = resolvedStartRow;
      resolvedStartRow = resolvedEndRow;
      resolvedEndRow = swapResolved;
    }

    let startColumn = await loadTableCellColumnIndex(
      context,
      startCell,
      values[resolvedStartRow] ?? [],
      selectionText,
      startCellText,
    );
    let endColumn = await loadTableCellColumnIndex(
      context,
      endCell,
      values[resolvedEndRow] ?? [],
      selectionText,
      endCellText,
    );

    if (startColumn !== null && endColumn !== null && startColumn > endColumn) {
      const swapColumn = startColumn;
      startColumn = endColumn;
      endColumn = swapColumn;
    }

    return {
      startRow: resolvedStartRow,
      endRow: resolvedEndRow,
      startColumn,
      endColumn,
      cellText: startCellText,
      adjusted: startResolution.adjusted || endResolution.adjusted,
    };
  }

  const parentCell = range.parentTableCellOrNullObject;
  parentCell.load(["isNullObject", "rowIndex", "value"]);
  await context.sync();

  let rawRowIndex = 0;
  let cellText = "";
  if (!parentCell.isNullObject) {
    rawRowIndex = parentCell.rowIndex;
    cellText = parentCell.value ?? "";
  } else {
    const matchedRow = findTableRowIndexBySelection(values, selectionText);
    if (matchedRow !== null) {
      rawRowIndex = matchedRow;
    }
  }

  const rowResolution = resolveSelectionRowIndex(rawRowIndex, values, selectionText, cellText);
  const columnIndex = !parentCell.isNullObject
    ? await loadTableCellColumnIndex(
        context,
        parentCell,
        values[rowResolution.rowIndex] ?? [],
        selectionText,
        cellText,
      )
    : inferTableColumnIndex(values[rowResolution.rowIndex] ?? [], selectionText, cellText);

  return {
    startRow: rowResolution.rowIndex,
    endRow: rowResolution.rowIndex,
    startColumn: columnIndex,
    endColumn: columnIndex,
    cellText,
    adjusted: rowResolution.adjusted,
  };
}

function clampTableRowIndex(rowIndex: number, rowCount: number): number {
  if (rowCount < 1) return 0;
  return Math.max(0, Math.min(rowIndex, rowCount - 1));
}

async function resolveTableRowSpanLight(
  context: Word.RequestContext,
  range: Word.Range,
  rowCount: number,
): Promise<TableRowSpanResolution> {
  const rangeStart = range.getRange(Word.RangeLocation.start);
  const rangeEnd = range.getRange(Word.RangeLocation.end);
  const startCell = rangeStart.parentTableCellOrNullObject;
  const endCell = rangeEnd.parentTableCellOrNullObject;
  startCell.load(["isNullObject", "rowIndex", "value"]);
  endCell.load(["isNullObject", "rowIndex", "value"]);
  await context.sync();

  if (!startCell.isNullObject && !endCell.isNullObject) {
    let startRow = clampTableRowIndex(startCell.rowIndex, rowCount);
    let endRow = clampTableRowIndex(endCell.rowIndex, rowCount);
    if (startRow > endRow) {
      const swapRow = startRow;
      startRow = endRow;
      endRow = swapRow;
    }

    const startCellText = startCell.value ?? "";
    let startColumn = await loadTableCellColumnIndex(
      context,
      startCell,
      [],
      "",
      startCellText,
    );
    let endColumn = await loadTableCellColumnIndex(
      context,
      endCell,
      [],
      "",
      endCell.value ?? "",
    );

    if (startColumn !== null && endColumn !== null && startColumn > endColumn) {
      const swapColumn = startColumn;
      startColumn = endColumn;
      endColumn = swapColumn;
    }

    return {
      startRow,
      endRow,
      startColumn,
      endColumn,
      cellText: startCellText,
      adjusted: false,
    };
  }

  const parentCell = range.parentTableCellOrNullObject;
  parentCell.load(["isNullObject", "rowIndex", "value"]);
  await context.sync();

  const resolvedRow = parentCell.isNullObject
    ? 0
    : clampTableRowIndex(parentCell.rowIndex, rowCount);
  const cellText = parentCell.isNullObject ? "" : parentCell.value ?? "";
  const columnIndex = !parentCell.isNullObject
    ? await loadTableCellColumnIndex(context, parentCell, [], "", cellText)
    : null;

  return {
    startRow: resolvedRow,
    endRow: resolvedRow,
    startColumn: columnIndex,
    endColumn: columnIndex,
    cellText,
    adjusted: false,
  };
}

function collectSelectedRowValues(values: string[][], startRow: number, endRow: number): string[][] {
  const rows: string[][] = [];
  for (let rowIndex = startRow; rowIndex <= endRow; rowIndex += 1) {
    rows.push(values[rowIndex] ? values[rowIndex].slice() : []);
  }
  return rows;
}

async function resolveTableDimensionsLight(
  table: Word.Table,
  context: Word.RequestContext,
): Promise<{ rows: number; columns: number }> {
  table.load(["values", "rowCount", "isUniform"]);
  await context.sync();

  const rawValues = (table.values as string[][]) ?? [];
  const rows = table.rowCount > 0 ? table.rowCount : rawValues.length;
  let columns = rows > 0 ? maxColumnCountFromValues(rawValues, Math.min(rows, 1)) : 0;

  if (table.isUniform === false && rows > 0) {
    const probed = await probeTableColumnCount(table, 1, context);
    columns = Math.max(columns, probed);
  }

  return { rows, columns };
}

async function readTableRowValues(
  table: Word.Table,
  context: Word.RequestContext,
  rowIndex: number,
  columns: number,
  isUniform: boolean,
): Promise<string[]> {
  if (columns < 1) return [];

  if (isUniform !== false) {
    table.load("values");
    await context.sync();
    const rawValues = (table.values as string[][]) ?? [];
    const grid = normalizeTableValues(rawValues, rawValues.length, columns);
    return grid[rowIndex] ? grid[rowIndex].slice() : new Array(columns).fill("");
  }

  const cells: Word.TableCell[] = [];
  for (let colIndex = 0; colIndex < columns; colIndex += 1) {
    const cell = table.getCellOrNullObject(rowIndex, colIndex);
    cells.push(cell);
    cell.load(["isNullObject", "value"]);
  }
  await context.sync();

  const line: string[] = [];
  for (let colIndex = 0; colIndex < columns; colIndex += 1) {
    line.push(!cells[colIndex].isNullObject ? cells[colIndex].value ?? "" : "");
  }
  return line;
}

async function resolveParentTableIndexLight(
  context: Word.RequestContext,
  tables: Word.TableCollection,
  parentTable: Word.Table,
  rows: number,
  columns: number,
): Promise<ParentTableIndexResolution> {
  for (let index = 0; index < tables.items.length; index += 1) {
    if (tables.items[index] === parentTable) {
      return { index, method: "reference" };
    }
  }

  for (let index = 0; index < tables.items.length; index += 1) {
    tables.items[index].load(["values", "rowCount"]);
  }
  await context.sync();

  const dimensionCandidates: number[] = [];
  for (let index = 0; index < tables.items.length; index += 1) {
    const dimensions = getTableWriteDimensions(
      (tables.items[index].values as string[][]) ?? [],
      tables.items[index].rowCount,
    );
    if (dimensions.rows === rows && dimensions.columns === columns) {
      dimensionCandidates.push(index);
    }
  }
  if (dimensionCandidates.length === 1) {
    return { index: dimensionCandidates[0], method: "dimensions_match" };
  }

  return { index: 0, method: "dimensions_match" };
}

async function readTableContextFromRange(
  context: Word.RequestContext,
  range: Word.Range,
  selectionText: string,
  selectionSource: TableSelectionSource,
  bookmark?: string,
  includeTableValues = false,
): Promise<TableSelectionContext | null> {
  const parentTable = range.parentTableOrNullObject;
  parentTable.load("isNullObject");
  await context.sync();

  if (parentTable.isNullObject) {
    return null;
  }

  const tables = context.document.body.tables;
  tables.load("items");
  parentTable.load(["values", "rowCount", "isUniform"]);
  await context.sync();

  const isUniform = parentTable.isUniform !== false;
  const dimensions = includeTableValues
    ? await resolveTableDimensions(parentTable, context)
    : await resolveTableDimensionsLight(parentTable, context);

  let values: string[][] | undefined;
  if (includeTableValues) {
    values = await readTableValuesGrid(parentTable, context, dimensions.rows, dimensions.columns);
  }

  const rowSpan = includeTableValues
    ? await resolveTableRowSpanFromRange(context, range, values as string[][], selectionText)
    : await resolveTableRowSpanLight(context, range, dimensions.rows);

  const rowIndex = rowSpan.startRow;
  const rowIndexEnd = rowSpan.endRow;
  const selectedRowCount = rowIndexEnd - rowIndex + 1;

  let rowValues: string[];
  let selectedRowValues: string[][] | undefined;

  if (includeTableValues && values) {
    rowValues = values[rowIndex] ? values[rowIndex].slice() : [];
    selectedRowValues =
      selectedRowCount > 1 ? collectSelectedRowValues(values, rowIndex, rowIndexEnd) : undefined;
  } else {
    rowValues = await readTableRowValues(
      parentTable,
      context,
      rowIndex,
      dimensions.columns,
      isUniform,
    );
  }

  const tableResolution = includeTableValues
    ? await resolveParentTableIndex(
        context,
        tables,
        parentTable,
        values as string[][],
        dimensions.rows,
        dimensions.columns,
        rowIndex,
        rowValues,
      )
    : await resolveParentTableIndexLight(
        context,
        tables,
        parentTable,
        dimensions.rows,
        dimensions.columns,
      );

  return {
    tableIndex: tableResolution.index,
    rowIndex,
    rowIndexEnd: rowIndexEnd !== rowIndex ? rowIndexEnd : undefined,
    columnIndex: rowSpan.startColumn,
    columnIndexEnd:
      rowSpan.endColumn !== null &&
      rowSpan.startColumn !== null &&
      rowSpan.endColumn !== rowSpan.startColumn
        ? rowSpan.endColumn
        : undefined,
    selectedRowCount: selectedRowCount > 1 ? selectedRowCount : undefined,
    rows: dimensions.rows,
    columns: dimensions.columns,
    isUniform,
    selectionText,
    cellText: rowSpan.cellText,
    rowValues,
    ...(values ? { tableValues: values } : {}),
    selectedRowValues,
    tableIndexResolution: tableResolution.method,
    rowIndexAdjusted: rowSpan.adjusted,
    selectionSource,
    bookmark,
  };
}

export async function readTableSelectionContext(
  sourceOrOptions: TableSelectionReadSource | TableSelectionReadOptions = "pinned_or_live",
): Promise<TableSelectionContext | null> {
  const options = normalizeTableSelectionReadOptions(sourceOrOptions);
  const source = options.source;
  const includeTableValues = options.includeTableValues;

  assertWordAvailable();
  try {
    return await Word.run(async (context) => {
      if (source !== "live") {
        const pinnedRange = context.document.getBookmarkRangeOrNullObject(USER_SELECTION_BOOKMARK);
        pinnedRange.load(["isNullObject", "text"]);
        await context.sync();
        if (!pinnedRange.isNullObject) {
          const pinnedContext = await readTableContextFromRange(
            context,
            pinnedRange,
            pinnedRange.text ?? "",
            "pinned",
            USER_SELECTION_BOOKMARK,
            includeTableValues,
          );
          if (pinnedContext) {
            return pinnedContext;
          }
        }
        if (source === "pinned") {
          return null;
        }
      }

      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return readTableContextFromRange(
        context,
        selection,
        selection.text ?? "",
        "live",
        undefined,
        includeTableValues,
      );
    });
  } catch (error) {
    wrapWordError(error, "Failed to read table selection context.");
  }
}

export async function readPinnedTableSelectionContext(): Promise<TableSelectionContext | null> {
  return readTableSelectionContext({ source: "pinned", includeTableValues: false });
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

const MAX_TABLE_PROBE_COLUMNS = 20;
export const MAX_TABLE_ROW_INSERT = 20;

export interface TableRowInsertPlan {
  atRow: number;
  count: number;
  mode: "before" | "after";
  /** When set, seed text on insert (insert-only mode). */
  rowValues?: string[][];
}

function maxColumnCountFromValues(values: string[][], rows: number): number {
  let columns = 0;
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    columns = Math.max(columns, values[rowIndex]?.length ?? 0);
  }
  return columns;
}

/** Dimensions Word accepts for table.values writes (merged cells may return ragged rows on read). */
function getTableWriteDimensions(
  values: string[][],
  rowCount?: number,
): { rows: number; columns: number } {
  const rows = rowCount !== undefined && rowCount > 0 ? rowCount : values.length;
  const columns = rows > 0 ? maxColumnCountFromValues(values, rows) : 0;
  return { rows, columns };
}

async function probeTableColumnCount(
  table: Word.Table,
  rowCount: number,
  context: Word.RequestContext,
): Promise<number> {
  if (rowCount < 1) return 0;

  const probes: { rowIndex: number; colIndex: number; cell: Word.TableCell }[] = [];
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    for (let colIndex = 0; colIndex < MAX_TABLE_PROBE_COLUMNS; colIndex += 1) {
      const cell = table.getCellOrNullObject(rowIndex, colIndex);
      probes.push({ rowIndex, colIndex, cell });
      cell.load("isNullObject");
    }
  }
  await context.sync();

  let maxColumns = 0;
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    let rowColumns = 0;
    for (let colIndex = 0; colIndex < MAX_TABLE_PROBE_COLUMNS; colIndex += 1) {
      const probe = probes[rowIndex * MAX_TABLE_PROBE_COLUMNS + colIndex];
      if (!probe.cell.isNullObject) {
        rowColumns = colIndex + 1;
      }
    }
    maxColumns = Math.max(maxColumns, rowColumns);
  }
  return maxColumns;
}

async function resolveTableDimensions(
  table: Word.Table,
  context: Word.RequestContext,
): Promise<{ rows: number; columns: number }> {
  table.load(["values", "rowCount", "isUniform"]);
  await context.sync();

  const rawValues = (table.values as string[][]) ?? [];
  const rows = table.rowCount > 0 ? table.rowCount : rawValues.length;
  let columns = maxColumnCountFromValues(rawValues, rows);

  if (table.isUniform === false && rows > 0) {
    const probed = await probeTableColumnCount(table, rows, context);
    columns = Math.max(columns, probed);
  }

  return { rows, columns };
}

async function readTableValuesGrid(
  table: Word.Table,
  context: Word.RequestContext,
  rows: number,
  columns: number,
): Promise<string[][]> {
  table.load(["values", "isUniform"]);
  await context.sync();

  const rawValues = (table.values as string[][]) ?? [];

  if (table.isUniform !== false) {
    return normalizeTableValues(rawValues, rows, columns);
  }

  const cells: Word.TableCell[] = [];
  const coords: { rowIndex: number; colIndex: number }[] = [];
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    for (let colIndex = 0; colIndex < columns; colIndex += 1) {
      const cell = table.getCellOrNullObject(rowIndex, colIndex);
      cells.push(cell);
      coords.push({ rowIndex, colIndex });
      cell.load(["isNullObject", "value"]);
    }
  }
  await context.sync();

  const grid = normalizeTableValues(rawValues, rows, columns);
  for (let index = 0; index < cells.length; index += 1) {
    if (!cells[index].isNullObject) {
      grid[coords[index].rowIndex][coords[index].colIndex] = cells[index].value ?? "";
    }
  }
  return grid;
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
        const table = tables.items[index];
        const dimensions = await resolveTableDimensions(table, context);
        const values = await readTableValuesGrid(
          table,
          context,
          dimensions.rows,
          dimensions.columns,
        );
        infos.push({
          index,
          rows: dimensions.rows,
          columns: dimensions.columns,
          values,
          preview: formatTablePreview(values),
          isUniform: table.isUniform,
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
        const dimensions = await resolveTableDimensions(parentTable, context);
        const values = await readTableValuesGrid(
          parentTable,
          context,
          dimensions.rows,
          dimensions.columns,
        );
        const parentCell = selection.parentTableCellOrNullObject;
        parentCell.load(["isNullObject", "rowIndex", "value"]);
        selection.load("text");
        await context.sync();

        let rowIndex = 0;
        let rowValues: string[] = [];
        if (!parentCell.isNullObject) {
          const rowResolution = resolveSelectionRowIndex(
            parentCell.rowIndex,
            values,
            selection.text ?? "",
            parentCell.value ?? "",
          );
          rowIndex = rowResolution.rowIndex;
        }
        rowValues = values[rowIndex] ? values[rowIndex].slice() : [];

        const resolved = await resolveParentTableIndex(
          context,
          tables,
          parentTable,
          values,
          dimensions.rows,
          dimensions.columns,
          rowIndex,
          rowValues,
        );
        return resolved.index;
      }

      return 0;
    });
  } catch (error) {
    wrapWordError(error, "Failed to resolve table.");
  }
}

export async function findTableIndexForRowPatch(
  preferredIndex: number,
  startRow: number,
  patch: string[][],
): Promise<number | null> {
  assertWordAvailable();
  const tables = await listDocumentTables(20);
  if (!tables.length) return null;

  const preferred = tables.find((table) => table.index === preferredIndex);
  if (
    preferred &&
    startRow >= 0 &&
    startRow + patch.length <= preferred.rows &&
    rowPatchCompatible(preferred.values[startRow] ?? [], patch[0] ?? [])
  ) {
    return preferred.index;
  }

  const selection = await readTableSelectionContext();
  if (selection) {
    const selected = tables.find((table) => table.index === selection.tableIndex);
    if (
      selected &&
      startRow >= 0 &&
      startRow + patch.length <= selected.rows &&
      rowPatchCompatible(selected.values[startRow] ?? [], patch[0] ?? [])
    ) {
      return selected.index;
    }
  }

  const matches: number[] = [];
  for (let index = 0; index < tables.length; index += 1) {
    const table = tables[index];
    if (startRow < 0 || startRow + patch.length > table.rows) continue;
    if (!rowPatchCompatible(table.values[startRow] ?? [], patch[0] ?? [])) continue;
    matches.push(table.index);
  }

  if (matches.length === 1) return matches[0];
  if (selection && matches.indexOf(selection.tableIndex) >= 0) {
    return selection.tableIndex;
  }
  return null;
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
      const dimensions = await resolveTableDimensions(table, context);
      const values = await readTableValuesGrid(
        table,
        context,
        dimensions.rows,
        dimensions.columns,
      );
      table.load("isUniform");
      await context.sync();

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

async function insertTableRowsPlan(
  table: Word.Table,
  context: Word.RequestContext,
  plan: TableRowInsertPlan,
  columns: number,
): Promise<void> {
  const count = Math.min(MAX_TABLE_ROW_INSERT, Math.max(1, Math.floor(plan.count)));
  if (count < 1) return;

  table.load("rowCount");
  table.rows.load("items");
  await context.sync();

  const rowCount = table.rowCount;
  const valueGrid = plan.rowValues?.length
    ? normalizeCellGrid(plan.rowValues.slice(0, count), count, columns)
    : null;

  if (plan.mode === "after" && plan.atRow >= rowCount - 1) {
    table.addRows(Word.InsertLocation.end, count, valueGrid ?? undefined);
    await context.sync();
    return;
  }

  if (rowCount < 1) {
    throw new WordOperationError("Cannot insert rows into an empty table.");
  }

  const anchorIndex = Math.min(Math.max(0, plan.atRow), rowCount - 1);
  const insertLocation =
    plan.mode === "before" ? Word.InsertLocation.before : Word.InsertLocation.after;
  table.rows.items[anchorIndex].insertRows(insertLocation, count, valueGrid ?? undefined);
  await context.sync();
}

export async function deleteTableRowsAtIndex(
  tableIndex: number,
  rowIndex: number,
  rowCount: number,
): Promise<void> {
  assertWordAvailable();
  const count = Math.max(1, Math.floor(rowCount));
  const index = Math.max(0, Math.floor(rowIndex));

  try {
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tableIndex < 0 || tableIndex >= tables.items.length) {
        throw new WordOperationError(`Cannot delete rows: table index ${tableIndex} is out of range.`);
      }

      tables.items[tableIndex].deleteRows(index, count);
      await context.sync();
    });
  } catch (error) {
    wrapWordError(error, "Failed to delete table rows.");
  }
}

export async function updateTableAtIndex(
  tableIndex: number,
  rows: number,
  columns: number,
  cellValues: string[][],
  rowInsert?: TableRowInsertPlan,
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

      if (rowInsert?.count) {
        await insertTableRowsPlan(table, context, rowInsert, columns);
      }

      const dimensions = await resolveTableDimensions(table, context);

      if (dimensions.rows !== rows || dimensions.columns !== columns) {
        throw new WordOperationError(
          `Table ${tableIndex} is ${dimensions.rows}x${dimensions.columns}. Pass matching rows and columns.`,
        );
      }

      const previous = await readTableValuesGrid(
        table,
        context,
        dimensions.rows,
        dimensions.columns,
      );
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
      const dimensions = await resolveTableDimensions(table, context);
      const previous = await readTableValuesGrid(
        table,
        context,
        dimensions.rows,
        dimensions.columns,
      );
      const grid = normalizeTableValues(cloneCellGrid(values), dimensions.rows, dimensions.columns);
      await writeTableValues(table, grid, previous, context);
    });
  } catch (error) {
    wrapWordError(error, "Failed to restore table.");
  }
}