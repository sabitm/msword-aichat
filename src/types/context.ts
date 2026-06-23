export type ContextMode = "selection" | "outline" | "none";

export type TableIndexResolution = "reference" | "values_match" | "row_match" | "dimensions_match";

export type TableSelectionSource = "live" | "pinned";

export interface TableSelectionContext {
  tableIndex: number;
  /** 0-based first selected row (inclusive). */
  rowIndex: number;
  /** 0-based last selected row when a multi-row range is pinned or selected. */
  rowIndexEnd?: number;
  columnIndex: number | null;
  /** 0-based last selected column for multi-cell ranges. */
  columnIndexEnd?: number;
  /** Number of table rows covered by the selection (1 for a single cell). */
  selectedRowCount?: number;
  rows: number;
  columns: number;
  isUniform: boolean;
  selectionText: string;
  cellText: string;
  rowValues: string[];
  /** Full table cell grid when requested (get_selection); omitted on lightweight UI refresh. */
  tableValues?: string[][];
  /** Cell values for each row in a multi-row selection, when available. */
  selectedRowValues?: string[][];
  /** How table_index was resolved when parentTable is not a direct body.tables reference. */
  tableIndexResolution?: TableIndexResolution;
  /** True when row_index was adjusted because parentCell.rowIndex was out of range. */
  rowIndexAdjusted?: boolean;
  /** Whether coordinates came from the live selection or the pinned bookmark. */
  selectionSource?: TableSelectionSource;
  /** Bookmark name when selectionSource is pinned. */
  bookmark?: string;
}

export interface DocumentContext {
  mode: ContextMode;
  text: string;
  tokenEstimate: number;
  truncated: boolean;
  empty: boolean;
  error?: string;
  tableSelection?: TableSelectionContext | null;
  /** True when context was read from the pinned Sync bookmark. */
  selectionPinned?: boolean;
  selectionBookmark?: string;
}

export const CONTEXT_MODE_LABELS: Record<ContextMode, string> = {
  selection: "Selection",
  outline: "Outline",
  none: "None",
};

export const MAX_CONTEXT_CHARS = 12_000;