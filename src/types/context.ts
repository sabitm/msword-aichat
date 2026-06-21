export type ContextMode = "selection" | "outline" | "none";

export interface TableSelectionContext {
  tableIndex: number;
  rowIndex: number;
  columnIndex: number | null;
  rows: number;
  columns: number;
  isUniform: boolean;
  selectionText: string;
  cellText: string;
  rowValues: string[];
}

export interface DocumentContext {
  mode: ContextMode;
  text: string;
  tokenEstimate: number;
  truncated: boolean;
  empty: boolean;
  error?: string;
  tableSelection?: TableSelectionContext | null;
}

export const CONTEXT_MODE_LABELS: Record<ContextMode, string> = {
  selection: "Selection",
  outline: "Outline",
  none: "None",
};

export const MAX_CONTEXT_CHARS = 12_000;