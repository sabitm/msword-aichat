export type ContextMode = "selection" | "outline" | "none";

export interface DocumentContext {
  mode: ContextMode;
  text: string;
  tokenEstimate: number;
  truncated: boolean;
  empty: boolean;
  error?: string;
}

export const CONTEXT_MODE_LABELS: Record<ContextMode, string> = {
  selection: "Selection",
  outline: "Outline",
  none: "None",
};

export const MAX_CONTEXT_CHARS = 12_000;