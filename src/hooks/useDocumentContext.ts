import { useCallback, useEffect, useState } from "react";
import type { ContextMode, DocumentContext } from "../types/context";
import { getDocumentContext } from "../word/context";

const EMPTY_PREVIEW: DocumentContext = {
  mode: "none",
  text: "",
  tokenEstimate: 0,
  truncated: false,
  empty: true,
};

export function useDocumentContext(mode: ContextMode) {
  const [context, setContext] = useState<DocumentContext>(EMPTY_PREVIEW);
  const [isLoading, setIsLoading] = useState(false);

  const refresh = useCallback(async () => {
    if (mode === "none") {
      setContext({ ...EMPTY_PREVIEW, mode: "none" });
      return;
    }

    setIsLoading(true);
    try {
      const next = await getDocumentContext(mode);
      setContext(next);
    } finally {
      setIsLoading(false);
    }
  }, [mode]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  return {
    context,
    isLoading,
    refresh,
  };
}