import { useCallback, useEffect, useState } from "react";
import { getDocumentKey } from "../word/document-key";

export function useDocumentKey() {
  const [docKey, setDocKey] = useState("loading");

  const refresh = useCallback(async () => {
    setDocKey(await getDocumentKey());
  }, []);

  useEffect(() => {
    void refresh();
    window.addEventListener("focus", refresh);
    return () => window.removeEventListener("focus", refresh);
  }, [refresh]);

  return { docKey, refresh };
}