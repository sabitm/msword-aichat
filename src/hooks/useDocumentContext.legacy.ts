import * as React from "react";
import type { ContextMode, DocumentContext } from "../types/context";
import { getDocumentContext } from "../word/context";

var EMPTY_PREVIEW: DocumentContext = {
  mode: "none",
  text: "",
  tokenEstimate: 0,
  truncated: false,
  empty: true,
};

export function useDocumentContext(mode: ContextMode): {
  context: DocumentContext;
  isLoading: boolean;
  refresh: () => void;
} {
  var _a = React.useState<DocumentContext>(EMPTY_PREVIEW);
  var context = _a[0];
  var setContext = _a[1];

  var _b = React.useState(false);
  var isLoading = _b[0];
  var setIsLoading = _b[1];

  var refresh = React.useCallback(function () {
    if (mode === "none") {
      setContext(Object.assign({}, EMPTY_PREVIEW, { mode: "none" }));
      return;
    }

    setIsLoading(true);
    getDocumentContext(mode)
      .then(function (next) {
        setContext(next);
      })
      .then(function () {
        setIsLoading(false);
      });
  }, [mode]);

  React.useEffect(function () {
    refresh();
  }, [refresh]);

  return {
    context: context,
    isLoading: isLoading,
    refresh: refresh,
  };
}