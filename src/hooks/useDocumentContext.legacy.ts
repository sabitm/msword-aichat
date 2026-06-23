import * as React from "react";
import type { ContextMode, DocumentContext } from "../types/context";
import { getDocumentContext } from "../word/context";
import { upsertUserSelectionBookmark } from "../word/ranges";

var EMPTY_PREVIEW: DocumentContext = {
  mode: "selection",
  text: "",
  tokenEstimate: 0,
  truncated: false,
  empty: true,
};

export function useDocumentContext(mode: ContextMode): {
  context: DocumentContext;
  isLoading: boolean;
  refresh: () => void;
  pinAndRefresh: () => void;
} {
  var _a = React.useState<DocumentContext>(EMPTY_PREVIEW);
  var context = _a[0];
  var setContext = _a[1];

  var _b = React.useState(false);
  var isLoading = _b[0];
  var setIsLoading = _b[1];

  var refresh = React.useCallback(function () {
    setIsLoading(true);
    getDocumentContext(mode)
      .then(function (next) {
        setContext(next);
      })
      .then(function () {
        setIsLoading(false);
      });
  }, [mode]);

  var pinAndRefresh = React.useCallback(function () {
    setIsLoading(true);
    upsertUserSelectionBookmark()
      .then(function (pinned) {
        return getDocumentContext(mode, {
          preferPinned: true,
          pinnedSelectionText: pinned.text,
        });
      })
      .then(function (next) {
        setContext(next);
      })
      .catch(function (error) {
        const message = error instanceof Error ? error.message : "Failed to pin selection";
        setContext(
          Object.assign({}, EMPTY_PREVIEW, {
            mode: mode,
            error: message,
          }),
        );
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
    pinAndRefresh: pinAndRefresh,
  };
}