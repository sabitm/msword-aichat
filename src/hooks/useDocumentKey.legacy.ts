import * as React from "react";
import { getDocumentKey } from "../word/document-key.legacy";

export function useDocumentKey(): { docKey: string; refresh: () => void } {
  var _a = React.useState("loading");
  var docKey = _a[0];
  var setDocKey = _a[1];

  var refresh = React.useCallback(function () {
    getDocumentKey().then(function (key) {
      setDocKey(key);
    });
  }, []);

  React.useEffect(function () {
    refresh();
    window.addEventListener("focus", refresh);
    return function () {
      window.removeEventListener("focus", refresh);
    };
  }, [refresh]);

  return { docKey: docKey, refresh: refresh };
}