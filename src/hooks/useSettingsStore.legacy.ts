import * as React from "react";
import { settingsStore, type SettingsState } from "../settings/store.legacy";

export function useSettingsStore(): SettingsState {
  var _a = React.useState(0);
  var setTick = _a[1];

  React.useEffect(function () {
    return settingsStore.subscribe(function () {
      setTick(function (tick) {
        return tick + 1;
      });
    });
  }, []);

  return settingsStore.getState();
}