import { settingsStore } from "../settings/store.legacy";

export type TelemetryEvent =
  | "onboarding_completed"
  | "connection_test"
  | "chat_message_sent"
  | "agent_run"
  | "edit_applied"
  | "error_occurred";

export function trackEvent(
  name: TelemetryEvent,
  properties?: Record<string, string | number | boolean>,
): void {
  var preferences = settingsStore.getPreferences();
  if (!preferences.telemetryEnabled) {
    return;
  }

  var payload = {
    event: name,
    timestamp: new Date().toISOString(),
    properties: properties || {},
  };

  if (process.env.NODE_ENV !== "production") {
    console.debug("[telemetry]", payload);
  }

  var endpoint = preferences.telemetryEndpoint.trim();
  if (!endpoint) {
    return;
  }

  fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  }).catch(function () {
    // Telemetry must never break the add-in.
  });
}