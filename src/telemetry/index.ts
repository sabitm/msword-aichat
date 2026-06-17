import { useSettingsStore } from "../settings/store";

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
  const preferences = useSettingsStore.getState().getPreferences();
  if (!preferences.telemetryEnabled) return;

  const payload = {
    event: name,
    timestamp: new Date().toISOString(),
    properties: properties ?? {},
  };

  if (import.meta.env.DEV) {
    console.debug("[telemetry]", payload);
  }

  const endpoint = preferences.telemetryEndpoint.trim();
  if (!endpoint) return;

  void fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
    keepalive: true,
  }).catch(() => {
    // Telemetry must never break the add-in.
  });
}