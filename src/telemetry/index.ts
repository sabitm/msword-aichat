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
  const { telemetryEnabled } = useSettingsStore.getState().getPreferences();
  if (!telemetryEnabled) return;

  if (import.meta.env.DEV) {
    console.debug("[telemetry]", name, properties ?? {});
  }
}