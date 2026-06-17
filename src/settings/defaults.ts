import type { InteractionMode } from "../types/agent";
import type { ProviderConfig } from "../types/llm";

export const DEFAULT_PROVIDER_CONFIG: ProviderConfig = {
  kind: "openai",
  baseUrl: "https://api.openai.com/v1",
  apiKey: "",
  model: "gpt-4o-mini",
  maxTokens: 2048,
  temperature: 0.7,
};

export interface AppPreferences {
  interactionMode: InteractionMode;
  autoApplyEdits: boolean;
}

export const DEFAULT_PREFERENCES: AppPreferences = {
  interactionMode: "agent",
  autoApplyEdits: false,
};

export const SETTINGS_STORAGE_KEY = "msword-aichat:provider-config";
export const API_KEY_STORAGE_KEY = "msword-aichat:api-key";
export const PREFERENCES_STORAGE_KEY = "msword-aichat:preferences";