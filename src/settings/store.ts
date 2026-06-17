import { create } from "zustand";
import type { InteractionMode } from "../types/agent";
import type { ProviderConfig, ProviderKind } from "../types/llm";
import {
  API_KEY_STORAGE_KEY,
  DEFAULT_PREFERENCES,
  DEFAULT_PROVIDER_CONFIG,
  PREFERENCES_STORAGE_KEY,
  SETTINGS_STORAGE_KEY,
  type AppPreferences,
} from "./defaults";

interface SettingsState {
  config: ProviderConfig;
  preferences: AppPreferences;
  isConfigured: boolean;
  load: () => void;
  update: (patch: Partial<ProviderConfig>) => void;
  updatePreferences: (patch: Partial<AppPreferences>) => void;
  setKind: (kind: ProviderKind) => void;
  setInteractionMode: (mode: InteractionMode) => void;
  completeOnboarding: () => void;
  save: () => void;
  getConfig: () => ProviderConfig;
  getPreferences: () => AppPreferences;
}

function readStoredConfig(): ProviderConfig {
  const defaults = { ...DEFAULT_PROVIDER_CONFIG };
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    const apiKey = localStorage.getItem(API_KEY_STORAGE_KEY) ?? "";
    if (!raw) {
      return { ...defaults, apiKey };
    }
    const parsed = JSON.parse(raw) as Partial<ProviderConfig>;
    return {
      ...defaults,
      ...parsed,
      apiKey,
    };
  } catch {
    return defaults;
  }
}

function readStoredPreferences(): AppPreferences {
  try {
    const raw = localStorage.getItem(PREFERENCES_STORAGE_KEY);
    if (!raw) return { ...DEFAULT_PREFERENCES };
    const parsed = JSON.parse(raw) as Partial<AppPreferences>;
    return { ...DEFAULT_PREFERENCES, ...parsed };
  } catch {
    return { ...DEFAULT_PREFERENCES };
  }
}

function isConfigReady(config: ProviderConfig): boolean {
  return Boolean(config.baseUrl.trim() && config.apiKey.trim() && config.model.trim());
}

export const useSettingsStore = create<SettingsState>((set, get) => ({
  config: DEFAULT_PROVIDER_CONFIG,
  preferences: DEFAULT_PREFERENCES,
  isConfigured: false,

  load: () => {
    const config = readStoredConfig();
    const preferences = readStoredPreferences();
    set({ config, preferences, isConfigured: isConfigReady(config) });
  },

  update: (patch) => {
    set((state) => {
      const config = { ...state.config, ...patch };
      return { config, isConfigured: isConfigReady(config) };
    });
  },

  updatePreferences: (patch) => {
    set((state) => ({
      preferences: { ...state.preferences, ...patch },
    }));
  },

  setKind: (kind) => {
    const patch: Partial<ProviderConfig> = { kind };
    if (kind === "openai") {
      patch.baseUrl = "https://api.openai.com/v1";
    } else {
      patch.baseUrl = "https://api.anthropic.com/v1";
    }
    get().update(patch);
  },

  setInteractionMode: (mode) => {
    get().updatePreferences({ interactionMode: mode });
  },

  completeOnboarding: () => {
    get().updatePreferences({ onboardingCompleted: true });
    get().save();
  },

  save: () => {
    const { config, preferences } = get();
    const { apiKey, ...rest } = config;
    localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(rest));
    localStorage.setItem(API_KEY_STORAGE_KEY, apiKey);
    localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify(preferences));
    set({ isConfigured: isConfigReady(config) });
  },

  getConfig: () => get().config,
  getPreferences: () => get().preferences,
}));