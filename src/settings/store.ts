import { create } from "zustand";
import type { ProviderConfig, ProviderKind } from "../types/llm";
import {
  API_KEY_STORAGE_KEY,
  DEFAULT_PROVIDER_CONFIG,
  SETTINGS_STORAGE_KEY,
} from "./defaults";

interface SettingsState {
  config: ProviderConfig;
  isConfigured: boolean;
  load: () => void;
  update: (patch: Partial<ProviderConfig>) => void;
  setKind: (kind: ProviderKind) => void;
  save: () => void;
  getConfig: () => ProviderConfig;
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

function isConfigReady(config: ProviderConfig): boolean {
  return Boolean(config.baseUrl.trim() && config.apiKey.trim() && config.model.trim());
}

export const useSettingsStore = create<SettingsState>((set, get) => ({
  config: DEFAULT_PROVIDER_CONFIG,
  isConfigured: false,

  load: () => {
    const config = readStoredConfig();
    set({ config, isConfigured: isConfigReady(config) });
  },

  update: (patch) => {
    set((state) => {
      const config = { ...state.config, ...patch };
      return { config, isConfigured: isConfigReady(config) };
    });
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

  save: () => {
    const { config } = get();
    const { apiKey, ...rest } = config;
    localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(rest));
    localStorage.setItem(API_KEY_STORAGE_KEY, apiKey);
    set({ isConfigured: isConfigReady(config) });
  },

  getConfig: () => get().config,
}));