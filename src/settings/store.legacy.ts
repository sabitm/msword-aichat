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

export interface SettingsState {
  config: ProviderConfig;
  preferences: AppPreferences;
  isConfigured: boolean;
}

type Listener = () => void;

function readStoredConfig(): ProviderConfig {
  var defaults = Object.assign({}, DEFAULT_PROVIDER_CONFIG);
  try {
    var raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    var apiKey = localStorage.getItem(API_KEY_STORAGE_KEY) || "";
    if (!raw) {
      return Object.assign({}, defaults, { apiKey: apiKey });
    }
    var parsed = JSON.parse(raw) as Partial<ProviderConfig>;
    return Object.assign({}, defaults, parsed, { apiKey: apiKey });
  } catch (_error) {
    return defaults;
  }
}

function readStoredPreferences(): AppPreferences {
  try {
    var raw = localStorage.getItem(PREFERENCES_STORAGE_KEY);
    if (!raw) {
      return Object.assign({}, DEFAULT_PREFERENCES);
    }
    var parsed = JSON.parse(raw) as Partial<AppPreferences>;
    return Object.assign({}, DEFAULT_PREFERENCES, parsed);
  } catch (_error) {
    return Object.assign({}, DEFAULT_PREFERENCES);
  }
}

function isConfigReady(config: ProviderConfig): boolean {
  return Boolean(config.baseUrl.trim() && config.apiKey.trim() && config.model.trim());
}

var state: SettingsState = {
  config: DEFAULT_PROVIDER_CONFIG,
  preferences: DEFAULT_PREFERENCES,
  isConfigured: false,
};

var listeners: Listener[] = [];

function notify(): void {
  for (var i = 0; i < listeners.length; i++) {
    listeners[i]();
  }
}

function setState(patch: Partial<SettingsState>): void {
  state = Object.assign({}, state, patch);
  notify();
}

export var settingsStore = {
  getState: function (): SettingsState {
    return state;
  },

  subscribe: function (listener: Listener): () => void {
    listeners.push(listener);
    return function () {
      var index = listeners.indexOf(listener);
      if (index >= 0) {
        listeners.splice(index, 1);
      }
    };
  },

  load: function (): void {
    var config = readStoredConfig();
    var preferences = readStoredPreferences();
    setState({
      config: config,
      preferences: preferences,
      isConfigured: isConfigReady(config),
    });
  },

  update: function (patch: Partial<ProviderConfig>): void {
    var config = Object.assign({}, state.config, patch);
    setState({
      config: config,
      isConfigured: isConfigReady(config),
    });
  },

  updatePreferences: function (patch: Partial<AppPreferences>): void {
    setState({
      preferences: Object.assign({}, state.preferences, patch),
    });
  },

  setKind: function (kind: ProviderKind): void {
    var patch: Partial<ProviderConfig> = { kind: kind };
    if (kind === "openai") {
      patch.baseUrl = "https://api.openai.com/v1";
    } else {
      patch.baseUrl = "https://api.anthropic.com/v1";
    }
    settingsStore.update(patch);
  },

  setInteractionMode: function (mode: InteractionMode): void {
    settingsStore.updatePreferences({ interactionMode: mode });
  },

  completeOnboarding: function (): void {
    settingsStore.updatePreferences({ onboardingCompleted: true });
    settingsStore.save();
  },

  save: function (): void {
    var config = state.config;
    var preferences = state.preferences;
    var rest = Object.assign({}, config);
    var apiKey = rest.apiKey;
    delete (rest as { apiKey?: string }).apiKey;
    localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(rest));
    localStorage.setItem(API_KEY_STORAGE_KEY, apiKey);
    localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify(preferences));
    setState({ isConfigured: isConfigReady(config) });
  },

  getConfig: function (): ProviderConfig {
    return state.config;
  },

  getPreferences: function (): AppPreferences {
    return state.preferences;
  },
};