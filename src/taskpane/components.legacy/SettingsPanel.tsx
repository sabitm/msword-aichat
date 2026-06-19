import {
  Checkbox,
  DefaultButton,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Slider,
  Spinner,
  SpinnerSize,
  Text,
  TextField,
} from "@fluentui/react";
import * as React from "react";
import { useSettingsStore } from "../../hooks/useSettingsStore.legacy";
import { createProvider } from "../../llm/factory";
import { fetchModelList } from "../../llm/models";
import { settingsStore } from "../../settings/store.legacy";
import { trackEvent } from "../../telemetry/telemetry.legacy";
import type { PingResult, ProviderKind } from "../../types/llm";
import { IeSelect } from "./IeSelect";

var providerLabels: Record<ProviderKind, string> = {
  openai: "OpenAI-compatible",
  anthropic: "Anthropic-compatible",
};

var providerOptions = [
  { value: "openai", label: providerLabels.openai },
  { value: "anthropic", label: providerLabels.anthropic },
];

export function SettingsPanel(): React.ReactElement {
  var storeState = useSettingsStore();
  var config = storeState.config;
  var preferences = storeState.preferences;
  var isConfigured = storeState.isConfigured;

  var _a = React.useState<PingResult | null>(null);
  var pingResult = _a[0];
  var setPingResult = _a[1];

  var _b = React.useState(false);
  var isPinging = _b[0];
  var setIsPinging = _b[1];

  var _c = React.useState(false);
  var isFetchingModels = _c[0];
  var setIsFetchingModels = _c[1];

  var _d = React.useState<string | null>(null);
  var modelListMessage = _d[0];
  var setModelListMessage = _d[1];

  var _e = React.useState<string[]>([]);
  var availableModels = _e[0];
  var setAvailableModels = _e[1];

  var modelOptions = availableModels.map(function (model) {
    return { value: model, label: model };
  });

  function handleSave(): void {
    settingsStore.save();
    setPingResult(null);
  }

  function handlePing(): void {
    setIsPinging(true);
    setPingResult(null);
    createProvider(config)
      .ping()
      .then(function (result) {
        setPingResult(result);
        trackEvent("connection_test", { ok: result.ok, source: "settings" });
      })
      .catch(function (error) {
        setPingResult({
          ok: false,
          message: error instanceof Error ? error.message : "Ping failed",
        });
      })
      .then(function () {
        setIsPinging(false);
      });
  }

  function handleFetchModels(): void {
    setIsFetchingModels(true);
    setModelListMessage(null);
    fetchModelList(config)
      .then(function (result) {
        setAvailableModels(result.models);
        setModelListMessage(result.message);
        if (result.ok && result.models.length && result.models.indexOf(config.model) < 0) {
          settingsStore.update({ model: result.models[0] });
        }
      })
      .then(function () {
        setIsFetchingModels(false);
      });
  }

  return (
    <div className="settings-form">
      <Text variant="xLarge" block>
        AI Provider
      </Text>
      <Text variant="medium" block>
        Connect to any OpenAI or Anthropic-compatible endpoint. API keys are stored locally on this
        device.
      </Text>

      <IeSelect
        label="Provider type"
        value={config.kind}
        options={providerOptions}
        onChange={function (value) {
          settingsStore.setKind(value as ProviderKind);
        }}
      />

      <TextField
        label="Base URL"
        description="e.g. https://api.openai.com/v1 or your gateway"
        value={config.baseUrl}
        onChange={function (_event, value) {
          settingsStore.update({ baseUrl: value || "" });
        }}
      />

      <TextField
        label="API key"
        type="password"
        value={config.apiKey}
        onChange={function (_event, value) {
          settingsStore.update({ apiKey: value || "" });
        }}
      />

      {availableModels.length ? (
        <IeSelect
          label="Model"
          value={config.model}
          options={modelOptions}
          onChange={function (value) {
            settingsStore.update({ model: value });
          }}
        />
      ) : (
        <TextField
          label="Model"
          description="Fetch from /models or type a model id"
          value={config.model}
          onChange={function (_event, value) {
            settingsStore.update({ model: value || "" });
          }}
        />
      )}

      <Text variant="medium" block>
        {"Max tokens: " + config.maxTokens}
      </Text>
      <Slider
        min={256}
        max={8192}
        step={256}
        value={config.maxTokens}
        showValue={false}
        onChange={function (value) {
          settingsStore.update({ maxTokens: value });
        }}
      />

      <Text variant="medium" block>
        {"Temperature: " + config.temperature.toFixed(1)}
      </Text>
      <Slider
        min={0}
        max={2}
        step={0.1}
        value={config.temperature}
        showValue={false}
        onChange={function (value) {
          settingsStore.update({ temperature: value });
        }}
      />

      <Text variant="large" block>
        Instructions &amp; behavior
      </Text>
      <TextField
        label="Custom instructions"
        description="Appended to every chat and agent system prompt"
        multiline
        autoAdjustHeight
        value={preferences.customInstructions}
        onChange={function (_event, value) {
          settingsStore.updatePreferences({ customInstructions: value || "" });
        }}
      />

      <Text variant="large" block>
        Agent behavior
      </Text>
      <Checkbox
        checked={preferences.autoApplyEdits}
        onChange={function (_event, checked) {
          settingsStore.updatePreferences({ autoApplyEdits: Boolean(checked) });
        }}
        label="Auto-apply document edits without preview"
      />
      <Text variant="small" block>
        When off, insert and replace operations show a before/after preview with Apply and Reject.
      </Text>
      <Checkbox
        checked={preferences.reviewModeAsComments}
        onChange={function (_event, checked) {
          settingsStore.updatePreferences({ reviewModeAsComments: Boolean(checked) });
        }}
        label="Review mode — prefer comments over body edits"
      />
      <Checkbox
        checked={preferences.persistConversations}
        onChange={function (_event, checked) {
          settingsStore.updatePreferences({ persistConversations: Boolean(checked) });
        }}
        label="Remember conversation per document"
      />

      <Text variant="large" block>
        Telemetry
      </Text>
      <Checkbox
        checked={preferences.telemetryEnabled}
        onChange={function (_event, checked) {
          settingsStore.updatePreferences({ telemetryEnabled: Boolean(checked) });
        }}
        label="Send anonymous usage events"
      />
      <TextField
        label="Telemetry endpoint (optional)"
        description="POST JSON events when telemetry is on"
        placeholder="https://telemetry.example.com/events"
        value={preferences.telemetryEndpoint}
        onChange={function (_event, value) {
          settingsStore.updatePreferences({ telemetryEndpoint: value || "" });
        }}
      />

      <div className="settings-actions">
        <PrimaryButton onClick={handleSave}>Save settings</PrimaryButton>
        <DefaultButton
          disabled={!config.apiKey.trim() || isFetchingModels}
          onClick={function () {
            handleFetchModels();
          }}
        >
          {isFetchingModels ? <Spinner size={SpinnerSize.small} /> : null}
          Fetch models
        </DefaultButton>
        <DefaultButton disabled={!isConfigured || isPinging} onClick={handlePing}>
          {isPinging ? <Spinner size={SpinnerSize.small} /> : null}
          Test connection
        </DefaultButton>
      </div>

      {modelListMessage ? (
        <MessageBar
          messageBarType={availableModels.length ? MessageBarType.success : MessageBarType.warning}
        >
          {modelListMessage}
        </MessageBar>
      ) : null}

      {pingResult ? (
        <MessageBar messageBarType={pingResult.ok ? MessageBarType.success : MessageBarType.error}>
          {pingResult.ok
            ? "Connected (" + pingResult.latencyMs + "ms): " + pingResult.message
            : pingResult.message}
        </MessageBar>
      ) : null}

      {!isConfigured ? (
        <MessageBar messageBarType={MessageBarType.warning}>
          Enter a base URL, API key, and model name, then save to enable chat.
        </MessageBar>
      ) : null}
    </div>
  );
}