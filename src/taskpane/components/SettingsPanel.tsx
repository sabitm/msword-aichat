import {
  Button,
  Checkbox,
  Dropdown,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Option,
  Slider,
  Spinner,
  Text,
} from "@fluentui/react-components";
import { useState } from "react";
import { createProvider } from "../../llm/factory";
import { fetchModelList } from "../../llm/models";
import { useSettingsStore } from "../../settings/store";
import { trackEvent } from "../../telemetry";
import type { PingResult, ProviderKind } from "../../types/llm";

const providerLabels: Record<ProviderKind, string> = {
  openai: "OpenAI-compatible",
  anthropic: "Anthropic-compatible",
};

export function SettingsPanel() {
  const config = useSettingsStore((s) => s.config);
  const preferences = useSettingsStore((s) => s.preferences);
  const update = useSettingsStore((s) => s.update);
  const updatePreferences = useSettingsStore((s) => s.updatePreferences);
  const setKind = useSettingsStore((s) => s.setKind);
  const save = useSettingsStore((s) => s.save);
  const isConfigured = useSettingsStore((s) => s.isConfigured);

  const [pingResult, setPingResult] = useState<PingResult | null>(null);
  const [isPinging, setIsPinging] = useState(false);
  const [isFetchingModels, setIsFetchingModels] = useState(false);
  const [modelListMessage, setModelListMessage] = useState<string | null>(null);
  const [availableModels, setAvailableModels] = useState<string[]>([]);

  const handleSave = () => {
    save();
    setPingResult(null);
  };

  const handlePing = async () => {
    setIsPinging(true);
    setPingResult(null);
    try {
      const provider = createProvider(config);
      const result = await provider.ping();
      setPingResult(result);
      trackEvent("connection_test", { ok: result.ok, source: "settings" });
    } catch (error) {
      setPingResult({
        ok: false,
        message: error instanceof Error ? error.message : "Ping failed",
      });
    } finally {
      setIsPinging(false);
    }
  };

  const handleFetchModels = async () => {
    setIsFetchingModels(true);
    setModelListMessage(null);
    const result = await fetchModelList(config);
    setAvailableModels(result.models);
    setModelListMessage(result.message);
    if (result.ok && result.models.length && !result.models.includes(config.model)) {
      update({ model: result.models[0] });
    }
    setIsFetchingModels(false);
  };

  return (
    <div className="settings-form">
      <Text size={500} weight="semibold">
        AI Provider
      </Text>
      <Text size={300}>
        Connect to any OpenAI or Anthropic-compatible endpoint. API keys are stored locally on
        this device.
      </Text>

      <Field label="Provider type">
        <Dropdown
          value={providerLabels[config.kind]}
          selectedOptions={[config.kind]}
          onOptionSelect={(_event, data) => {
            const kind = data.optionValue as ProviderKind | undefined;
            if (kind) setKind(kind);
          }}
        >
          <Option value="openai">{providerLabels.openai}</Option>
          <Option value="anthropic">{providerLabels.anthropic}</Option>
        </Dropdown>
      </Field>

      <Field label="Base URL" hint="e.g. https://api.openai.com/v1 or your gateway">
        <Input
          value={config.baseUrl}
          onChange={(_event, data) => update({ baseUrl: data.value })}
        />
      </Field>

      <Field label="API key">
        <Input
          type="password"
          value={config.apiKey}
          onChange={(_event, data) => update({ apiKey: data.value })}
        />
      </Field>

      <Field label="Model" hint="Fetch from /models or type a model id">
        {availableModels.length ? (
          <Dropdown
            value={config.model}
            selectedOptions={[config.model]}
            onOptionSelect={(_event, data) => {
              if (data.optionValue) update({ model: data.optionValue });
            }}
          >
            {availableModels.map((model) => (
              <Option key={model} value={model}>
                {model}
              </Option>
            ))}
          </Dropdown>
        ) : (
          <Input
            value={config.model}
            onChange={(_event, data) => update({ model: data.value })}
          />
        )}
      </Field>

      <Field label={`Max tokens: ${config.maxTokens}`}>
        <Slider
          min={256}
          max={8192}
          step={256}
          value={config.maxTokens}
          onChange={(_event, data) => update({ maxTokens: data.value })}
        />
      </Field>

      <Field label={`Temperature: ${config.temperature.toFixed(1)}`}>
        <Slider
          min={0}
          max={2}
          step={0.1}
          value={config.temperature}
          onChange={(_event, data) => update({ temperature: data.value })}
        />
      </Field>

      <Text size={400} weight="semibold">
        Agent behavior
      </Text>
      <Checkbox
        checked={preferences.autoApplyEdits}
        onChange={(_event, data) => updatePreferences({ autoApplyEdits: Boolean(data.checked) })}
        label="Auto-apply document edits without preview"
      />
      <Text size={200}>
        When off, insert and replace operations show a before/after preview with Apply and Reject.
      </Text>

      <Checkbox
        checked={preferences.telemetryEnabled}
        onChange={(_event, data) => updatePreferences({ telemetryEnabled: Boolean(data.checked) })}
        label="Send anonymous usage events (opt-in, local debug only for now)"
      />

      <div className="settings-actions">
        <Button appearance="primary" onClick={handleSave}>
          Save settings
        </Button>
        <Button
          appearance="secondary"
          disabled={!config.apiKey.trim() || isFetchingModels}
          onClick={() => void handleFetchModels()}
        >
          {isFetchingModels ? <Spinner size="tiny" /> : null}
          Fetch models
        </Button>
        <Button appearance="secondary" disabled={!isConfigured || isPinging} onClick={handlePing}>
          {isPinging ? <Spinner size="tiny" /> : null}
          Test connection
        </Button>
      </div>

      {modelListMessage ? (
        <MessageBar intent={availableModels.length ? "success" : "warning"}>
          <MessageBarBody>{modelListMessage}</MessageBarBody>
        </MessageBar>
      ) : null}

      {pingResult ? (
        <MessageBar intent={pingResult.ok ? "success" : "error"}>
          <MessageBarBody>
            {pingResult.ok
              ? `Connected (${pingResult.latencyMs}ms): ${pingResult.message}`
              : pingResult.message}
          </MessageBarBody>
        </MessageBar>
      ) : null}

      {!isConfigured ? (
        <MessageBar intent="warning">
          <MessageBarBody>
            Enter a base URL, API key, and model name, then save to enable chat.
          </MessageBarBody>
        </MessageBar>
      ) : null}
    </div>
  );
}