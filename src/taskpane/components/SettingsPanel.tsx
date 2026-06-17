import {
  Button,
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
import { useSettingsStore } from "../../settings/store";
import type { PingResult, ProviderKind } from "../../types/llm";

const providerLabels: Record<ProviderKind, string> = {
  openai: "OpenAI-compatible",
  anthropic: "Anthropic-compatible",
};

export function SettingsPanel() {
  const config = useSettingsStore((s) => s.config);
  const update = useSettingsStore((s) => s.update);
  const setKind = useSettingsStore((s) => s.setKind);
  const save = useSettingsStore((s) => s.save);
  const isConfigured = useSettingsStore((s) => s.isConfigured);

  const [pingResult, setPingResult] = useState<PingResult | null>(null);
  const [isPinging, setIsPinging] = useState(false);

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
    } catch (error) {
      setPingResult({
        ok: false,
        message: error instanceof Error ? error.message : "Ping failed",
      });
    } finally {
      setIsPinging(false);
    }
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

      <Field label="Model">
        <Input
          value={config.model}
          onChange={(_event, data) => update({ model: data.value })}
        />
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

      <div className="settings-actions">
        <Button appearance="primary" onClick={handleSave}>
          Save settings
        </Button>
        <Button appearance="secondary" disabled={!isConfigured || isPinging} onClick={handlePing}>
          {isPinging ? <Spinner size="tiny" /> : null}
          Test connection
        </Button>
      </div>

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