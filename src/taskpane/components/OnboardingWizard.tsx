import {
  Button,
  Checkbox,
  Dropdown,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Option,
  Spinner,
  Text,
} from "@fluentui/react-components";
import { useState } from "react";
import { createProvider } from "../../llm/factory";
import { fetchModelList } from "../../llm/models";
import { useSettingsStore } from "../../settings/store";
import { trackEvent } from "../../telemetry";
import type { InteractionMode } from "../../types/agent";
import type { PingResult, ProviderKind } from "../../types/llm";

const providerLabels: Record<ProviderKind, string> = {
  openai: "OpenAI-compatible",
  anthropic: "Anthropic-compatible",
};

const STEPS = ["Welcome", "Provider", "Connection", "Modes", "Ready"] as const;

interface OnboardingWizardProps {
  onComplete: () => void;
}

export function OnboardingWizard({ onComplete }: OnboardingWizardProps) {
  const config = useSettingsStore((s) => s.config);
  const preferences = useSettingsStore((s) => s.preferences);
  const update = useSettingsStore((s) => s.update);
  const updatePreferences = useSettingsStore((s) => s.updatePreferences);
  const setKind = useSettingsStore((s) => s.setKind);
  const save = useSettingsStore((s) => s.save);
  const completeOnboarding = useSettingsStore((s) => s.completeOnboarding);
  const isConfigured = useSettingsStore((s) => s.isConfigured);

  const [step, setStep] = useState(0);
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
      trackEvent("connection_test", { ok: result.ok, source: "onboarding" });
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

  const handleFinish = () => {
    save();
    completeOnboarding();
    trackEvent("onboarding_completed", {
      mode: preferences.interactionMode,
      provider: config.kind,
    });
    onComplete();
  };

  const canAdvanceFromProvider = isConfigured;
  const canAdvanceFromConnection = pingResult?.ok === true;

  return (
    <div className="onboarding-overlay">
      <div className="onboarding-card">
        <Text size={500} weight="semibold">
          Welcome to Word AI Chat
        </Text>
        <Text size={200} className="onboarding-steps">
          Step {step + 1} of {STEPS.length}: {STEPS[step]}
        </Text>

        {step === 0 ? (
          <div className="onboarding-body">
            <Text>
              Chat with your document using any OpenAI- or Anthropic-compatible endpoint. Agent mode
              can read and edit Word with your approval.
            </Text>
            <ul className="onboarding-list">
              <li>API keys stay on this device</li>
              <li>Edits are previewed before applying (unless you enable auto-apply)</li>
              <li>Works on Word Desktop and Word on the web</li>
            </ul>
          </div>
        ) : null}

        {step === 1 ? (
          <div className="onboarding-body onboarding-form">
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

            <Field label="Base URL">
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

            <div className="settings-actions">
              <Button appearance="secondary" onClick={handleSave}>
                Save
              </Button>
              <Button
                appearance="secondary"
                disabled={!config.apiKey.trim() || isFetchingModels}
                onClick={() => void handleFetchModels()}
              >
                {isFetchingModels ? <Spinner size="tiny" /> : null}
                Fetch models
              </Button>
            </div>

            {modelListMessage ? (
              <MessageBar intent={availableModels.length ? "success" : "warning"}>
                <MessageBarBody>{modelListMessage}</MessageBarBody>
              </MessageBar>
            ) : null}
          </div>
        ) : null}

        {step === 2 ? (
          <div className="onboarding-body">
            <Text>Verify your endpoint responds before chatting.</Text>
            <div className="settings-actions">
              <Button appearance="secondary" onClick={handleSave}>
                Save settings
              </Button>
              <Button
                appearance="primary"
                disabled={!isConfigured || isPinging}
                onClick={() => void handlePing()}
              >
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
          </div>
        ) : null}

        {step === 3 ? (
          <div className="onboarding-body onboarding-form">
            <Field label="Default interaction mode">
              <Dropdown
                value={preferences.interactionMode === "agent" ? "Agent" : "Chat"}
                selectedOptions={[preferences.interactionMode]}
                onOptionSelect={(_event, data) => {
                  const mode = data.optionValue as InteractionMode | undefined;
                  if (mode) updatePreferences({ interactionMode: mode });
                }}
              >
                <Option value="agent">Agent — reads and edits with tools</Option>
                <Option value="chat">Chat — streaming Q&A, no tools</Option>
              </Dropdown>
            </Field>
            <Text size={200}>
              <strong>Agent</strong> runs a tool loop to search, replace, style, and format your
              document. <strong>Chat</strong> streams answers with optional document context only.
            </Text>
            <Checkbox
              checked={preferences.autoApplyEdits}
              onChange={(_event, data) =>
                updatePreferences({ autoApplyEdits: Boolean(data.checked) })
              }
              label="Auto-apply document edits without preview"
            />
            <Checkbox
              checked={preferences.telemetryEnabled}
              onChange={(_event, data) =>
                updatePreferences({ telemetryEnabled: Boolean(data.checked) })
              }
              label="Send anonymous usage events (opt-in, local debug only for now)"
            />
          </div>
        ) : null}

        {step === 4 ? (
          <div className="onboarding-body">
            <Text>You are ready to chat. Open the Chat tab, select document context if needed, and send a message.</Text>
            <ul className="onboarding-list">
              <li>Use Selection context for quick actions on highlighted text</li>
              <li>Agent edits show a before/after preview — click Apply to write to Word</li>
              <li>Change provider or mode anytime in Settings</li>
            </ul>
          </div>
        ) : null}

        <div className="onboarding-footer">
          <Button
            appearance="subtle"
            disabled={step === 0}
            onClick={() => setStep((current) => Math.max(0, current - 1))}
          >
            Back
          </Button>
          {step < STEPS.length - 1 ? (
            <Button
              appearance="primary"
              disabled={
                (step === 1 && !canAdvanceFromProvider) ||
                (step === 2 && !canAdvanceFromConnection)
              }
              onClick={() => setStep((current) => current + 1)}
            >
              Next
            </Button>
          ) : (
            <Button appearance="primary" onClick={handleFinish}>
              Get started
            </Button>
          )}
        </div>
      </div>
    </div>
  );
}