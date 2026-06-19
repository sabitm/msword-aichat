import type { ModelListResult, ProviderConfig } from "../types/llm";
import { formatNetworkError, normalizeBaseUrl } from "./provider";

interface OpenAIModelsResponse {
  data?: Array<{ id?: string }>;
  error?: { message?: string };
}

export async function fetchModelList(config: ProviderConfig): Promise<ModelListResult> {
  if (!config.baseUrl.trim() || !config.apiKey.trim()) {
    return {
      ok: false,
      models: [],
      message: "Enter a base URL and API key first.",
    };
  }

  const baseUrl = normalizeBaseUrl(config.baseUrl);
  const url = `${baseUrl}/models`;
  const headers: Record<string, string> = {
    Accept: "application/json",
  };

  if (config.kind === "openai") {
    headers.Authorization = `Bearer ${config.apiKey}`;
  } else {
    headers["x-api-key"] = config.apiKey;
    headers["anthropic-version"] = "2023-06-01";
  }

  try {
    const response = await fetch(url, { headers });
    const body = (await response.json()) as OpenAIModelsResponse;

    if (!response.ok) {
      return {
        ok: false,
        models: [],
        message: body.error?.message ?? `HTTP ${response.status}`,
      };
    }

    const models = (body.data ?? [])
      .map((entry) => entry.id?.trim())
      .filter((id): id is string => Boolean(id))
      .sort((a, b) => a.localeCompare(b));

    if (!models.length) {
      return {
        ok: false,
        models: [],
        message: "No models returned. Enter a model name manually.",
      };
    }

    return {
      ok: true,
      models,
      message: `Found ${models.length} model${models.length === 1 ? "" : "s"}.`,
    };
  } catch (error) {
    return {
      ok: false,
      models: [],
      message: formatNetworkError(error),
    };
  }
}