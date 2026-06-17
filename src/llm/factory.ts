import type { ProviderConfig } from "../types/llm";
import { AnthropicCompatibleProvider } from "./anthropic-compatible";
import { OpenAICompatibleProvider } from "./openai-compatible";
import type { LLMProvider } from "./provider";

export function createProvider(config: ProviderConfig): LLMProvider {
  switch (config.kind) {
    case "openai":
      return new OpenAICompatibleProvider(config);
    case "anthropic":
      return new AnthropicCompatibleProvider(config);
    default: {
      const exhaustive: never = config.kind;
      throw new Error(`Unsupported provider kind: ${exhaustive}`);
    }
  }
}