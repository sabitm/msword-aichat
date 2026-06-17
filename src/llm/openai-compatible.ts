import type { ChatEvent, ChatRequest, PingResult, ProviderConfig } from "../types/llm";
import type { LLMProvider } from "./provider";
import { normalizeBaseUrl } from "./provider";

interface OpenAIStreamChunk {
  choices?: Array<{
    delta?: { content?: string };
    finish_reason?: string | null;
  }>;
  error?: { message?: string };
}

export class OpenAICompatibleProvider implements LLMProvider {
  readonly kind = "openai" as const;

  constructor(private readonly config: ProviderConfig) {}

  async *chat(request: ChatRequest): AsyncIterable<ChatEvent> {
    const url = `${normalizeBaseUrl(this.config.baseUrl)}/chat/completions`;
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${this.config.apiKey}`,
      },
      body: JSON.stringify({
        model: this.config.model,
        messages: request.messages,
        stream: request.stream ?? true,
        max_tokens: request.maxTokens ?? this.config.maxTokens,
        temperature: request.temperature ?? this.config.temperature,
      }),
    });

    if (!response.ok) {
      const body = await response.text();
      yield { type: "error", error: `HTTP ${response.status}: ${body}` };
      return;
    }

    if (!request.stream) {
      const json = (await response.json()) as OpenAIStreamChunk & {
        choices?: Array<{ message?: { content?: string } }>;
      };
      const text = json.choices?.[0]?.message?.content ?? "";
      if (text) {
        yield { type: "text_delta", text };
      }
      yield { type: "done" };
      return;
    }

    const reader = response.body?.getReader();
    if (!reader) {
      yield { type: "error", error: "No response body" };
      return;
    }

    const decoder = new TextDecoder();
    let buffer = "";

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop() ?? "";

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed.startsWith("data:")) continue;

          const payload = trimmed.slice(5).trim();
          if (payload === "[DONE]") {
            yield { type: "done" };
            return;
          }

          try {
            const chunk = JSON.parse(payload) as OpenAIStreamChunk;
            if (chunk.error?.message) {
              yield { type: "error", error: chunk.error.message };
              return;
            }
            const text = chunk.choices?.[0]?.delta?.content;
            if (text) {
              yield { type: "text_delta", text };
            }
          } catch {
            // Ignore malformed SSE chunks.
          }
        }
      }

      yield { type: "done" };
    } finally {
      reader.releaseLock();
    }
  }

  async ping(): Promise<PingResult> {
    const started = performance.now();
    let responseText = "";

    for await (const event of this.chat({
      messages: [{ role: "user", content: "Reply with exactly: pong" }],
      stream: false,
      maxTokens: 16,
      temperature: 0,
    })) {
      if (event.type === "error") {
        return { ok: false, message: event.error ?? "Unknown error" };
      }
      if (event.type === "text_delta" && event.text) {
        responseText += event.text;
      }
    }

    return {
      ok: true,
      message: responseText.trim() || "Connection successful",
      latencyMs: Math.round(performance.now() - started),
    };
  }
}