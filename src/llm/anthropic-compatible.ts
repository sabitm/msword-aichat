import type { AgentMessage, CompletionResult, ToolCall } from "../types/agent";
import type { ChatEvent, ChatRequest, PingResult, ProviderConfig } from "../types/llm";
import type { LLMProvider } from "./provider";
import { normalizeBaseUrl } from "./provider";

interface AnthropicStreamEvent {
  type: string;
  delta?: { type?: string; text?: string };
  error?: { message?: string };
  message?: {
    content?: Array<
      | { type: "text"; text?: string }
      | { type: "tool_use"; id?: string; name?: string; input?: Record<string, unknown> }
    >;
  };
}

interface AnthropicCompletionResponse {
  content?: Array<
    | { type: "text"; text?: string }
    | { type: "tool_use"; id?: string; name?: string; input?: Record<string, unknown> }
  >;
  stop_reason?: string;
  error?: { message?: string };
}

export class AnthropicCompatibleProvider implements LLMProvider {
  readonly kind = "anthropic" as const;
  readonly supportsTools = true;

  constructor(private readonly config: ProviderConfig) {}

  private toAnthropicMessages(messages: AgentMessage[]) {
    const result: Array<Record<string, unknown>> = [];

    for (const message of messages) {
      if (message.role === "system") continue;

      if (message.role === "tool") {
        result.push({
          role: "user",
          content: [
            {
              type: "tool_result",
              tool_use_id: message.toolCallId,
              content: message.content,
            },
          ],
        });
        continue;
      }

      if (message.role === "assistant" && message.toolCalls?.length) {
        const content: Array<Record<string, unknown>> = [];
        if (message.content.trim()) {
          content.push({ type: "text", text: message.content });
        }
        for (const call of message.toolCalls) {
          content.push({
            type: "tool_use",
            id: call.id,
            name: call.name,
            input: JSON.parse(call.arguments || "{}"),
          });
        }
        result.push({ role: "assistant", content });
        continue;
      }

      result.push({
        role: message.role,
        content: message.content,
      });
    }

    return result;
  }

  private toAnthropicTools(tools: ChatRequest["tools"]) {
    return tools?.map((tool) => ({
      name: tool.name,
      description: tool.description,
      input_schema: tool.parameters,
    }));
  }

  async complete(request: Parameters<LLMProvider["complete"]>[0]): Promise<CompletionResult> {
    const url = `${normalizeBaseUrl(this.config.baseUrl)}/messages`;
    const systemMessage = request.messages.find((m) => m.role === "system");

    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": this.config.apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: this.config.model,
        system: systemMessage?.content,
        messages: this.toAnthropicMessages(request.messages),
        tools: this.toAnthropicTools(request.tools),
        stream: false,
        max_tokens: request.maxTokens ?? this.config.maxTokens,
        temperature: request.temperature ?? this.config.temperature,
      }),
    });

    if (!response.ok) {
      const body = await response.text();
      return {
        text: "",
        toolCalls: [],
        stopReason: "error",
        error: `HTTP ${response.status}: ${body}`,
      };
    }

    const json = (await response.json()) as AnthropicCompletionResponse;
    if (json.error?.message) {
      return {
        text: "",
        toolCalls: [],
        stopReason: "error",
        error: json.error.message,
      };
    }

    const text =
      json.content
        ?.filter((block) => block.type === "text")
        .map((block) => block.text ?? "")
        .join("") ?? "";

    const toolCalls: ToolCall[] =
      json.content
        ?.filter((block) => block.type === "tool_use")
        .map((block) => ({
          id: block.id ?? "",
          name: block.name ?? "",
          arguments: JSON.stringify(block.input ?? {}),
        }))
        .filter((call) => call.id && call.name) ?? [];

    return {
      text,
      toolCalls,
      stopReason: json.stop_reason === "tool_use" ? "tool_calls" : "end",
    };
  }

  async *chat(request: ChatRequest): AsyncIterable<ChatEvent> {
    const url = `${normalizeBaseUrl(this.config.baseUrl)}/messages`;
    const systemMessage = request.messages.find((m) => m.role === "system");
    const nonSystemMessages = request.messages
      .filter((m) => m.role !== "system")
      .map((m) => ({
        role: m.role as "user" | "assistant",
        content: m.content,
      }));

    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": this.config.apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: this.config.model,
        system: systemMessage?.content,
        messages: nonSystemMessages,
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
      const json = (await response.json()) as AnthropicStreamEvent;
      const text =
        json.message?.content
          ?.filter((block) => block.type === "text")
          .map((block) => block.text ?? "")
          .join("") ?? "";
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
    let currentEvent = "";

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop() ?? "";

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed) continue;

          if (trimmed.startsWith("event:")) {
            currentEvent = trimmed.slice(6).trim();
            continue;
          }

          if (!trimmed.startsWith("data:")) continue;

          const payload = trimmed.slice(5).trim();
          try {
            const chunk = JSON.parse(payload) as AnthropicStreamEvent;
            if (chunk.error?.message) {
              yield { type: "error", error: chunk.error.message };
              return;
            }

            if (
              currentEvent === "content_block_delta" &&
              chunk.delta?.type === "text_delta" &&
              chunk.delta.text
            ) {
              yield { type: "text_delta", text: chunk.delta.text };
            }

            if (currentEvent === "message_stop" || chunk.type === "message_stop") {
              yield { type: "done" };
              return;
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