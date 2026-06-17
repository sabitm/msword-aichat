import type {
  AgentMessage,
  AgentRunResult,
  AgentStep,
  CompletionResult,
  PendingEdit,
} from "../types/agent";
import { MAX_AGENT_STEPS } from "../types/agent";
import type { ProviderConfig } from "../types/llm";
import { createProvider } from "../llm/factory";
import { TOOL_DEFINITIONS, executeTool } from "./tools/registry";

function createId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function serializeToolOutput(output: Record<string, unknown>): string {
  return JSON.stringify(output);
}

export async function runAgent(options: {
  config: ProviderConfig;
  messages: AgentMessage[];
  autoApplyEdits: boolean;
  onStep?: (step: AgentStep) => void;
}): Promise<AgentRunResult> {
  const provider = createProvider(options.config);
  const loopMessages = [...options.messages];
  const steps: AgentStep[] = [];
  let pendingEdit: PendingEdit | undefined;
  let finalText = "";

  const emit = (step: AgentStep) => {
    steps.push(step);
    options.onStep?.(step);
  };

  for (let stepIndex = 0; stepIndex < MAX_AGENT_STEPS; stepIndex += 1) {
    const completion: CompletionResult = await provider.complete({
      messages: loopMessages,
      tools: TOOL_DEFINITIONS,
      temperature: options.config.temperature,
      maxTokens: options.config.maxTokens,
    });

    if (completion.stopReason === "error") {
      return {
        text: finalText,
        steps,
        pendingEdit,
        error: completion.error ?? "Agent request failed",
      };
    }

    if (completion.text.trim()) {
      finalText = completion.text.trim();
      emit({
        id: createId(),
        type: "text",
        text: finalText,
        status: "done",
      });
    }

    if (completion.toolCalls.length === 0 || completion.stopReason === "end") {
      return { text: finalText, steps, pendingEdit };
    }

    loopMessages.push({
      role: "assistant",
      content: completion.text,
      toolCalls: completion.toolCalls,
    });

    for (const toolCall of completion.toolCalls) {
      const callStepId = createId();
      emit({
        id: callStepId,
        type: "tool_call",
        toolName: toolCall.name,
        input: toolCall.arguments,
        status: "running",
      });

      const result = await executeTool(toolCall.name, toolCall.arguments, {
        autoApplyEdits: options.autoApplyEdits,
      });

      const resultStep: AgentStep = {
        id: createId(),
        type: "tool_result",
        toolName: toolCall.name,
        output: serializeToolOutput(result.output),
        status: result.success ? "done" : "error",
        error: result.error,
      };
      emit(resultStep);

      const callStep = steps.find((s) => s.id === callStepId);
      if (callStep) {
        callStep.status = result.success ? "done" : "error";
        callStep.error = result.error;
      }

      loopMessages.push({
        role: "tool",
        content: serializeToolOutput(result.output),
        toolCallId: toolCall.id,
        name: toolCall.name,
      });

      if (result.pendingEdit) {
        pendingEdit = result.pendingEdit;
        emit({
          id: createId(),
          type: "tool_result",
          toolName: toolCall.name,
          output: "Edit pending user approval.",
          status: "pending",
        });
        return {
          text:
            finalText ||
            "I prepared a document edit. Review the preview below and click Apply or Reject.",
          steps,
          pendingEdit,
        };
      }

      if (!result.success) {
        return {
          text: finalText || result.error || "A tool failed.",
          steps,
          pendingEdit,
          error: result.error,
        };
      }
    }
  }

  return {
    text: finalText || "Agent stopped after reaching the maximum number of steps.",
    steps,
    pendingEdit,
    error: "Maximum agent steps reached",
  };
}