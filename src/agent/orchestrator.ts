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

function buildTurnTranscript(
  loopMessages: AgentMessage[],
  turnStartIndex: number,
  finalText: string,
): AgentMessage[] {
  const transcript = loopMessages.slice(turnStartIndex);

  if (!finalText.trim()) {
    return transcript;
  }

  const last = transcript[transcript.length - 1];
  if (!last) {
    transcript.push({ role: "assistant", content: finalText });
    return transcript;
  }

  if (last.role === "tool" || (last.role === "assistant" && last.toolCalls?.length)) {
    transcript.push({ role: "assistant", content: finalText });
    return transcript;
  }

  if (last.role === "assistant" && !last.content.trim()) {
    last.content = finalText;
  }

  return transcript;
}

export async function runAgent(options: {
  config: ProviderConfig;
  messages: AgentMessage[];
  autoApplyEdits: boolean;
  onStep?: (step: AgentStep) => void;
}): Promise<AgentRunResult> {
  const provider = createProvider(options.config);
  const loopMessages = [...options.messages];
  const turnStartIndex = loopMessages.length;
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
        transcript: buildTurnTranscript(loopMessages, turnStartIndex, finalText),
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
      return {
        text: finalText,
        steps,
        pendingEdit,
        transcript: buildTurnTranscript(loopMessages, turnStartIndex, finalText),
      };
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
          transcript: buildTurnTranscript(loopMessages, turnStartIndex, finalText),
        };
      }

      if (!result.success) {
        return {
          text: finalText || result.error || "A tool failed.",
          steps,
          pendingEdit,
          error: result.error,
          transcript: buildTurnTranscript(loopMessages, turnStartIndex, finalText),
        };
      }
    }
  }

  return {
    text: finalText || "Agent stopped after reaching the maximum number of steps.",
    steps,
    pendingEdit,
    error: "Maximum agent steps reached",
    transcript: buildTurnTranscript(loopMessages, turnStartIndex, finalText),
  };
}