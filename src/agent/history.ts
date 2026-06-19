import type { AgentMessage } from "../types/agent";

export interface AgentHistoryUiMessage {
  role: "user" | "assistant";
  content: string;
  agentTranscript?: AgentMessage[];
}

export function appendUiMessageToAgentHistory(
  agentMessages: AgentMessage[],
  message: AgentHistoryUiMessage,
): void {
  if (message.role === "user") {
    agentMessages.push({ role: "user", content: message.content });
    return;
  }

  if (message.agentTranscript && message.agentTranscript.length > 0) {
    for (var i = 0; i < message.agentTranscript.length; i++) {
      agentMessages.push(message.agentTranscript[i]);
    }
    return;
  }

  if (message.content.trim()) {
    agentMessages.push({ role: "assistant", content: message.content });
  }
}

export function buildAgentHistoryFromUi(messages: AgentHistoryUiMessage[]): AgentMessage[] {
  var agentMessages: AgentMessage[] = [];
  for (var i = 0; i < messages.length; i++) {
    appendUiMessageToAgentHistory(agentMessages, messages[i]);
  }
  return agentMessages;
}