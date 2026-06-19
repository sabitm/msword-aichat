export interface SlashCommand {
  name: string;
  description: string;
  prompt: string;
}

export const SLASH_COMMANDS: SlashCommand[] = [
  {
    name: "fix",
    description: "Fix grammar, spelling, and clarity",
    prompt:
      "Fix grammar, spelling, and clarity in the current selection. Use get_selection first, then replace_text with the corrected text.",
  },
  {
    name: "table",
    description: "Insert a structured table",
    prompt:
      "Create a well-structured table for the described data. Call insert_table once with rows, columns, and the full cells 2D array (including headers). Do not insert an empty table and fill cells in a follow-up step.",
  },
  {
    name: "toc",
    description: "Outline or table-of-contents help",
    prompt:
      "Analyze document headings and help structure a table of contents. Use get_document_text or search_document as needed, then insert_text with a clear outline.",
  },
  {
    name: "summarize",
    description: "Summarize selection or context",
    prompt: "Summarize the current selection or relevant document context concisely.",
  },
  {
    name: "formal",
    description: "Rewrite in a formal tone",
    prompt:
      "Rewrite the current selection in a formal professional tone. Use get_selection, then replace_text.",
  },
  {
    name: "comment",
    description: "Add review comments on selection",
    prompt:
      "Review the current selection and add Word comments with specific suggestions using insert_comment. Do not change body text unless asked.",
  },
];

const commandMap = new Map(SLASH_COMMANDS.map((command) => [command.name, command]));

export function getSlashCommandHint(input: string): SlashCommand | null {
  const match = input.match(/^\/(\w*)$/);
  if (!match) return null;
  const partial = match[1];
  if (!partial) return null;
  return SLASH_COMMANDS.find((command) => command.name.startsWith(partial)) ?? null;
}

export function listMatchingSlashCommands(input: string): SlashCommand[] {
  const match = input.match(/^\/(\w*)$/);
  if (!match) return [];
  const partial = match[1];
  if (!partial) return SLASH_COMMANDS;
  return SLASH_COMMANDS.filter((command) => command.name.startsWith(partial));
}

export function expandSlashCommand(input: string): {
  displayText: string;
  promptText: string;
  command?: string;
} {
  const trimmed = input.trim();
  const match = trimmed.match(/^\/(\w+)(?:\s+([\s\S]*))?$/);
  if (!match) {
    return { displayText: trimmed, promptText: trimmed };
  }

  const [, commandName, args] = match;
  const command = commandMap.get(commandName);
  if (!command) {
    return { displayText: trimmed, promptText: trimmed };
  }

  const details = args?.trim();
  const promptText = details ? `${command.prompt}\n\nDetails: ${details}` : command.prompt;
  return {
    displayText: trimmed,
    promptText,
    command: command.name,
  };
}