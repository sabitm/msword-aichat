import type { AppPreferences } from "../settings/defaults";

export interface PromptOptions {
  customInstructions?: string;
  reviewModeAsComments?: boolean;
}

export function promptOptionsFromPreferences(preferences: AppPreferences): PromptOptions {
  return {
    customInstructions: preferences.customInstructions.trim() || undefined,
    reviewModeAsComments: preferences.reviewModeAsComments,
  };
}

export function appendCustomInstructions(base: string, options: PromptOptions): string {
  const parts = [base];

  if (options.customInstructions) {
    parts.push("", "User instructions:", options.customInstructions);
  }

  if (options.reviewModeAsComments) {
    parts.push(
      "",
      "Review mode is enabled: prefer insert_comment for suggestions instead of replace_text unless the user explicitly asks to change body text.",
    );
  }

  return parts.join("\n");
}