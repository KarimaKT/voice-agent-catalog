import type { AgentDefinition } from "./types";

export interface HarnessSupport {
  supported: boolean;
  label: string;
}

export function getHarnessSupport(harness: string | undefined): HarnessSupport {
  const normalized = harness?.trim().toLowerCase();
  if (normalized === "standard") {
    return { supported: true, label: "Standard" };
  }
  return {
    supported: false,
    label: harness?.trim() || "Unknown harness",
  };
}

function normalizeCompletionText(value: string): string {
  return value.trim().replace(/\s+/g, " ").toLocaleLowerCase();
}

export function completesConversation(
  agent: Pick<AgentDefinition, "endsConversation" | "completionPhrase">,
  responseText: string,
): boolean {
  return agent.endsConversation === true && agent.completionPhrase.trim().length > 0 &&
    normalizeCompletionText(responseText) ===
      normalizeCompletionText(agent.completionPhrase);
}

export function voiceInterfaceInstructions(locale: string): string {
  return `Interface context: this conversation uses speech and may show an avatar. ` +
    `For the rest of this conversation, answer in ${locale}, even when the user uses another language. ` +
    `Keep spoken replies concise and natural. Avoid Markdown, decorative symbols, and reading out URLs or file contents. ` +
    `Show cards, citations, images, and files in chat instead of narrating them. ` +
    `Follow your existing instructions and authorization rules. ` +
    `This is interface setup, not a request to execute a tool or complete a business task. ` +
    `Briefly introduce yourself and invite the user's first request.`;
}
