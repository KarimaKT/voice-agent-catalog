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

export function voiceInterfaceInstructions(locale: string, inputLocale = locale): string {
  return `Remember the following across this conversation: This conversation is happening over a voice channel. ` +
    `Keep all conversational responses short and natural, with no markup or decorative symbols. ` +
    `The user will be speaking ${inputLocale}, but you must always answer in ${locale}. ` +
    `Keep cards, citations, files and HTML tool payloads in their structured outputs; do not read them aloud. ` +
    `Follow your existing instructions and authorization rules. ` +
    `This interface context is not a request to execute a tool. ` +
    `Here is the user's response to your request:\n\n`;
}
