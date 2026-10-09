import React from "react";
import * as AdaptiveCards from "adaptivecards";
import "adaptivecards/lib/adaptivecards.css";

import { safeHttpUrl, sanitizeAdaptiveCard } from "./adaptive-card-policy";
import type { AgentAction } from "./types";

// No external Markdown/HTML processor: TextBlock content must never become
// agent-controlled HTML or initiate unfiltered image/link requests.
AdaptiveCards.AdaptiveCard.onProcessMarkdown = (text, result) => {
  const escaped = document.createElement("div");
  escaped.textContent = text;
  result.outputHtml = escaped.innerHTML.replace(/\n/g, "<br>");
  result.didProcess = true;
};

export function AdaptiveCard({
  content,
  name,
  disabled,
  onAction,
}: {
  content: Record<string, unknown>;
  name?: string;
  disabled: boolean;
  onAction: (action: AgentAction) => void;
}) {
  const hostRef = React.useRef<HTMLDivElement>(null);
  const callbackRef = React.useRef(onAction);
  const disabledRef = React.useRef(disabled);
  const [warnings, setWarnings] = React.useState<string[]>([]);
  callbackRef.current = onAction;
  disabledRef.current = disabled;

  React.useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const policy = sanitizeAdaptiveCard(content);
    const notices = [...policy.warnings];
    host.replaceChildren();
    if (policy.card) {
      try {
        const card = new AdaptiveCards.AdaptiveCard();
        card.hostConfig = new AdaptiveCards.HostConfig({
          supportsInteractivity: true,
          fontFamily: '"Segoe UI", system-ui, sans-serif',
          actions: { maxActions: 20 },
        });
        card.onAnchorClicked = (_element, anchor) => {
          const url = safeHttpUrl(anchor.getAttribute("href"));
          if (url) window.open(url, "_blank", "noopener,noreferrer");
          return true;
        };
        card.onExecuteAction = (action) => {
          if (action instanceof AdaptiveCards.OpenUrlAction) {
            const url = safeHttpUrl(action.url);
            if (url) window.open(url, "_blank", "noopener,noreferrer");
          } else if (action instanceof AdaptiveCards.SubmitAction && !disabledRef.current) {
            callbackRef.current({
              title: action.title || "Submit",
              kind: "submit",
              value: action.data ?? {},
            });
          }
        };
        const context = new AdaptiveCards.SerializationContext(AdaptiveCards.Versions.v1_6);
        card.parse(policy.card, context);
        for (let index = 0; index < context.eventCount; index++) {
          notices.push(`Card notice: ${context.getEventAt(index).message}`);
        }
        for (const event of card.validateProperties().validationEvents) {
          notices.push(`Card notice: ${event.message}`);
        }
        const rendered = card.render();
        if (rendered) host.append(rendered);
        else notices.push("This card could not be displayed. Use text to respond.");
      } catch {
        host.replaceChildren();
        notices.push("This card could not be displayed. Use text to respond.");
      }
    }
    setWarnings([...new Set(notices)]);
    return () => host.replaceChildren();
  }, [content]);

  return (
    <section className="adaptive-card" aria-label={name || "Adaptive Card"}>
      <fieldset className="adaptive-card-controls" disabled={disabled}>
        <div ref={hostRef} />
      </fieldset>
      {warnings.map((warning, index) => (
        <p className="card-fallback" role="note" key={index}>{warning}</p>
      ))}
    </section>
  );
}
