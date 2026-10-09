import React from "react";
import { createRoot } from "react-dom/client";
import { AdaptiveCard } from "../src/Tab/AdaptiveCard";
import { createMessageActivity } from "../src/Tab/copilot";
import type { AgentAction } from "../src/Tab/types";

// Run in a real browser against the Vite server:
// (await import("/tabs/home/tests/adaptive-card.browser.ts")).runCardBrowserChecks()
export async function runCardBrowserChecks(): Promise<string[]> {
  const host = document.createElement("div");
  document.body.append(host);
  const root = createRoot(host);
  const submitted: AgentAction[] = [];
  const results: string[] = [];
  const check = (condition: boolean, name: string) => {
    if (!condition) throw new Error(name);
    results.push(name);
  };
  const waitFor = async (predicate: () => boolean) => {
    for (let index = 0; index < 100; index++) {
      if (predicate()) return;
      await new Promise((resolve) => setTimeout(resolve, 20));
    }
    throw new Error("Timed out waiting for the card renderer.");
  };
  const content = {
    type: "AdaptiveCard", version: "1.5",
    body: [
      { type: "TextBlock", text: '<img src="https://example.com/unwanted" onerror="alert(1)">' },
      { type: "Input.Text", id: "customer", label: "Customer email", isRequired: true, regex: "^.+@.+$", errorMessage: "Enter a valid email." },
      { type: "Input.Number", id: "quantity", label: "Quantity", min: 1, max: 5, value: 2 },
      { type: "Input.Toggle", id: "approved", title: "Approved", valueOn: "yes", valueOff: "no", value: "yes" },
      { type: "Image", url: "javascript:alert(1)" },
    ],
    actions: [
      { type: "Action.Submit", title: "Save", data: { command: "save", order: { id: 42 } } },
      { type: "Action.Submit", title: "Cancel", associatedInputs: "none", data: { command: "cancel" } },
      { type: "Action.OpenUrl", title: "Bad link", url: "data:text/html,unsafe" },
      { type: "Action.Execute", title: "Unsupported" },
    ],
  };
  const render = (disabled = false) => root.render(React.createElement(AdaptiveCard, {
    content, disabled, onAction: (action) => submitted.push(action),
  }));
  try {
    render();
    await waitFor(() => Boolean(host.querySelector("input")));
    const email = host.querySelector<HTMLInputElement>('input[type="text"]')!;
    const save = [...host.querySelectorAll("button")].find((button) => button.textContent === "Save")!;
    const cancel = [...host.querySelectorAll("button")].find((button) => button.textContent === "Cancel")!;
    check(Boolean(email && save && cancel), "Official SDK renders labelled inputs and Submit buttons");
    save.click();
    check(submitted.length === 0, "Required input validation blocks empty Submit");
    check(host.textContent?.includes("Enter a valid email.") === true, "SDK displays input validation error");
    email.value = "invalid";
    email.dispatchEvent(new Event("input", { bubbles: true }));
    save.click();
    check(submitted.length === 0, "Regex validation blocks invalid Submit");
    cancel.click();
    check(submitted.length === 1 && (submitted[0].value as Record<string, unknown>).command === "cancel",
      "associatedInputs none supports Cancel without validating the form");
    email.value = "customer@example.com";
    email.dispatchEvent(new Event("input", { bubbles: true }));
    save.click();
    const activity = createMessageActivity(submitted[1], "continuous-1");
    const value = activity.value as Record<string, unknown>;
    check(submitted.length === 2 && value.customer === "customer@example.com" &&
      value.command === "save" && (value.order as { id: number }).id === 42 &&
      value.quantity === "2" && value.approved === "yes" && activity.text === undefined,
    "SDK merges hidden object data and entered inputs into Activity.value");
    check(host.querySelectorAll("img").length === 0, "Unsafe images and embedded HTML never create image elements");
    check(host.textContent?.includes("unsafe card link") === true &&
      host.textContent?.includes("Unsupported card action") === true,
    "Unsafe and unsupported actions have visible fallbacks");
    render(true);
    await waitFor(() => host.querySelector("fieldset")?.disabled === true);
    save.click();
    check(submitted.length === 2, "Busy/completed card cannot submit again");
    check(email.value === "customer@example.com", "Disabling card actions preserves entered form values");
    return results;
  } finally {
    root.unmount();
    host.remove();
  }
}
