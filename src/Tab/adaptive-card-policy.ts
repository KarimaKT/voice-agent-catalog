const supportedElements = new Set([
  "TextBlock", "RichTextBlock", "TextRun", "Image", "ImageSet", "FactSet",
  "Container", "Column", "ColumnSet", "ActionSet", "Table", "TableRow", "TableCell",
  "Input.Text", "Input.Number", "Input.Date", "Input.Time", "Input.Toggle", "Input.ChoiceSet",
]);
const supportedActions = new Set([
  "Action.Submit", "Action.OpenUrl", "Action.ShowCard", "Action.ToggleVisibility",
]);

export interface CardPolicyResult {
  card?: Record<string, unknown>;
  warnings: string[];
}

export function safeHttpUrl(value: unknown): string | undefined {
  if (typeof value !== "string" || /[\u0000-\u001f\u007f]/.test(value)) {
    return undefined;
  }
  try {
    const url = new URL(value);
    return (url.protocol === "http:" || url.protocol === "https:") &&
      !url.username && !url.password ? url.toString() : undefined;
  } catch {
    return undefined;
  }
}

function record(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

// The SDK owns layout, inputs and validation. This host only bounds the document
// and restricts capabilities before any resource can be rendered or action run.
export function sanitizeAdaptiveCard(content: unknown): CardPolicyResult {
  const warnings = new Set<string>();
  const warn = (message: string) => warnings.add(`${message} Use text to respond if needed.`);
  let nodes = 0;

  function fallback(item: Record<string, unknown>, action: boolean, depth: number): unknown {
    if (record(item.fallback)) {
      return walk(item.fallback, depth + 1, action);
    }
    return action ? undefined : { type: "TextBlock", text: "Unsupported card content.", wrap: true };
  }

  function walk(value: unknown, depth: number, action = false): unknown {
    if (++nodes > 2000 || depth > 24) {
      throw new Error("This card exceeds the supported size or nesting limit.");
    }
    if (Array.isArray(value)) {
      return value.map((item) => walk(item, depth + 1, action)).filter((item) => item !== undefined);
    }
    if (!record(value)) {
      return value;
    }
    const item = value;
    const type = typeof item.type === "string" ? item.type : "";
    if (action || type.startsWith("Action.")) {
      if (!supportedActions.has(type)) {
        warn(`Unsupported card action: ${type || "unknown"}.`);
        return fallback(item, true, depth);
      }
      if (type === "Action.OpenUrl" && !safeHttpUrl(item.url)) {
        warn("An unsafe card link was blocked.");
        return fallback(item, true, depth);
      }
    } else if (type && type !== "AdaptiveCard" && !supportedElements.has(type)) {
      warn(`Unsupported card element: ${type}.`);
      return fallback(item, false, depth);
    }
    if (type === "Image" && !safeHttpUrl(item.url)) {
      warn("An unsafe card image was blocked.");
      return { type: "TextBlock", text: "Card image unavailable.", wrap: true };
    }
    const result: Record<string, unknown> = {};
    for (const [key, child] of Object.entries(item)) {
      if (key === "refresh" || key === "authentication" || key === "resources") {
        warn(`Unsupported card capability: ${key}.`);
        continue;
      }
      if (key === "url" || key === "iconUrl") {
        const url = safeHttpUrl(child);
        if (url) result[key] = url;
        else warn(key === "iconUrl" ? "An unsafe action icon was blocked." : "An unsafe card resource was blocked.");
      } else if (key === "backgroundImage") {
        const url = safeHttpUrl(record(child) ? child.url : child);
        if (url) result[key] = record(child) ? { ...child, url } : url;
        else warn("An unsafe card background was blocked.");
      } else if (key === "data" || key === "metadata") {
        // Submit data is opaque business data, not a card subtree.
        result[key] = child;
      } else {
        let boundedChild = child;
        if (key === "actions" && Array.isArray(child) && child.length > 20) {
          warn("Only the first 20 card actions are supported.");
          boundedChild = child.slice(0, 20);
        }
        const sanitized = walk(boundedChild, depth + 1,
          key === "actions" || key === "selectAction" || key === "inlineAction");
        if (sanitized !== undefined) result[key] = sanitized;
      }
    }
    return result;
  }

  try {
    const serialized = JSON.stringify(content);
    if (!serialized || serialized.length > 128_000) {
      throw new Error("This card exceeds the supported size limit.");
    }
    const original: unknown = JSON.parse(serialized);
    if (!record(original) || original.type !== "AdaptiveCard") {
      throw new Error("This attachment is not a supported Adaptive Card.");
    }
    const version = original.version;
    if (typeof version !== "string" || !/^1\.[0-6]$/.test(version)) {
      throw new Error("This card requires an unsupported Adaptive Card version.");
    }
    if (("body" in original && !Array.isArray(original.body)) ||
        ("actions" in original && !Array.isArray(original.actions))) {
      throw new Error("This card has an unsupported body or action format.");
    }
    const card = walk(original, 0) as Record<string, unknown>;
    return { card, warnings: [...warnings] };
  } catch (error) {
    warn(error instanceof Error ? error.message : "This card could not be read.");
    return { warnings: [...warnings] };
  }
}
