import { normalizeUserEmail, type VerifiedUser } from "./auth";

type CatalogRecord = Record<string, unknown>;

export interface SafeCatalogAgent {
  Id: string;
  Title: string;
  Description?: string;
  EnvironmentId: string;
  SchemaName: string;
  Enabled: true;
  CompletionPhrase: string;
  EndsConversation: boolean;
  WelcomeMessage?: string;
  Locale: string;
  VoiceName: string;
  AvatarCharacter: string;
  AvatarStyle: string;
  Audience: "Everyone" | "Restricted";
  Harness?: string;
}

const environmentIdPattern =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const schemaNamePattern = /^[A-Za-z][A-Za-z0-9_.-]{0,199}$/;

function optionalString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function requiredString(value: unknown, field: string, itemId: string): string {
  const text = optionalString(value);
  if (!text) {
    throw new Error(`Catalog item ${itemId} has no ${field}.`);
  }
  return text;
}

function requiredBoolean(value: unknown, field: string, itemId: string): boolean {
  if (typeof value === "string") {
    value = value.trim().toLowerCase();
  }
  if (value === true || value === 1 || value === "1" || value === "true") {
    return true;
  }
  if (value === false || value === 0 || value === "0" || value === "false") {
    return false;
  }
  throw new Error(`Catalog item ${itemId} has an invalid ${field}.`);
}

function allowedEmails(value: unknown): string[] {
  const values = Array.isArray(value) ? value : value == null ? [] : [value];
  return values.flatMap((entry) => {
    if (typeof entry === "string") {
      const email = normalizeUserEmail(entry);
      return email ? [email] : [];
    }
    if (!entry || typeof entry !== "object") {
      return [];
    }
    const user = entry as CatalogRecord;
    const email =
      normalizeUserEmail(user.Email) ||
      normalizeUserEmail(user.email) ||
      normalizeUserEmail(user.UserPrincipalName) ||
      normalizeUserEmail(user.userPrincipalName);
    return email ? [email] : [];
  });
}

export function validateAndAuthorizeCatalog(
  records: unknown[],
  user: VerifiedUser,
): SafeCatalogAgent[] {
  return records.flatMap((raw, index) => {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
      throw new Error(`Catalog item ${index + 1} is invalid.`);
    }
    const record = raw as CatalogRecord;
    const itemId = optionalString(record.Id) || optionalString(record.id) || String(index + 1);
    if (!requiredBoolean(record.Enabled, "Enabled", itemId)) {
      return [];
    }

    const audienceValue = requiredString(record.Audience, "Audience", itemId).toLowerCase();
    if (audienceValue !== "everyone" && audienceValue !== "restricted") {
      throw new Error(`Catalog item ${itemId} has an invalid Audience.`);
    }
    const audience = audienceValue === "restricted" ? "Restricted" : "Everyone";
    const users = allowedEmails(record.AllowedUsers);
    if (audience === "Restricted" && users.length === 0) {
      throw new Error(`Catalog item ${itemId} is Restricted but has no valid AllowedUsers.`);
    }

    const environmentId = requiredString(record.EnvironmentId, "EnvironmentId", itemId);
    const schemaName = requiredString(record.SchemaName, "SchemaName", itemId);
    if (!environmentIdPattern.test(environmentId)) {
      throw new Error(`Catalog item ${itemId} has an invalid EnvironmentId.`);
    }
    if (!schemaNamePattern.test(schemaName)) {
      throw new Error(`Catalog item ${itemId} has an invalid SchemaName.`);
    }

    const endsConversation = requiredBoolean(record.EndsConversation, "EndsConversation", itemId);
    const safe: SafeCatalogAgent = {
      Id: itemId,
      Title: requiredString(record.Title, "Title", itemId),
      Description: optionalString(record.Description),
      EnvironmentId: environmentId,
      SchemaName: schemaName,
      Enabled: true,
      CompletionPhrase: endsConversation
        ? requiredString(record.CompletionPhrase, "CompletionPhrase", itemId)
        : optionalString(record.CompletionPhrase) || "",
      EndsConversation: endsConversation,
      WelcomeMessage: optionalString(record.WelcomeMessage),
      Locale: requiredString(record.Locale, "Locale", itemId),
      VoiceName: requiredString(record.VoiceName, "VoiceName", itemId),
      AvatarCharacter: requiredString(record.AvatarCharacter, "AvatarCharacter", itemId),
      AvatarStyle: requiredString(record.AvatarStyle, "AvatarStyle", itemId),
      Audience: audience,
      Harness: optionalString(record.Harness),
    };

    return audience === "Restricted" && (!user.email || !users.includes(user.email))
      ? []
      : [safe];
  });
}
