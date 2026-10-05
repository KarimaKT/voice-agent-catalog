# Maker guide

This guide is for Copilot Studio makers, Power Platform makers, content owners,
and catalog administrators who configure agents without changing the React or
Azure code.

## Table of contents

- [What makers control](#what-makers-control)
- [Import the sample agent](#import-the-sample-agent)
- [Configure Pat's Outlook action](#configure-pats-outlook-action)
- [Create another agent](#create-another-agent)
- [Add an agent to the catalog](#add-an-agent-to-the-catalog)
- [Choose a voice](#choose-a-voice)
- [Choose an avatar](#choose-an-avatar)
- [Test before publishing](#test-before-publishing)
- [Operate the catalog](#operate-the-catalog)

## What makers control

Makers own:

- Copilot Studio instructions, topics, knowledge, tools, and sharing;
- agent welcome and completion contracts;
- Outlook or other connector actions;
- SharePoint catalog rows;
- locale, voice, avatar character, and avatar style; and
- publication of draft agent changes.

Changing a catalog row does not require rebuilding or reinstalling the Teams
application.

## Import the sample agent

Use the real Power Platform solution package:

```text
copilot-studio/packages/pat-manager-handoff-solution.zip
```

[Download the sample solution ZIP](../../copilot-studio/packages/pat-manager-handoff-solution.zip).

Import it from **Solutions** in the target Power Platform environment. The
package creates **Pat Manager Handoff Sample** with the four-topic interview
instructions and standard system topics.

The package intentionally omits live connector connections and recipient
values. After import, add and configure the Outlook action before publishing.
The editable package source is under
`copilot-studio/sample-agent/solution-source/`.

Web browsing, file analysis, semantic search, and general model knowledge are
disabled in the sample so the manager handoff is based only on interview
answers. Enable additional capabilities only after reviewing the data and
grounding requirements for the destination scenario.

## Configure Pat's Outlook action

Follow `copilot-studio/outlook-action.md`:

1. Add **Office 365 Outlook - Send an email (V2)**.
2. Select the approved maker-owned or user-owned connection policy.
3. Set the installation-specific manager or shared-mailbox recipient.
4. Let Pat generate the subject and Outlook-compatible HTML body.
5. Ensure the action runs once, only after all four topics are sufficient.
6. Use the success completion phrase only after the connector succeeds.

Never embed a sample author's address in an exported solution.

## Create another agent

An agent can implement any compatible conversation. It must:

- use Microsoft authentication for production Copilot invocation;
- be published and shared with intended users;
- have a stable completion phrase;
- keep spoken responses concise;
- expose any required connector action to the orchestrator; and
- avoid claiming an action succeeded until the action confirms success.

Use `copilot-studio/order-agent-instructions.txt` as a second scenario example.

## Add an agent to the catalog

Add one enabled row to the SharePoint catalog:

| Column | Purpose |
| --- | --- |
| `Title` | Dropdown display name |
| `Description` | Short visible purpose |
| `EnvironmentId` | Published Copilot Studio environment GUID |
| `SchemaName` | Published agent schema name |
| `Enabled` | Whether the app displays the agent |
| `CompletionPhrase` | Exact successful completion sentence |
| `WelcomeMessage` | Fallback first message |
| `Locale` | Speech recognition and synthesis locale |
| `VoiceName` | Azure neural voice name |
| `AvatarCharacter` | Azure standard avatar character |
| `AvatarStyle` | Style supported by the selected character |

See [catalog administration](../../catalog/README.md) and the sample
`catalog/Pat Agent Catalog.csv`.

## Choose a voice

These are useful English (US) examples:

| Voice name | Voice | Type | Available speaking styles |
| --- | --- | --- | --- |
| `en-US-AvaMultilingualNeural` | Ava (female) | Multilingual | Default |
| `en-US-AndrewMultilingualNeural` | Andrew (male) | Multilingual | `empathetic`, `relieved` |
| `en-US-JennyNeural` | Jenny (female) | Standard neural | `assistant`, `chat`, `cheerful`, `customerservice`, `friendly`, `newscast`, and more |
| `en-US-GuyNeural` | Guy (male) | Standard neural | `cheerful`, `friendly`, `newscast`, `whispering`, and more |
| `en-US-AriaNeural` | Aria (female) | Standard neural | `chat`, `customerservice`, `empathetic`, `narration-professional`, and more |

The current catalog selects the voice name. Optional speaking styles require
SSML support in the synthesis path and are not a separate catalog field.

See Microsoft's complete
[language and voice support table](https://learn.microsoft.com/azure/ai-services/speech-service/language-support?tabs=tts).

## Choose an avatar

| Character | Example style values |
| --- | --- |
| `harry` | `casual`, `youthful` |
| `lisa` | `casual-sitting`, `graceful-sitting`, `graceful-standing`, `technical-sitting`, `technical-standing` |
| `lori` | `casual`, `graceful`, `formal` |
| `max` | `business`, `casual`, `formal` |
| `meg` | `business`, `casual`, `formal` |

Character and style values are case-sensitive. Availability varies by Azure
region. Review the complete Microsoft lists for
[standard avatars](https://learn.microsoft.com/azure/ai-services/speech-service/text-to-speech-avatar/standard-avatars)
and [real-time synthesis](https://learn.microsoft.com/azure/ai-services/speech-service/text-to-speech-avatar/real-time-synthesis-avatar).

## Test before publishing

For every agent:

1. test a normal answer;
2. test a vague answer and one concise follow-up;
3. test correction, skip, stop, and restart behavior;
4. test action failure explicitly;
5. confirm the completion phrase is emitted only on success;
6. verify spoken responses are short enough for voice use;
7. publish the draft; and
8. test as a nonauthor user who has only runtime access.

For Pat, confirm exactly one formatted HTML email is delivered.

## Operate the catalog

- Edit a row to change voice or avatar without rebuilding the app.
- Set `Enabled` to false to hide an agent for new sessions.
- Publish and share an agent before enabling its row.
- Keep connection ownership and recipient policies documented.
- Never store secrets or signed URLs in catalog rows.
- Re-test Speech profiles after changing Azure region.
