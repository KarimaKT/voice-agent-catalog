# Maker guide

This guide is for Copilot Studio makers, Power Platform makers, content owners,
and catalog administrators who configure agents without changing the React or
Azure code.

**Voice Agent Catalog is a shared front end, not a second agent.** The Teams
personal tab and standalone website load approved agents from the SharePoint
list. Copilot Studio owns the conversation, tools and permissions; the catalog
owns discovery, voice/avatar presentation and visibility.

## Maker quick start

1. Publish a **Standard-harness**, Microsoft-authenticated Copilot Studio agent
   and share it with your intended users.
2. Open the SharePoint maker page and choose **Register an agent**.
3. Paste the agent's Studio browser URL, for example
   `https://copilotstudio.microsoft.com/environments/<environment-guid>/agents/<agent-guid>`.
   Use your real URL, not this placeholder or a token/channel endpoint.
4. Save. Check `RegistrationStatus`; registration resolves technical metadata
   only in the configured environment and leaves the row disabled.
5. A catalog maintainer opens the row's **Edit** form, chooses a profile and
   audience using the exact values below, then enables the reviewed entry.
6. Users refresh/reload the catalog. No Teams package or application deployment
   is needed for list changes.

Use the advanced fields on **Edit**, not the single-field **New** form.
Defaults are choices for review, not information inferred from the agent URL.
Changing an existing URL does not rerun the item-created registration flow;
disable the row and have a maintainer revalidate its metadata.

## Table of contents

- [Build for chat; the catalog adapts the voice channel](#build-for-chat-the-catalog-adapts-the-voice-channel)
- [What makers control](#what-makers-control)
- [Import the sample agent](#import-the-sample-agent)
- [Configure Pat's Outlook action](#configure-pats-outlook-action)
- [Create another agent](#create-another-agent)
- [Add an agent to the catalog](#add-an-agent-to-the-catalog)
- [Choose a voice](#choose-a-voice)
- [Choose an avatar](#choose-an-avatar)
- [User controls versus list settings](#user-controls-versus-list-settings)
- [Browser microphone troubleshooting](#browser-microphone-troubleshooting)
- [Test before publishing](#test-before-publishing)
- [Operate the catalog](#operate-the-catalog)

## What makers control

Makers own:

- Copilot Studio instructions, topics, knowledge, tools, and sharing;
- agent welcome and completion contracts;
- Outlook or other connector actions;
- SharePoint catalog rows;
- locale, voice, avatar character, and avatar style; and
- catalog audience (`Everyone` or direct-user `Restricted`) and optional
  harness metadata; and
- publication of draft agent changes.

Changing a catalog row does not require rebuilding or reinstalling the Teams
application.

## Build for chat; the catalog adapts the voice channel

**Build a normal chat agent. You do not need a separate voice agent or a
voice-specific copy of its instructions.** Configure its business behavior,
knowledge, tools, confirmations and access in Copilot Studio as usual.
The web app knows it is providing the catalog's voice-capable experience.

The app starts the agent normally, displaying its native conversation-start
message (or the catalog fallback). At the beginning of the user's **first
typed or spoken text message**, the app prepends this interface context once:

```text
Remember the following across this conversation: This conversation is happening over a voice channel.
Keep all conversational responses short and natural, with no markup or decorative symbols.
The user will be speaking <selected spoken-input locale>, but you must always answer in <catalog Locale>.
Keep cards, citations, files and HTML tool payloads in their structured outputs; do not read them aloud.
Follow your existing instructions and authorization rules.
This interface context is not a request to execute a tool.
Here is the user's response to your request:

<the user's original first message, unchanged>
```

This is the catalog's voice-adaptation formula: concise, plain spoken replies
in the catalog output language, even when the user speaks another language.
For example, input `es-ES` with output `en-US` requests English answers to a
Spanish-speaking user. The UI shows only the user's original answer; it does
not clutter the transcript with this prefix. Later user text is sent unchanged.
There is no extra hidden setup message or additional setup model turn.

The prefix is a user-message presentation hint, **not a system-policy override
or a guarantee that any agent can be converted**. It cannot override safety,
sharing, connector permissions, fixed authored messages or unsupported harness
transports. Native greetings sent before the first user message still use their
authored content. Adaptive Card submissions retain their object-valued
`Activity.value`; the prefix waits for the first text message rather than
breaking card payloads. Rich cards and HTML emails remain structured content,
not something to read aloud.

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
- have a stable completion phrase if it is a terminating workflow;
- keep spoken responses concise;
- expose any required connector action to the orchestrator; and
- avoid claiming an action succeeded until the action confirms success.

Use `copilot-studio/order-agent-instructions.txt` as a second scenario example.

## Add an agent to the catalog

Use **Register an agent** on the maker page: paste the Studio agent URL and
save the single URL field. The metadata flow fills technical fields
and creates a disabled review draft with installer-selected profile defaults.
Makers do not need to copy environment IDs
or schema names manually. A catalog maintainer handles the advanced fields:

| Column | Exact values / what to enter | Effect and cautions |
| --- | --- | --- |
| `AgentUrl` (Copilot Studio agent URL) | Full Studio URL containing your environment and agent GUIDs | Registration input. Do not paste a secret, signed URL, published channel URL or Direct Line token endpoint. |
| `RegistrationStatus` | Generated by the flow; read, do not overwrite | Metadata lookup status, not proof of publication, sharing or runtime compatibility. |
| `Title` | Your human-readable name, e.g. `Pat` | Dropdown and speaker label. |
| `Description` | Short plain-text purpose | Visible subtitle; does not change agent instructions. |
| `EnvironmentId` | Generated environment GUID | Do not guess or copy from an unrelated environment. |
| `SchemaName` | Generated published schema name, e.g. `new_MyAgent` | Not the agent GUID or display name. |
| `Enabled` | `TRUE` / `FALSE` for a text column; Yes / No for a boolean column | Keep `FALSE` until review. Disabled rows are omitted. |
| `EndsConversation` | Yes / No (boolean); equivalent `true` / `false` | Use No for ongoing assistants. Use Yes only for a terminating workflow. |
| `CompletionPhrase` | The agent's full final success message | Required when `EndsConversation` is Yes. Matching normalizes case/whitespace but is whole-message, not substring matching. An agent EndOfConversation activity can also end the session. |
| `WelcomeMessage` | Plain-text fallback first prompt | Used only if the agent returns no first message; it does not replace a message the agent actually sends. |
| `Locale` | A locale code from a profile below, e.g. `en-US` | Requested output language and synthesis locale. The agent's own instructions still apply. |
| `VoiceName` | Exact `ShortName` from the voice reference, e.g. `en-US-AvaMultilingualNeural` | Paste the full string, including hyphens and capitalization. |
| `AvatarCharacter` | Exact lowercase character from the supported pairs below, e.g. `lisa` | Not a display name, photo filename or custom-avatar deployment ID. |
| `AvatarStyle` | A style paired with that character, e.g. `casual-sitting` | Not a voice speaking style. Never combine a character with another character's style. |
| `Audience` | `Everyone` or `Restricted` | Everyone means eligible signed-in users of this tenant, not anonymous/public access. |
| `AllowedUsers` | Select direct users in the people picker | Required for Restricted. Multiple people are supported; groups, pasted names and domain wildcards are not expanded. |
| `Harness` | `standard` for the working transport | `github-copilot` and unknown values are retained but cannot be invoked by this client. Do not relabel an incompatible agent as standard. |

See [catalog administration](../../catalog/README.md) and the sample
`catalog/Voice Agent Catalog.csv`.

`Locale` selects TTS/avatar synthesis and the output language requested by the
first-message prefix. Users choose a separate input language for STT, defaulting
to `Locale`. Keep the agent chat-first instead of baking one voice language into
its instructions; review any stricter language rules already authored in the agent.

Audience filtering improves the dropdown experience only. The app backend
passes a verified user identity to the maker flow, and enforces direct
`AllowedUsers` equality server-side. The flow checks SharePoint permissions and
returns `canManageCatalog` plus a configuration URL
for catalog/site administrators. **Copilot Studio sharing and runtime
authorization remain the final security boundary.**

The current Direct Line/Agents SDK live transport is verified only for the
standard Copilot Studio harness. GitHub Copilot/Copilot Chat harnesses do not
currently support this transport. Unknown values remain visible but disabled,
so future transports can be added without a list migration or accidental
invocation.

## Choose a voice

These exact GA voice identifiers were confirmed against the Azure Speech
West Europe voices-list API on **2026-10-10 UTC**, without synthesizing audio:

| Output language | `Locale` | `VoiceName` |
| --- | --- | --- |
| English (US), Ava | `en-US` | `en-US-AvaMultilingualNeural` |
| English (US), Andrew | `en-US` | `en-US-AndrewMultilingualNeural` |
| English (UK), Sonia | `en-GB` | `en-GB-SoniaNeural` |
| Spanish (Spain), Elvira | `es-ES` | `es-ES-ElviraNeural` |
| French (France), Denise | `fr-FR` | `fr-FR-DeniseNeural` |
| German (Germany), Katja | `de-DE` | `de-DE-KatjaNeural` |
| Italian (Italy), Elsa | `it-IT` | `it-IT-ElsaNeural` |
| Portuguese (Brazil), Francisca | `pt-BR` | `pt-BR-FranciscaNeural` |
| Hindi (India), Swara | `hi-IN` | `hi-IN-SwaraNeural` |
| Japanese (Japan), Nanami | `ja-JP` | `ja-JP-NanamiNeural` |
| Korean (Korea), SunHi | `ko-KR` | `ko-KR-SunHiNeural` |
| Chinese (Mandarin, Simplified), Xiaoxiao | `zh-CN` | `zh-CN-XiaoxiaoNeural` |

[Download the full voice-name reference CSV](speech-voices.csv): 958 identifiers
returned by that regional API, with `ShortName`, `Locale`, `Gender`, `VoiceType`
and `Status`. Copy `ShortName` into `VoiceName`. This is a dated availability
snapshot, not a claim that every entry is compatible with every avatar or has
the same price. Prefer `Status=GA`; preview, HD and specialist voices need
separate compatibility and pricing review. Availability can change with region.

The list has no SSML speaking-style, pitch, speed or voice-role columns.
Words such as `cheerful` are not valid substitutions for `VoiceName` or
`AvatarStyle`. The shell requests concise speech, but authored cards and agent
instructions remain in control of their content.

See Microsoft's complete
[language and voice support table](https://learn.microsoft.com/azure/ai-services/speech-service/language-support?tabs=tts).

## Choose an avatar

For this client's full-body, real-time Speech SDK path, enter these exact pairs:

| `AvatarCharacter` | `AvatarStyle` |
| --- | --- |
| `lisa` | `casual-sitting` |
| `harry` | `casual` |
| `harry` | `youthful` |
| `lori` | `casual` |
| `lori` | `graceful` |
| `lori` | `formal` |
| `max` | `business` |
| `max` | `casual` |
| `max` | `formal` |
| `meg` | `business` |
| `meg` | `casual` |
| `meg` | `formal` |

Character and style values are case-sensitive. Availability varies by Azure
region. Review the complete Microsoft lists for
[standard avatars](https://learn.microsoft.com/azure/ai-services/speech-service/text-to-speech-avatar/standard-avatars)
and [real-time synthesis](https://learn.microsoft.com/azure/ai-services/speech-service/text-to-speech-avatar/real-time-synthesis-avatar).

Microsoft's current table says Lisa's `graceful-sitting`, `graceful-standing`,
`technical-sitting` and `technical-standing` are **not supported by the
real-time API** and are deprecated for December 2026. Do not enter them here.
Jeff is also deprecated; do not use it for new registrations.

Newer entries whose style is shown as **N/A** (such as Rowan, Celine, Nia and
Malik) are not copy-ready profiles for this list: the current app requires a
nonempty style, and a display-table N/A is not an API identifier. Do not guess
`N/A`, `default` or `business` for those entries. Photo/Voice Live avatars,
custom avatar deployments and batch-only gestures are not configured by these
two columns. Region availability must be checked before enabling any new pair;
only the deployment's actually tested profiles should be called runtime-verified.

**Copy-ready baseline profile**

```text
Locale: en-US
VoiceName: en-US-AvaMultilingualNeural
AvatarCharacter: lisa
AvatarStyle: casual-sitting
Harness: standard
EndsConversation: No
Enabled: FALSE
```

Set Audience and sharing deliberately. For Pat's terminating email workflow,
use `EndsConversation=Yes` and the configured success message, for example
`Thank you. Your manager handoff summary has been emailed.`
Do not enable an agent merely because you have copied a profile.

## User controls versus list settings

- Text-only is the default. Users opt into **Read typed replies aloud**,
  **Start voice conversation**, or **Avatar video** after seeing cost estimates.
- Spoken input language is a user choice; it is independent of output `Locale`.
- **Avatar idle timeout** is a user control: `15`, `34` (default), or `45` seconds.
  It is not currently a SharePoint column. Agent work and playback pause the
  idle timer; connected waiting time is still billable.
- Idle expiry disconnects video only. It does not cancel the agent's work or
  business actions. Video can reconnect for a later reply. Stop, completion,
  avatar-off, errors and leaving/hiding the app release resources.
- Cost-notice wording comes from the host's `VOICE_COST_NOTICE` and
  `AVATAR_COST_NOTICE` app settings, not from list columns. An operator can
  change them without rebuilding Teams; saving settings restarts the host.
- Instructions, knowledge, tools, connector ownership and email recipients
  belong to Copilot Studio, not to the list.

## Browser microphone troubleshooting

1. Open the HTTPS app in a supported current browser. Choose
   **Sign in / Load catalog**, select an authorized agent, then **Start voice
   conversation**. A disabled Start button before catalog loading is normal.
2. In Edge/Chrome, open the site's address-bar permissions and allow
   **Microphone**, then reload. A global **Ask before accessing** setting is
   correct but does not grant permission to the individual site. A denied
   site/device policy cannot be fixed by the agent.
3. Select the intended headset in the browser microphone settings. In Windows,
   check **Settings > Privacy & security > Microphone**, including microphone
   access and permission for desktop apps such as Edge/Chrome.
4. For Teams web, also allow microphone access for the Teams site and follow
   the Teams app's device-permission prompt. Do not disable tenant policies.
   If embedded voice remains blocked, try the standalone app to distinguish
   host permissions from Speech/agent problems.
5. Keep the app visible while talking; hidden-page cleanup deliberately releases
   audio/video. Say a short answer and pause for automatic submission.
6. If text replies arrive but audio does not, check the output device, tab mute
   and browser playback permission. For typed replies, enable **Read typed
   replies aloud**. Avatar video is not required for voice.

If permission is managed by your organization, ask your device/browser
administrator rather than changing tenant-wide consent or Teams policies.
Use non-sensitive test input and a short audio-only check before paid avatar
testing. Do not infer successful microphone recognition from a typed interview.

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
9. verify restricted rows are hidden for a nonallowed user without treating
   that filtering as an authorization test.

For Pat, confirm exactly one formatted HTML email is delivered.

## Operate the catalog

- Edit a row to change voice or avatar without rebuilding the app.
- Set `Enabled` to false to hide an agent for new sessions.
- Publish and share an agent before enabling its row.
- Keep connection ownership and recipient policies documented.
- Never store secrets or signed URLs in catalog rows.
- Re-test Speech profiles after changing Azure region.

### Keep the SharePoint maker page current

The maker page should include this guide below its existing registration and
list web parts. Render this Markdown as HTML with your approved publishing
tools, retaining the tables, code examples and links. The reusable
[browser publisher](../../scripts/publish-maker-guide.browser.js) updates one
stable guide text web part using the signed-in site owner's session, an ETag
and a SharePoint request digest. It preserves existing controls, publishes and
reads back the result. Pass explicit site/page paths, item ID, title, owner
email and the rendered HTML; none belong in generic source.

Do not bypass sign-in, paste tokens into a page, or change tenant permissions
to publish documentation. If the owner session has expired, sign in normally
and retry. The voice reference links to the repository CSV; it contains only
public voice metadata, not customer data.
