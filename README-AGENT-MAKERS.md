# Publish an agent to Voice Agent Catalog

1. Build and test a Copilot Studio agent with Microsoft authentication.
2. Keep responses concise for speech. Make the agent always answer in the
   catalog output `Locale`, even when input is in another language.
3. Publish the agent and record its environment ID and schema name.
4. Share the agent with every intended runtime user or group. Catalog filtering
   is only dropdown UX; **Copilot Studio sharing is the final security
   boundary**.
5. In the SharePoint **Voice Agent Catalog** list, select **New** and enter:
   identity, welcome/completion text, output `Locale`, voice, avatar, and style.
6. Choose `Audience=Everyone`, or `Restricted` and select direct people in
   `AllowedUsers`. Enable the row only after sharing and publication are ready.
7. Test typed input, each supported spoken-input language, TTS/avatar output,
   rich responses, completion behavior, and access as a nonmaker.

`Harness` is optional metadata. The current live Direct Line/Agents SDK
transport is verified only for the standard Copilot Studio harness. GitHub
Copilot/Copilot Chat harnesses do not currently support that transport. Future
harness values are preserved and are not rejected by the catalog schema.

See [the full maker guide](docs/maker/README.md) and
[catalog setup](catalog/README.md). This repository intentionally contains no
tenant-specific list URL, IDs, or user addresses.
