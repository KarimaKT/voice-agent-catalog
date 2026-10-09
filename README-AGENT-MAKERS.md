# Publish an agent to Voice Agent Catalog

1. Build and test a Copilot Studio agent with Microsoft authentication.
2. Choose an output `Locale`. The shell requests concise, voice-friendly replies
   in that language automatically; agent-authored rules and cards still apply.
3. Publish the agent and record its environment ID and schema name.
4. Share the agent with every intended runtime user or group. Catalog filtering
   restricts catalog visibility; **Copilot Studio sharing is the final invocation
   boundary**.
5. On the catalog's maker page, select **Register an agent**. Paste the full
   Copilot Studio agent page URL copied from your browser, then save. That is
   the only required new-item field.
6. The maker-owned metadata flow fills the name, environment, schema and
   harness. Voice/avatar use installer-selected defaults. Check
   `RegistrationStatus`; entries remain disabled until catalog review.
   A maintainer reviews sharing, reply language/audience, voice/avatar,
   transport and completion before
   enabling the entry. General agents use `EndsConversation=false`.
7. Test typed input, each supported spoken-input language, TTS/avatar output,
   rich responses, completion behavior, and access as a nonmaker.

Set `Harness=standard` for a usable entry. The current live Direct Line/Agents SDK
transport is verified only for the standard Copilot Studio harness. GitHub
Copilot/Copilot Chat harnesses do not currently support that transport. Future
harness values are preserved but labelled unavailable in the picker.

See [the full maker guide](docs/maker/README.md) and
[catalog setup](catalog/README.md). This repository intentionally contains no
tenant-specific list URL, IDs, or user addresses.

The installer must configure the URL-registration flow and an authorized
metadata connection for the target environment. Other environments are not
guessed from a URL. If resolution fails, the row shows a needs-attention status
and stays disabled; a maintainer can use the advanced fields.
