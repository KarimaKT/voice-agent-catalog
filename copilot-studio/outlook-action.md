# Outlook action configuration

Add **Office 365 Outlook – Send an email (V2)** to the agent.

## Inputs

### To

Set the installation-specific recipient:

```text
<MANAGER_OR_SHARED_MAILBOX_ADDRESS>
```

The recipient is intentionally not stored in this repository. Configure it
during installation or resolve it through an approved environment-specific
mechanism.

### Subject

Have the agent generate a concise subject such as:

```text
Manager handoff - <interviewee name> - <role>
```

### Body

Have the agent generate a complete HTML fragment or document with:

- a clear manager-handoff title;
- interviewee name and role;
- most important responsibility;
- key process, document, or system;
- open item, risk, or recommendation; and
- simple inline styles supported by Outlook.

## Connection mode

For a small controlled POC, a maker-owned connection lets approved users run the
action without creating individual Outlook connections. Review mailbox,
connector-sharing, auditing, and data-handling requirements before broader use.

Use user-owned connections when each interviewee must send as themselves.

## Completion contract

The action must run exactly once after all topics are sufficient. Pat must wait
for the result and only use the deterministic success sentence after confirmed
delivery:

```text
Thank you. Your manager handoff summary has been emailed.
```
