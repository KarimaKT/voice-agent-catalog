# Copilot Studio setup

This folder is the source-controlled, tenant-neutral definition of the Pat
interview pattern. It intentionally contains no tenant IDs, email addresses,
connector connection references, environment URLs, or exported credentials.

## Required first edit

Before publishing the agent, decide who should receive the manager handoff
email. Replace `<MANAGER_OR_SHARED_MAILBOX_ADDRESS>` while configuring the
Office 365 Outlook action. Do not leave a sample author's address in an exported
solution or source-controlled action.

For a reusable deployment, prefer one of these recipient policies:

1. A deployment-owned shared mailbox configured during installation.
2. An environment variable resolved by a flow or action.
3. An authenticated manager address looked up from an approved directory
   source.
4. A maker-configured fixed address for a short-lived POC.

Do not let interviewees supply an arbitrary recipient unless that behavior is
intentional and protected against abuse.

## Create the agent

1. Create a standard Copilot Studio agent in the target Power Platform
   environment.
2. Name it for the local deployment.
3. Select Microsoft authentication.
4. Enable generative orchestration if required by the interview design.
5. Paste the contents of `agent-instructions.txt` into the agent instructions.
6. Add **Office 365 Outlook – Send an email (V2)**.
7. Configure the action as described in `outlook-action.md`.
8. Test all four interview topics and one insufficient-answer follow-up.
9. Confirm one email is sent and the final completion sentence is exact.
10. Publish the agent.
11. Record the environment ID and agent schema name in the local `.env.dev`
    file used by the Teams project.

## Sharing

Share the published agent with the same users or Entra security group that can
install the Teams application. End users need permission to use the agent but
do not need coauthor access.

## Power Platform solution packaging

For organizational reuse, add the agent, connection reference, and related
components to a dedicated unmanaged Power Platform solution. Export the
solution without live credentials or environment-specific recipient values.
On import:

1. Select or create the target Outlook connection.
2. Configure the recipient policy.
3. publish the imported agent;
4. record its environment ID and schema name; and
5. provision the Teams application with those values.

Connection authorization is always completed in the destination environment;
credentials are never included in source control.
