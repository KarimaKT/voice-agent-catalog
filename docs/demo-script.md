# Three-minute customer demo

## Before the call

1. Open **Voice Agentsdev** in Teams.
2. Allow microphone access.
3. Confirm **Avatar video** is on.
4. Close any previous conversation so the app starts cleanly.

## Part 1: Pat manager handoff

1. Select **Pat**, then **Start demo with Pat**.
2. Point out Lisa's avatar and Ava's voice.
3. Select **Start voice conversation** once.
4. Give these short answers, pausing after each:
   - "I'm Karima, a solution architect."
   - "My priority is preparing the customer voice-agent demonstration."
   - "I need help confirming the production authentication approach."
   - "The next step is to validate the two-agent catalog and hand it over."
5. Show the organized handoff summary.
6. Select **Open self-addressed email**. Outlook opens a message from the
   current Teams user to the same user. Select **Send**.

Demo mode intentionally opens an Outlook draft rather than storing Outlook
credentials in the sample.

## Part 2: Morgan order management

1. Select **Morgan** from the dropdown, then **Start demo with Morgan**.
2. Point out the switch to Harry's avatar and Andrew's voice.
3. Type or say: "Show pending orders."
4. Type or say: "Find order ORD-1002."
5. Type or say:
   "Place an order for 4 Surface Laptop 7 for Contoso Retail."
6. Show the generated order number and pending status.

The deployed POC uses a deterministic session tool so the demo remains
available without tenant consent. The customer architecture uses the
**Voice Agent Orders** SharePoint list and the included Morgan instructions as
the starting point for a maker-owned connector action.

## Close

Explain that agent name, description, voice, and avatar come from the catalog.
The Teams shell does not need to be rebuilt when a customer adds another
compatible agent.
