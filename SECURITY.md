# Security

This repository is a proof-of-concept sample, not a production service.

## Reporting

Report suspected vulnerabilities privately to the repository owner. Do not open
a public issue containing credentials, tenant identifiers, customer data, or
reproduction details that expose a deployed environment.

## Never commit

- customer emails or user lists;
- tenant, subscription, environment, agent, or connection-reference IDs;
- generated Teams packages;
- `env/.env.*` deployment files other than `.example` templates;
- `.azure/` deployment records;
- Outlook, Graph, Power Platform, or Azure credentials;
- Speech keys, access tokens, refresh tokens, or MFA material; or
- conversation transcripts and generated handoff content.

## Production hardening

Before production, add backend authentication to the Speech token route,
distributed rate limiting, formal data retention, security monitoring, privacy
review, accessibility testing, threat modeling, and penetration testing.
