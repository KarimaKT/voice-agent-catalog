# Security

This repository is a reference sample and still requires customer security
review before production use.

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

## API authorization

`/api/catalog` and `/api/speech/token` require the app's delegated
`access_as_user` bearer token. The server validates Entra signature, issuer,
tenant, audience, expiry, and scope. `/api/health` and `/api/config` are public
and return no secrets. Catalog authorization uses only verified claims; client
headers are not identity evidence.

Before production, add distributed rate limiting, formal data retention,
security monitoring, privacy review, accessibility testing, threat modeling,
and penetration testing.
