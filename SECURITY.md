# Security

This repository is intended to remain safe for public viewing.

## Do not commit

- API keys
- access tokens
- webhook URLs
- OAuth client secrets
- spreadsheet credentials
- service-account private keys
- personal data

Runtime secrets belong in Google Apps Script **Script Properties**.

## If a secret is exposed

1. Revoke or rotate the exposed credential immediately.
2. Remove it from the working tree.
3. Check the repository history.
4. If necessary, rewrite the Git history and force-push the cleaned history.
5. Update the runtime configuration with the new credential.
