# Security Policy

## Reporting a Vulnerability

Please do not open a public issue for suspected security vulnerabilities.

Report security concerns privately to the repository owner through GitHub's private vulnerability reporting if enabled, or by contacting the maintainer directly.

## Sensitive Data

This project uses Microsoft device-code authentication through MSAL. Never commit or share:

- Microsoft access tokens or refresh tokens
- MSAL token cache files
- Authorization headers
- Private tenant, workspace, report, or semantic model identifiers unless intentionally redacted

The `powerbi_diagnostics` tool is designed to produce redacted troubleshooting output.
