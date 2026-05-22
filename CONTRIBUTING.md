# Contributing

Thanks for considering a contribution.

## Development

```bash
npm install
npm run check
npm test
```

Package a Claude Desktop extension with:

```bash
npm run pack:mcpb
```

## Pull Requests

- Keep Power BI data access routed through `https://api.fabric.microsoft.com/v1/mcp/powerbi`.
- Do not add direct `api.powerbi.com` calls.
- Do not commit Microsoft access tokens, refresh tokens, local token caches, or generated `.mcpb` bundles.
- Add or update tests for changes to auth, context defaults, tool mapping, or diagnostics.
- Run `npm run check`, `npm test`, and `npm audit --audit-level=low` before opening a PR.

## Security

Do not paste secrets or private tenant data into issues or pull requests. Use redacted diagnostics output where possible.
