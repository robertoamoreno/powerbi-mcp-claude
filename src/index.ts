#!/usr/bin/env node

import { AuthManager, TerminalDeviceCodePrompter } from "./auth.js";
import { authority, loadConfig } from "./config.js";
import { log } from "./logger.js";
import { LocalPowerBITools } from "./localTools.js";
import { RemoteMCPClient } from "./remote.js";
import { makeStdoutWriter, MCPProxyTransport } from "./transport.js";
import { VERSION } from "./version.js";

type ParsedArgs = {
  login: boolean;
  logout: boolean;
  status: boolean;
  help: boolean;
};

async function main(): Promise<number> {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printHelp();
    return 0;
  }

  const config = loadConfig();

  if (args.login || args.logout || args.status) {
    const auth = new AuthManager(config, new TerminalDeviceCodePrompter());

    if (args.logout) {
      const result = await auth.clearCache();
      process.stdout.write(`Cleared token cache at ${result.cachePath}\n`);
      process.stdout.write(`Accounts removed: ${result.accountsRemoved}\n`);
      process.stdout.write(`Remaining cached accounts: ${result.remainingAccounts}\n`);
      return 0;
    }

    if (args.status) {
      const info = await auth.tokenInfo();
      process.stdout.write(`Token cache: ${info.cachePath}\n`);
      process.stdout.write(`Token cache exists: ${info.cacheExists ? "yes" : "no"}\n`);
      process.stdout.write(`Cached accounts: ${info.accountCount}\n`);
      for (const username of info.usernames) {
        process.stdout.write(`- ${username}\n`);
      }
      return 0;
    }

    await auth.getAccessToken({ forceInteractive: true });
    process.stdout.write(`Authentication complete. Token cache: ${config.tokenCachePath}\n`);
    return 0;
  }

  let transport: MCPProxyTransport | undefined;
  const auth = new AuthManager(config, {
    supportsDeviceCodePrompt: () => transport?.supportsDeviceCodePrompt() ?? false,
    promptDeviceCode: async (response) => {
      if (!transport) {
        return "unavailable";
      }
      return await transport.promptDeviceCode(response);
    },
  });

  log(`Power BI MCP for Claude version ${VERSION}`);
  log(`Remote Power BI MCP URL: ${config.remoteUrl}`);
  log(`MSAL authority: ${authority(config)}`);
  log(`Token cache: ${config.tokenCachePath}`);
  log(
    `Default Power BI context: workspace=${config.defaultWorkspaceId ?? "not configured"}, semanticModel=${config.defaultSemanticModelId ?? "not configured"}`,
  );

  const remote = new RemoteMCPClient(config, auth);
  const localTools = new LocalPowerBITools(auth, config);
  transport = new MCPProxyTransport(remote, makeStdoutWriter(process.stdout), localTools);
  await transport.run(process.stdin);
  return 0;
}

function parseArgs(argv: string[]): ParsedArgs {
  return {
    login: argv.includes("--login"),
    logout: argv.includes("--logout"),
    status: argv.includes("--status"),
    help: argv.includes("--help") || argv.includes("-h"),
  };
}

function printHelp(): void {
  process.stdout.write(`Power BI MCP proxy for Claude

Usage:
  powerbi-mcp-claude            Run stdio MCP proxy
  powerbi-mcp-claude --login    Authenticate with Microsoft and cache tokens
  powerbi-mcp-claude --status   Show non-sensitive cache status
  powerbi-mcp-claude --logout   Delete local token cache

Environment:
  POWERBI_MCP_URL
  POWERBI_MCP_CLIENT_ID
  POWERBI_MCP_TENANT_ID
  POWERBI_MCP_SCOPES
  POWERBI_MCP_DEFAULT_WORKSPACE_ID
  POWERBI_MCP_DEFAULT_WORKSPACE_NAME
  POWERBI_MCP_DEFAULT_SEMANTIC_MODEL_ID
  POWERBI_MCP_DEFAULT_SEMANTIC_MODEL_NAME
  POWERBI_MCP_CACHE_DIR
  POWERBI_MCP_TIMEOUT_SECONDS
  POWERBI_MCP_ALLOW_INTERACTIVE_AUTH
`);
}

main()
  .then((code) => {
    process.exitCode = code;
  })
  .catch((error: unknown) => {
    log(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  });
