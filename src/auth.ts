import {
  chmodSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { basename, dirname, join } from "node:path";
import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration,
  type DeviceCodeRequest,
} from "@azure/msal-node";
import type { DeviceCodeResponse } from "@azure/msal-common/node";

import { authority, type ProxyConfig } from "./config.js";
import { withFileLock } from "./fileLock.js";
import { log, logBlock } from "./logger.js";

export class AuthenticationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "AuthenticationError";
  }
}

export type TokenInfo = {
  cachePath: string;
  cacheExists: boolean;
  accountCount: number;
  usernames: string[];
};

export type ClearCacheResult = {
  cachePath: string;
  cacheExists: boolean;
  accountsRemoved: number;
  remainingAccounts: number;
};

export type DeviceLoginStart =
  | {
      status: "authenticated";
      message: string;
    }
  | {
      status: "pending";
      userCode: string;
      verificationUri: string;
      expiresIn: number;
      message: string;
    };

export type DeviceLoginStatus =
  | {
      status: "not_started";
      authenticated: boolean;
      message: string;
    }
  | {
      status: "pending";
      authenticated: false;
      userCode: string;
      verificationUri: string;
      message: string;
    }
  | {
      status: "authenticated";
      authenticated: true;
      username?: string;
      message: string;
    }
  | {
      status: "failed";
      authenticated: false;
      message: string;
    };

export type DeviceCodePromptResult = "accepted" | "declined" | "cancelled" | "unavailable";

export interface DeviceCodePrompter {
  supportsDeviceCodePrompt(): boolean;
  promptDeviceCode(response: DeviceCodeResponse): Promise<DeviceCodePromptResult>;
}

export class TerminalDeviceCodePrompter implements DeviceCodePrompter {
  supportsDeviceCodePrompt(): boolean {
    return true;
  }

  async promptDeviceCode(response: DeviceCodeResponse): Promise<DeviceCodePromptResult> {
    logBlock(response.message);
    return "accepted";
  }
}

export class NullDeviceCodePrompter implements DeviceCodePrompter {
  supportsDeviceCodePrompt(): boolean {
    return false;
  }

  async promptDeviceCode(): Promise<DeviceCodePromptResult> {
    return "unavailable";
  }
}

export class AuthManager {
  private readonly app: PublicClientApplication;
  private activeTokenRequest: Promise<string> | undefined;
  private activeDeviceLogin: ActiveDeviceLogin | undefined;

  constructor(
    private readonly config: ProxyConfig,
    private readonly prompter: DeviceCodePrompter = new NullDeviceCodePrompter(),
  ) {
    const msalConfig: Configuration = {
      auth: {
        clientId: config.clientId,
        authority: authority(config),
      },
      system: {
        loggerOptions: {
          loggerCallback: (_level, message, containsPii) => {
            if (!containsPii && process.env.POWERBI_MCP_MSAL_DEBUG === "true") {
              log(`MSAL: ${message}`);
            }
          },
          piiLoggingEnabled: false,
        },
      },
    };

    this.app = new PublicClientApplication(msalConfig);
  }

  async getAccessToken(options: { forceRefresh?: boolean; forceInteractive?: boolean } = {}): Promise<string> {
    if (this.activeTokenRequest) {
      return this.activeTokenRequest;
    }

    this.activeTokenRequest = this.getAccessTokenOnce(options);
    try {
      return await this.activeTokenRequest;
    } finally {
      this.activeTokenRequest = undefined;
    }
  }

  async getCachedAccessToken(options: { forceRefresh?: boolean } = {}): Promise<string | undefined> {
    return await this.acquireSilent(options.forceRefresh === true);
  }

  async startDeviceCodeLogin(): Promise<DeviceLoginStart> {
    const cached = await this.getCachedAccessToken();
    if (cached) {
      return {
        status: "authenticated",
        message: "Microsoft authentication is already available.",
      };
    }

    if (this.activeDeviceLogin?.status === "pending") {
      return deviceLoginStartFromResponse(await this.activeDeviceLogin.started);
    }

    await this.loadCache();
    let resolveStarted: (response: DeviceCodeResponse) => void;
    let rejectStarted!: (error: Error) => void;
    const started = new Promise<DeviceCodeResponse>((resolve, reject) => {
      resolveStarted = resolve;
      rejectStarted = reject;
    });

    let login!: ActiveDeviceLogin;
    const request: DeviceCodeRequest = {
      scopes: this.config.scopes,
      timeout: this.config.deviceCodeTimeoutSeconds,
      deviceCodeCallback: (response: DeviceCodeResponse) => {
        login.response = response;
        resolveStarted(response);
      },
    };

    login = {
      status: "pending",
      started,
      request,
      response: undefined,
      error: undefined,
      username: undefined,
    };
    this.activeDeviceLogin = login;

    void this.completeBackgroundDeviceLogin(request, login, rejectStarted);

    const response = await started;
    return deviceLoginStartFromResponse(response);
  }

  async deviceLoginStatus(): Promise<DeviceLoginStatus> {
    const cached = await this.getCachedAccessToken();
    if (cached) {
      const accounts = await this.accounts();
      const account = accounts[0];
      const username = account ? accountLabel(account) : undefined;
      return {
        status: "authenticated",
        authenticated: true,
        ...(username ? { username } : {}),
        message: username ? `Authenticated with Microsoft as ${username}.` : "Authenticated with Microsoft.",
      };
    }

    if (!this.activeDeviceLogin) {
      return {
        status: "not_started",
        authenticated: false,
        message: "Microsoft authentication has not been started.",
      };
    }

    if (this.activeDeviceLogin.status === "pending" && this.activeDeviceLogin.response) {
      return {
        status: "pending",
        authenticated: false,
        userCode: this.activeDeviceLogin.response.userCode,
        verificationUri: this.activeDeviceLogin.response.verificationUri,
        message: `Still waiting for Microsoft sign-in. Open ${this.activeDeviceLogin.response.verificationUri} and enter code ${this.activeDeviceLogin.response.userCode}.`,
      };
    }

    if (this.activeDeviceLogin.status === "authenticated") {
      return {
        status: "not_started",
        authenticated: false,
        message: "No usable cached Microsoft token is available.",
      };
    }

    return {
      status: "failed",
      authenticated: false,
      message: this.activeDeviceLogin.error ?? "Microsoft authentication failed.",
    };
  }

  async clearCache(): Promise<ClearCacheResult> {
    if (this.activeDeviceLogin?.status === "pending") {
      this.activeDeviceLogin.request.cancel = true;
    }

    await this.loadCache();
    const accounts = await this.app.getTokenCache().getAllAccounts();
    for (const account of accounts) {
      await this.app.getTokenCache().removeAccount(account);
    }

    this.activeDeviceLogin = undefined;
    this.app.getTokenCache().deserialize("{}");
    await this.writeEmptyCache();
    this.removeCacheArtifacts();

    this.app.getTokenCache().deserialize(readFileSync(this.config.tokenCachePath, "utf8"));
    const remainingAccounts = await this.app.getTokenCache().getAllAccounts();

    return {
      cachePath: this.config.tokenCachePath,
      cacheExists: existsSync(this.config.tokenCachePath),
      accountsRemoved: accounts.length,
      remainingAccounts: remainingAccounts.length,
    };
  }

  async tokenInfo(): Promise<TokenInfo> {
    const accounts = await this.accounts();

    return {
      cachePath: this.config.tokenCachePath,
      cacheExists: existsSync(this.config.tokenCachePath),
      accountCount: accounts.length,
      usernames: accounts.map((account) => account.username).filter(Boolean),
    };
  }

  private async getAccessTokenOnce(options: { forceRefresh?: boolean; forceInteractive?: boolean }): Promise<string> {
    if (!options.forceInteractive) {
      const cached = await this.acquireSilent(options.forceRefresh === true);
      if (cached) {
        return cached;
      }
    }

    if (!this.config.allowInteractiveAuth) {
      throw new AuthenticationError(
        "No cached Microsoft token is available and interactive auth is disabled. Run the login helper first, then reconnect this MCP server.",
      );
    }

    if (!this.prompter.supportsDeviceCodePrompt()) {
      throw new AuthenticationError(
        "No cached Microsoft token is available. This Claude client did not advertise URL-mode MCP elicitation, so the server cannot show an in-app Microsoft sign-in prompt. Run `powerbi-mcp-claude --login` once, then reconnect.",
      );
    }

    return await this.acquireDeviceCode();
  }

  private async acquireSilent(forceRefresh: boolean): Promise<string | undefined> {
    await this.loadCache();
    const accounts = await this.app.getTokenCache().getAllAccounts();

    for (const account of accounts) {
      try {
        const result = await this.app.acquireTokenSilent({
          account,
          scopes: this.config.scopes,
          forceRefresh,
        });
        await this.persistCache();

        if (result.accessToken) {
          log(`Using cached Microsoft token for ${accountLabel(account)}.`);
          return result.accessToken;
        }
      } catch (error) {
        log(`Silent token acquisition failed for ${accountLabel(account)}: ${errorMessage(error)}`);
      }
    }

    return undefined;
  }

  private async accounts(): Promise<AccountInfo[]> {
    await this.loadCache();
    return await this.app.getTokenCache().getAllAccounts();
  }

  private async acquireDeviceCode(): Promise<string> {
    await this.loadCache();
    log("Microsoft authentication is required for Power BI MCP access.");

    const request: DeviceCodeRequest = {
      scopes: this.config.scopes,
      timeout: this.config.deviceCodeTimeoutSeconds,
      deviceCodeCallback: (response: DeviceCodeResponse) => {
        void this.handleDeviceCodePrompt(response, request);
      },
    };

    let result: AuthenticationResult | null;
    try {
      result = await this.app.acquireTokenByDeviceCode(request);
    } catch (error) {
      await this.persistCache();
      throw new AuthenticationError(`Microsoft device-code authentication failed: ${errorMessage(error)}`);
    }

    await this.persistCache();
    if (!result?.accessToken) {
      throw new AuthenticationError("Microsoft device-code authentication completed without an access token.");
    }

    log(`Authenticated with Microsoft${result.account ? ` as ${accountLabel(result.account)}` : ""}.`);
    return result.accessToken;
  }

  private async completeBackgroundDeviceLogin(
    request: DeviceCodeRequest,
    login: ActiveDeviceLogin,
    rejectStarted: (error: Error) => void,
  ): Promise<void> {
    try {
      const result = await this.app.acquireTokenByDeviceCode(request);
      await this.persistCache();
      if (!result?.accessToken) {
        throw new AuthenticationError("Microsoft device-code authentication completed without an access token.");
      }

      login.status = "authenticated";
      login.username = result.account ? accountLabel(result.account) : undefined;
      log(`Authenticated with Microsoft${login.username ? ` as ${login.username}` : ""}.`);
    } catch (error) {
      const message = errorMessage(error);
      if (!login.response) {
        rejectStarted(new AuthenticationError(`Could not start Microsoft device-code authentication: ${message}`));
      }
      login.status = "failed";
      login.error = message;
      await this.persistCache();
      log(`Microsoft device-code authentication failed: ${message}`);
    }
  }

  private async handleDeviceCodePrompt(response: DeviceCodeResponse, request: DeviceCodeRequest): Promise<void> {
    try {
      const result = await this.prompter.promptDeviceCode(response);
      if (result === "declined" || result === "cancelled" || result === "unavailable") {
        request.cancel = true;
      }
    } catch (error) {
      request.cancel = true;
      log(`Could not show device-code prompt: ${errorMessage(error)}`);
    }
  }

  private async loadCache(): Promise<void> {
    const lockPath = `${this.config.tokenCachePath}.lock`;
    mkdirSync(dirname(this.config.tokenCachePath), { recursive: true });

    await withFileLock(lockPath, async () => {
      if (!existsSync(this.config.tokenCachePath)) {
        this.app.getTokenCache().deserialize("{}");
        return;
      }

      try {
        this.app.getTokenCache().deserialize(readFileSync(this.config.tokenCachePath, "utf8"));
      } catch (error) {
        log(`Warning: could not load token cache ${this.config.tokenCachePath}: ${errorMessage(error)}`);
        this.app.getTokenCache().deserialize("{}");
      }
    });
  }

  private async persistCache(): Promise<void> {
    const cache = this.app.getTokenCache();
    if (!cache.hasChanged()) {
      return;
    }

    const lockPath = `${this.config.tokenCachePath}.lock`;
    mkdirSync(dirname(this.config.tokenCachePath), { recursive: true });

    await withFileLock(lockPath, async () => {
      const tmpPath = `${this.config.tokenCachePath}.${process.pid}.tmp`;
      writeFileSync(tmpPath, cache.serialize(), { encoding: "utf8", mode: 0o600 });
      chmodSync(tmpPath, 0o600);
      renameSync(tmpPath, this.config.tokenCachePath);
      chmodSync(this.config.tokenCachePath, 0o600);
    });
  }

  private async writeEmptyCache(): Promise<void> {
    const lockPath = `${this.config.tokenCachePath}.lock`;
    mkdirSync(dirname(this.config.tokenCachePath), { recursive: true });

    await withFileLock(lockPath, async () => {
      writeFileSync(this.config.tokenCachePath, "{}", { encoding: "utf8", mode: 0o600 });
      chmodSync(this.config.tokenCachePath, 0o600);
    });
  }

  private removeCacheArtifacts(): void {
    const cacheDir = dirname(this.config.tokenCachePath);
    const cacheName = basename(this.config.tokenCachePath);

    try {
      for (const entry of readdirSync(cacheDir)) {
        if (entry === `${cacheName}.lock` || (entry.startsWith(`${cacheName}.`) && entry.endsWith(".tmp"))) {
          rmSync(join(cacheDir, entry), { force: true });
        }
      }
    } catch {
      // Best effort cleanup only; the empty cache file above is the source of truth.
    }
  }
}

type ActiveDeviceLogin = {
  status: "pending" | "authenticated" | "failed";
  started: Promise<DeviceCodeResponse>;
  request: DeviceCodeRequest;
  response: DeviceCodeResponse | undefined;
  error: string | undefined;
  username: string | undefined;
};

function deviceLoginStartFromResponse(response: DeviceCodeResponse): DeviceLoginStart {
  return {
    status: "pending",
    userCode: response.userCode,
    verificationUri: response.verificationUri,
    expiresIn: response.expiresIn,
    message: `Open ${response.verificationUri} and enter code ${response.userCode} to authenticate with Microsoft.`,
  };
}

function accountLabel(account: Pick<AccountInfo, "username" | "homeAccountId">): string {
  return account.username || account.homeAccountId || "cached account";
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
