import { closeSync, openSync, unlinkSync } from "node:fs";

const DEFAULT_TIMEOUT_MS = 5000;
const DEFAULT_INTERVAL_MS = 50;

export async function withFileLock<T>(lockPath: string, fn: () => Promise<T>): Promise<T> {
  const fd = await acquireLock(lockPath);
  try {
    return await fn();
  } finally {
    closeSync(fd);
    try {
      unlinkSync(lockPath);
    } catch {
      // Lock cleanup is best-effort. A later process will fail fast if the lock stays stale.
    }
  }
}

async function acquireLock(lockPath: string): Promise<number> {
  const started = Date.now();

  while (true) {
    try {
      return openSync(lockPath, "wx", 0o600);
    } catch (error) {
      if (!isFileExistsError(error) || Date.now() - started > DEFAULT_TIMEOUT_MS) {
        throw error;
      }
      await sleep(DEFAULT_INTERVAL_MS);
    }
  }
}

function isFileExistsError(error: unknown): boolean {
  return typeof error === "object" && error !== null && "code" in error && error.code === "EEXIST";
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
