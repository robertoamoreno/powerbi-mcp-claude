export function log(message: string): void {
  process.stderr.write(`[${new Date().toISOString()}] ${message}\n`);
}

export function logBlock(message: string): void {
  for (const line of message.split(/\r?\n/)) {
    if (line.trim()) {
      log(line);
    }
  }
}
