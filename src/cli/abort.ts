export class CommandAbortedError extends Error {
  constructor(message = "Command aborted") {
    super(message);
    this.name = "CommandAbortedError";
  }
}

const signalTimeouts = new WeakMap<
  AbortSignal,
  { deadlineAt: number; message: string }
>();

function reasonToMessage(reason: unknown): string {
  if (reason instanceof Error) return reason.message;
  if (typeof reason === "string" && reason.trim()) return reason;
  return "Command aborted";
}

export function registerSignalTimeout(
  signal: AbortSignal,
  timeoutMs: number,
  message?: string,
): void {
  signalTimeouts.set(signal, {
    deadlineAt: Date.now() + timeoutMs,
    message:
      message ?? `Command timed out after ${Math.ceil(timeoutMs / 1000)}s`,
  });
}

export function throwIfAborted(signal?: AbortSignal | null): void {
  if (!signal) return;

  if (signal.aborted) {
    throw new CommandAbortedError(reasonToMessage(signal.reason));
  }

  const timeout = signalTimeouts.get(signal);
  if (timeout && Date.now() >= timeout.deadlineAt) {
    throw new CommandAbortedError(timeout.message);
  }
}

export function isAbortError(err: unknown): boolean {
  return err instanceof CommandAbortedError;
}
