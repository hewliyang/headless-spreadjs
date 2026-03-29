import { createRequire } from "node:module";
import { discoverHooks, setHookWriters } from "../hooks.js";
import { registerSignalTimeout } from "./abort.js";
import { spawnDaemon, tryDaemon, tryExistingDaemon } from "./client.js";
import { dispatch, USAGE } from "./dispatch.js";
import {
  createIoContext,
  runWithIo,
  writeStderr,
  writeStdout,
} from "./output.js";

const DEFAULT_TIMEOUT_MS = 30_000;

function hasFlag(args: string[], name: string): boolean {
  return args.includes(name);
}

function parseTimeoutValue(raw: string): number {
  const value = raw.trim();
  if (!/^\d+(\.\d+)?$/.test(value)) {
    throw new Error("Invalid --timeout value (expected seconds)");
  }

  const seconds = Number.parseFloat(value);
  if (!Number.isFinite(seconds) || seconds <= 0) {
    throw new Error("Invalid --timeout value (expected seconds)");
  }
  return Math.floor(seconds * 1000);
}

function parseGlobalOptions(args: string[]): {
  args: string[];
  timeoutMs: number;
  noHooks: boolean;
} {
  const out: string[] = [];
  let timeoutMs = DEFAULT_TIMEOUT_MS;
  let noHooks = false;

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];

    if (arg === "--timeout") {
      const raw = args[i + 1];
      if (!raw || raw.startsWith("--")) {
        throw new Error("Usage: --timeout <seconds>");
      }
      timeoutMs = parseTimeoutValue(raw);
      i++;
      continue;
    }

    if (arg.startsWith("--timeout=")) {
      timeoutMs = parseTimeoutValue(arg.slice("--timeout=".length));
      continue;
    }

    if (arg === "--no-hooks") {
      noHooks = true;
      continue;
    }

    out.push(arg);
  }

  return { args: out, timeoutMs, noHooks };
}

async function runWithTimeout<T>(
  fn: (signal: AbortSignal) => Promise<T>,
  timeoutMs: number,
): Promise<T> {
  const controller = new AbortController();
  registerSignalTimeout(
    controller.signal,
    timeoutMs,
    `Command timed out after ${Math.ceil(timeoutMs / 1000)}s`,
  );
  let timer: ReturnType<typeof setTimeout> | null = null;

  try {
    return await Promise.race([
      fn(controller.signal),
      new Promise<never>((_, reject) => {
        timer = setTimeout(() => {
          const message = `Command timed out after ${Math.ceil(timeoutMs / 1000)}s`;
          controller.abort(new Error(message));
          reject(new Error(message));
        }, timeoutMs);
      }),
    ]);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

async function writeTo(
  stream: NodeJS.WriteStream,
  data: string,
): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    let settled = false;

    const finish = (err?: NodeJS.ErrnoException | null) => {
      if (settled) return;
      settled = true;
      stream.off("error", onError);

      if (
        err?.code === "EPIPE" ||
        err?.message?.toUpperCase().includes("EPIPE")
      ) {
        resolve();
        return;
      }
      if (err) {
        reject(err);
        return;
      }
      resolve();
    };

    const onError = (err: NodeJS.ErrnoException) => finish(err);

    stream.once("error", onError);
    try {
      stream.write(data, (err?: Error | null) => {
        finish((err as NodeJS.ErrnoException | null | undefined) ?? null);
      });
    } catch (err) {
      finish(err as NodeJS.ErrnoException);
    }
  });
}

async function flushStdStreams(): Promise<void> {
  await Promise.all([writeTo(process.stdout, ""), writeTo(process.stderr, "")]);
}

async function exitWith(
  code: number,
  output?: { stdout?: string; stderr?: string },
): Promise<never> {
  if (output?.stdout) {
    await writeTo(process.stdout, output.stdout);
  }
  if (output?.stderr) {
    await writeTo(process.stderr, output.stderr);
  }

  await flushStdStreams();
  process.exit(code);
}

export async function main(): Promise<void> {
  let args: string[];
  let timeoutMs: number;
  let noHooks: boolean;

  try {
    const parsed = parseGlobalOptions(process.argv.slice(2));
    args = parsed.args;
    timeoutMs = parsed.timeoutMs;
    noHooks = parsed.noHooks;
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return exitWith(1, { stderr: `${JSON.stringify({ error: message })}\n` });
  }

  if (args.length === 0 || args[0] === "--help" || args[0] === "-h") {
    return exitWith(0, { stdout: `${USAGE}\n` });
  }

  if (args[0] === "--version" || args[0] === "-v") {
    const require = createRequire(import.meta.url);
    const { version } = require("../../package.json") as { version: string };
    return exitWith(0, { stdout: `${version}\n` });
  }

  const noDaemon = hasFlag(args, "--no-daemon");
  const filteredArgs = args.filter(
    (a) => a !== "--no-daemon" && a !== "--no-hooks",
  );

  if (filteredArgs[0] === "daemon") {
    const sub = filteredArgs[1];
    if (sub === "start") {
      try {
        await spawnDaemon(timeoutMs);
        return exitWith(0, {
          stdout: `${JSON.stringify({ daemon: "started" })}\n`,
        });
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        return exitWith(1, {
          stderr: `${JSON.stringify({ error: message })}\n`,
        });
      }
    }

    if (sub === "stop" || sub === "status" || sub === "flush") {
      const result = await tryExistingDaemon(
        filteredArgs,
        process.cwd(),
        undefined,
        timeoutMs,
      );
      if (result) {
        return exitWith(result.exitCode, {
          stdout: result.stdout,
          stderr: result.stderr,
        });
      }

      if (sub === "status") {
        return exitWith(0, {
          stdout: `${JSON.stringify({ running: false })}\n`,
        });
      }

      return exitWith(1, {
        stdout: `${JSON.stringify({ error: "No daemon running" })}\n`,
      });
    }

    return exitWith(1, {
      stderr: "Usage: hsx daemon start|stop|status|flush\n",
    });
  }

  // Wire hook output through CLI's IO layer, then discover
  setHookWriters(
    (data) => writeStdout(data),
    (data) => writeStderr(data),
  );
  if (!noHooks) {
    await discoverHooks();
  }

  if (!noDaemon) {
    let stdin: string | undefined;
    const command = filteredArgs[0];
    const setJsonArg = filteredArgs[3];
    const evalCodeArg = filteredArgs[2];
    const needsStdin =
      (command === "set" && (!setJsonArg || setJsonArg === "-")) ||
      (command === "eval" && (!evalCodeArg || evalCodeArg === "-"));

    if (needsStdin && !process.stdin.isTTY) {
      const chunks: Buffer[] = [];
      for await (const chunk of process.stdin) {
        chunks.push(chunk as Buffer);
      }
      stdin = Buffer.concat(chunks).toString("utf-8").trim();
    }

    const result = await tryDaemon(
      filteredArgs,
      process.cwd(),
      stdin,
      timeoutMs,
    );
    if (result) {
      return exitWith(result.exitCode, {
        stdout: result.stdout,
        stderr: result.stderr,
      });
    }

    if (stdin !== undefined) {
      const io = createIoContext(stdin);
      try {
        await runWithTimeout(
          (signal) =>
            Promise.resolve(
              runWithIo(io, () => dispatch(filteredArgs, { signal })),
            ),
          timeoutMs,
        );
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        io.stderr += `${JSON.stringify({ error: message })}\n`;
        return exitWith(1, { stdout: io.stdout, stderr: io.stderr });
      }
      return exitWith(0, { stdout: io.stdout, stderr: io.stderr });
    }
  }

  try {
    await runWithTimeout(
      (signal) => dispatch(filteredArgs, { signal }),
      timeoutMs,
    );
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return exitWith(1, { stderr: `${JSON.stringify({ error: message })}\n` });
  }

  return exitWith(0);
}
