import { createRequire } from "node:module";
import { spawnDaemon, tryDaemon, tryExistingDaemon } from "./client.js";
import { dispatch, USAGE } from "./dispatch.js";
import { createIoContext, runWithIo } from "./output.js";

function hasFlag(args: string[], name: string): boolean {
  return args.includes(name);
}

async function writeTo(
  stream: NodeJS.WriteStream,
  data: string,
): Promise<void> {
  if (!data) return;

  await new Promise<void>((resolve, reject) => {
    const onError = (err: NodeJS.ErrnoException) => {
      stream.off("error", onError);
      if (err.code === "EPIPE") {
        resolve();
        return;
      }
      reject(err);
    };

    stream.once("error", onError);
    stream.write(data, () => {
      stream.off("error", onError);
      resolve();
    });
  });
}

async function flushStdStreams(): Promise<void> {
  await Promise.all([
    new Promise<void>((resolve) => process.stdout.write("", () => resolve())),
    new Promise<void>((resolve) => process.stderr.write("", () => resolve())),
  ]);
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
  const args = process.argv.slice(2);

  if (args.length === 0 || args[0] === "--help" || args[0] === "-h") {
    return exitWith(0, { stdout: `${USAGE}\n` });
  }

  if (args[0] === "--version" || args[0] === "-v") {
    const require = createRequire(import.meta.url);
    const { version } = require("../../package.json") as { version: string };
    return exitWith(0, { stdout: `${version}\n` });
  }

  const noDaemon = hasFlag(args, "--no-daemon");
  const filteredArgs = args.filter((a) => a !== "--no-daemon");

  if (filteredArgs[0] === "daemon") {
    const sub = filteredArgs[1];
    if (sub === "start") {
      try {
        await spawnDaemon();
        return exitWith(0, { stdout: `${JSON.stringify({ daemon: "started" })}\n` });
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        return exitWith(1, { stderr: `${JSON.stringify({ error: message })}\n` });
      }
    }

    if (sub === "stop" || sub === "status" || sub === "flush") {
      const result = await tryExistingDaemon(filteredArgs, process.cwd());
      if (result) {
        return exitWith(result.exitCode, {
          stdout: result.stdout,
          stderr: result.stderr,
        });
      }

      if (sub === "status") {
        return exitWith(0, { stdout: `${JSON.stringify({ running: false })}\n` });
      }

      return exitWith(1, {
        stdout: `${JSON.stringify({ error: "No daemon running" })}\n`,
      });
    }

    return exitWith(1, { stderr: "Usage: hsx daemon start|stop|status|flush\n" });
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

    const result = await tryDaemon(filteredArgs, process.cwd(), stdin);
    if (result) {
      return exitWith(result.exitCode, {
        stdout: result.stdout,
        stderr: result.stderr,
      });
    }

    if (stdin !== undefined) {
      const io = createIoContext(stdin);
      try {
        await runWithIo(io, () => dispatch(filteredArgs));
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        io.stderr += `${JSON.stringify({ error: message })}\n`;
        return exitWith(1, { stdout: io.stdout, stderr: io.stderr });
      }
      return exitWith(0, { stdout: io.stdout, stderr: io.stderr });
    }
  }

  try {
    await dispatch(filteredArgs);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return exitWith(1, { stderr: `${JSON.stringify({ error: message })}\n` });
  }

  return exitWith(0);
}
