import { createRequire } from "node:module";
import { spawnDaemon, tryDaemon } from "./client.js";
import { dispatch, USAGE } from "./dispatch.js";
import { createIoContext, runWithIo } from "./output.js";

function hasFlag(args: string[], name: string): boolean {
  return args.includes(name);
}

export async function main(): Promise<void> {
  const args = process.argv.slice(2);

  if (args.length === 0 || args[0] === "--help" || args[0] === "-h") {
    console.log(USAGE);
    process.exit(0);
  }

  if (args[0] === "--version" || args[0] === "-v") {
    const require = createRequire(import.meta.url);
    const { version } = require("../../package.json") as { version: string };
    console.log(version);
    process.exit(0);
  }

  const noDaemon = hasFlag(args, "--no-daemon");
  const filteredArgs = args.filter((a) => a !== "--no-daemon");

  if (filteredArgs[0] === "daemon") {
    const sub = filteredArgs[1];
    if (sub === "start") {
      try {
        await spawnDaemon();
        console.log(JSON.stringify({ daemon: "started" }));
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        console.error(JSON.stringify({ error: message }));
        process.exit(1);
      }
      return;
    }

    if (sub === "stop" || sub === "status") {
      const result = await tryDaemon(filteredArgs, process.cwd());
      if (result) {
        if (result.stdout) process.stdout.write(result.stdout);
        if (result.stderr) process.stderr.write(result.stderr);
        process.exit(result.exitCode);
      }

      if (sub === "stop") {
        console.log(JSON.stringify({ error: "No daemon running" }));
      } else {
        console.log(JSON.stringify({ running: false }));
      }
      process.exit(sub === "stop" ? 1 : 0);
    }

    console.error("Usage: hsx daemon start|stop|status");
    process.exit(1);
  }

  if (!noDaemon) {
    let stdin: string | undefined;
    const command = filteredArgs[0];
    const needsStdin =
      (command === "set" && !filteredArgs[3]) ||
      (command === "eval" && !filteredArgs[2]);

    if (needsStdin && !process.stdin.isTTY) {
      const chunks: Buffer[] = [];
      for await (const chunk of process.stdin) {
        chunks.push(chunk as Buffer);
      }
      stdin = Buffer.concat(chunks).toString("utf-8").trim();
    }

    const result = await tryDaemon(filteredArgs, process.cwd(), stdin);
    if (result) {
      if (result.stdout) process.stdout.write(result.stdout);
      if (result.stderr) process.stderr.write(result.stderr);
      process.exit(result.exitCode);
    }

    if (stdin) {
      const io = createIoContext(stdin);
      try {
        await runWithIo(io, () => dispatch(filteredArgs));
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        io.stderr += `${JSON.stringify({ error: message })}\n`;
        if (io.stdout) process.stdout.write(io.stdout);
        if (io.stderr) process.stderr.write(io.stderr);
        process.exit(1);
      }
      if (io.stdout) process.stdout.write(io.stdout);
      if (io.stderr) process.stderr.write(io.stderr);
      process.exit(0);
    }
  }

  try {
    await dispatch(filteredArgs);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    process.stderr.write(`${JSON.stringify({ error: message })}\n`);
    process.exit(1);
  }
  process.exit(0);
}
