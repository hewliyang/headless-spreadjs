#!/usr/bin/env node

const MAX_ERROR_MESSAGE_LENGTH = 500;
const UNSUPPORTED_SPREADJS_FEATURE_ERROR =
  "SpreadJS failed (the file may contain unsupported features such as form controls)";

function toErrorMessage(reason: unknown): string {
  const message =
    reason instanceof Error ? reason.message : String(reason ?? "");
  const lower = message.toLowerCase();
  const looksLikeUnsupportedSpreadJsFeature =
    lower.includes("form control") ||
    lower.includes("shape") ||
    lower.includes("chart") ||
    lower.includes("canvas") ||
    lower.includes("getcontext") ||
    lower.includes("todataurl");

  if (
    looksLikeUnsupportedSpreadJsFeature &&
    message.length > MAX_ERROR_MESSAGE_LENGTH
  ) {
    return UNSUPPORTED_SPREADJS_FEATURE_ERROR;
  }

  if (message.length > MAX_ERROR_MESSAGE_LENGTH) {
    return `${message.slice(0, MAX_ERROR_MESSAGE_LENGTH)}...`;
  }

  return message;
}

function reportFatalError(reason: unknown): never {
  const err = reason as NodeJS.ErrnoException;
  if (err?.code === "EPIPE") {
    process.exit(0);
  }

  process.stderr.write(
    `${JSON.stringify({ error: toErrorMessage(reason) })}\n`,
  );
  process.exit(1);
}

function handlePipeError(err: NodeJS.ErrnoException): void {
  if (err.code === "EPIPE") {
    process.exit(0);
  }
}

process.stdout.on("error", handlePipeError);
process.stderr.on("error", handlePipeError);

process.on("uncaughtException", reportFatalError);
process.on("unhandledRejection", reportFatalError);

const { main } = await import("./main.js");
await main().catch(reportFatalError);
