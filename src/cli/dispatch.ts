import { clear } from "./commands/clear.js";
import { copy } from "./commands/copy.js";
import { create } from "./commands/create.js";
import { csv } from "./commands/csv.js";
import { evalCode } from "./commands/eval.js";
import { get } from "./commands/get.js";
import { info } from "./commands/info.js";
import { objects } from "./commands/objects.js";
import { resize } from "./commands/resize.js";
import { type RcDim, type RcOp, rowsCols } from "./commands/rows-cols.js";
import { search } from "./commands/search.js";
import { set } from "./commands/set.js";
import { type SheetOp, sheet } from "./commands/sheet.js";
import { writeStderr, writeStdout } from "./output.js";

export const USAGE = `Usage: hsx <command> [args]

Commands:
  create <file>                              Create a new Excel file
  info <file>                                Show workbook metadata
  get <file> <ref>                           Read cells (Sheet1!A1:C10)
  csv <file> <ref>                           Read range as CSV
  set <file> <ref> [json]                    Write cells (JSON from arg or stdin)
  clear <file> <ref> [--type all|styles]     Clear a range (default: values only)
  search <file> <term> [--sheet S] [--regex] Search for values across sheets
  copy <file> <src> <dst>                    Copy range (formulas + styles)
  sheet <file> <op> [args]                   list | create | delete | rename
  rc <file> <op> <dim> [--ref R] [--count N] Insert/delete/hide/freeze rows or columns
  resize <file> [--columns A:D] [--width N]  Resize column widths or row heights
  objects <file> [--sheet <name>]             List charts, tables, pivots
  eval <file> [code]                         Execute arbitrary JS (code from arg or stdin)
  daemon start                               Start the background daemon
  daemon stop                                Stop the background daemon
  daemon status                              Show daemon status
  daemon flush                               Flush buffered workbook writes

Options:
  --no-daemon                                Skip daemon, run directly
  --timeout <seconds>                        Command timeout (default: 30s)

Reference format:
  Sheet1!A1:C10    range on named sheet
  A1:C10           range on active sheet
  A1               single cell

Globals available in eval:
  workbook   SpreadJS Workbook instance
  sheet      Active worksheet
  GC         GC.Spread.Sheets namespace
  file       ExcelFile wrapper (batch, save, toJSON)`;

function flag(args: string[], name: string): string | undefined {
  const idx = args.indexOf(name);
  return idx !== -1 ? args[idx + 1] : undefined;
}

function hasFlag(args: string[], name: string): boolean {
  return args.includes(name);
}

function requireArg(args: string[], index: number, usage: string): string {
  const value = args[index];
  if (!value || value.startsWith("--")) {
    throw new Error(`Usage: hsx ${usage}`);
  }
  return value;
}

function parseOptionalInt(
  value: string | undefined,
  optionName: string,
): number | undefined {
  if (value === undefined) return undefined;
  const parsed = Number.parseInt(value, 10);
  if (Number.isNaN(parsed)) {
    throw new Error(`Invalid value for ${optionName}: ${value}`);
  }
  return parsed;
}

function parseOptionalFloat(
  value: string | undefined,
  optionName: string,
): number | undefined {
  if (value === undefined) return undefined;
  const parsed = Number.parseFloat(value);
  if (Number.isNaN(parsed)) {
    throw new Error(`Invalid value for ${optionName}: ${value}`);
  }
  return parsed;
}

function parseClearType(
  value: string | undefined,
): "values" | "styles" | "all" {
  if (!value) return "values";
  if (value === "values" || value === "styles" || value === "all") {
    return value;
  }
  throw new Error(`Invalid value for --type: ${value}`);
}

export async function dispatch(args: string[]): Promise<void> {
  const command = args[0];
  const rest = args.slice(1);

  switch (command) {
    case "create": {
      const file = requireArg(rest, 0, "create <file>");
      await create(file);
      break;
    }

    case "info": {
      const file = requireArg(rest, 0, "info <file>");
      await info(file);
      break;
    }

    case "get": {
      const file = requireArg(rest, 0, "get <file> <ref>");
      const ref = requireArg(rest, 1, "get <file> <ref>");
      const styles = !hasFlag(rest, "--no-styles");
      await get(file, ref, { styles });
      break;
    }

    case "csv": {
      const file = requireArg(rest, 0, "csv <file> <ref>");
      const ref = requireArg(rest, 1, "csv <file> <ref>");
      await csv(file, ref);
      break;
    }

    case "set": {
      const file = requireArg(rest, 0, "set <file> <ref> [json]");
      const ref = requireArg(rest, 1, "set <file> <ref> [json]");
      const json = rest[2];
      await set(file, ref, json);
      break;
    }

    case "clear": {
      const file = requireArg(rest, 0, "clear <file> <ref>");
      const ref = requireArg(rest, 1, "clear <file> <ref>");
      await clear(file, ref, parseClearType(flag(rest, "--type")));
      break;
    }

    case "search": {
      const file = requireArg(rest, 0, "search <file> <term>");
      const term = requireArg(rest, 1, "search <file> <term>");
      await search(file, term, {
        sheet: flag(rest, "--sheet"),
        matchCase: hasFlag(rest, "--match-case"),
        regex: hasFlag(rest, "--regex"),
        maxResults: parseOptionalInt(flag(rest, "--max"), "--max"),
      });
      break;
    }

    case "copy": {
      const file = requireArg(rest, 0, "copy <file> <src> <dst>");
      const src = requireArg(rest, 1, "copy <file> <src> <dst>");
      const dst = requireArg(rest, 2, "copy <file> <src> <dst>");
      await copy(file, src, dst);
      break;
    }

    case "sheet": {
      const file = requireArg(rest, 0, "sheet <file> <op> [args]");
      const op = requireArg(rest, 1, "sheet <file> <op> [args]");
      const opArgs = rest.slice(2);
      await sheet(file, op as SheetOp, opArgs);
      break;
    }

    case "rc": {
      const file = requireArg(rest, 0, "rc <file> <op> <dim>");
      const op = requireArg(rest, 1, "rc <file> <op> <dim>") as RcOp;
      const dim = requireArg(rest, 2, "rc <file> <op> <dim>") as RcDim;
      await rowsCols(file, op, dim, {
        sheet: flag(rest, "--sheet"),
        ref: flag(rest, "--ref"),
        count: parseOptionalInt(flag(rest, "--count"), "--count"),
      });
      break;
    }

    case "resize": {
      const file = requireArg(rest, 0, "resize <file> [options]");
      await resize(file, flag(rest, "--sheet"), {
        columns: flag(rest, "--columns"),
        rows: flag(rest, "--rows"),
        width: parseOptionalFloat(flag(rest, "--width"), "--width"),
        height: parseOptionalFloat(flag(rest, "--height"), "--height"),
      });
      break;
    }

    case "objects": {
      const file = requireArg(rest, 0, "objects <file>");
      await objects(file, flag(rest, "--sheet"));
      break;
    }

    case "eval": {
      const file = requireArg(rest, 0, "eval <file> [code]");
      const code = rest[1];
      await evalCode(file, code);
      break;
    }

    default:
      writeStderr(`Unknown command: ${command}\n`);
      writeStdout(`${USAGE}\n`);
      throw new Error(`Unknown command: ${command}`);
  }
}
