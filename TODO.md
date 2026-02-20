# TODO

Triage outcome for feature requests (as of 2026-02-19), incorporating follow-up discussion.

## Decisions

- [x] **(2) `hsx get` JSON truncation**
  - Status: **already fixed** (treat as closed unless a new repro appears).

- [x] **(1) A1 helper in `eval`** — **shipped (minimal passthrough form)**
  - Added `range(ref)` inside eval runtime:
    - resolves A1 refs on active sheet by default
    - supports explicit/quoted sheet refs (`Sheet2!A1`, `'My Sheet'!A1:B2`)
    - returns native SpreadJS `Range` object (passthrough, no wrapper shape)
  - Deferred extras (`cell`/`formula`/`setCell`) to keep surface minimal.

- [x] **(7) `hsx csv --formulas`** — **shipped**
  - Added formula-aware export mode for quick audits:
    - `--mode value|formula|both`
  - Added `--formulas` as shorthand for `--mode formula`.

- [x] **(6) `hsx diff`** — **shipped (v1 value + formula)**
  - Compares two workbooks by sheet/cell (`A1`)
  - Reports diffs for `value` + `formula`
  - Emits a short summary string (`summary`) at top-level output
  - If diff volume exceeds inline threshold, writes full NDJSON diff to temp file (`diffFile`) for grep-friendly inspection
  - Styles/formatting/object diffs deferred

- [ ] **(5) Dependency tracing** — **Scoped implementation only**
  - `deps <cell>` (what this cell references):
    - Use SpreadJS `getPrecedents(row, col)` directly.
  - `refs <cell>` (what references this cell):
    - **Do NOT rely on SpreadJS `getDependents` naively** (can be prohibitively expensive).
    - Implement ourselves:
      1. Iterate only used/formula cells per sheet.
      2. For each formula cell, call `getPrecedents`.
      3. Match target against precedent ranges.
  - Label results as **one-hop / best-effort** (dynamic formulas like `INDIRECT` can limit precision).

- [x] **(4) `hsx formulas` command** — **Not necessary now**
  - Decision: deprioritize/skip for now since `hsx get` already returns formulas.
  - Reconsider only if repeated user demand appears.

- [x] **(3) eval reserved names (`sheet`)** — **No action needed now**
  - Decision: acceptable UX; users learn after first failure.
  - Optional docs polish later, but not prioritized.

## Suggested execution order

1. (6) Diff v1 (value + formula)
2. (5) Deps/refs (one-hop, bounded)

## Notes / guardrails

- Keep new commands audit-friendly and deterministic.
- Prefer bounded scans over full-sheet brute force.
- Make limitations explicit in help text for dependency tracing.
- Add regression tests for each command before release.
