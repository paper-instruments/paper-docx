# Phase 4 — Documentation hygiene

**Created:** 2026-08-18
**Status:** Ready

## Motivation

The Word verdict ledger is the durable record of which guards are justified, and it currently
describes the prefixed-archive hole as open with the fork on the wrong side of it. Once Phases
1–3 land that is false, and a ledger that is wrong about a shipped fix is worse than one with a
gap. This phase closes the record while the evidence is still fresh, and removes the generator
text that still tells a human to expect the old behaviour.

## Context

**What exists today:**
- `verifying-against-word/WORD-VERDICTS.md` records `prefix_data` and
  `3-fork-output/07`/`08` as measured Word refusals with paper-docx on the permissive side, and
  lists the save-path decisions under work still outstanding.
- `verifying-against-word/scripts/make_save_variants.py` embeds README text telling the reader
  what to expect from files 07 and 08.
- `docs/proposals/word-openable-packages/spec.md` is `status: planned` and still carries two
  decisions this plan resolved: the `patch_save` verbatim gate and the truncate gate.
- Two parallel specs — `word-oracle-alignment.md`, `word-oracle-fixes.md` — reference the load
  check as this branch's or theirs; whichever landed, the ownership note should match reality.

**What this phase delivers:**
- A ledger that reads correctly cold: the prefixed-archive hole is closed on both the read and
  the write side.
- Generator text that matches shipped behaviour.
- No stale references to removed code.

**Two files in this phase live outside the repo.** `WORD-VERDICTS.md` and
`make_save_variants.py` are under `~/.claude/skills/verifying-against-word/`, so those edits will
not appear in the PR diff and cannot be reviewed with it — flag them to the reviewer separately.
They are also concurrently edited by other sessions: **re-read each file immediately before
writing it** and merge into its current state rather than assuming the state described here.

**Reference files to study before starting:**
- `verifying-against-word/WORD-VERDICTS.md` — the save-path section, the container subsection, and
  "Rows that still need a human"
- `verifying-against-word/scripts/make_save_variants.py` — the README block and the per-file
  expectation lines for 07 and 08
- `docs/proposals/word-openable-packages/spec.md` — frontmatter and open questions

## Steps

### Step 1 — Update the verdict ledger

**Goal:** the ledger states that the byte-zero rule shipped, on both sides, and nothing in it
claims the fork accepts a prefixed archive.

**Work:**
Record in the save-path section and the container subsection that paper-docx now refuses a
prefixed archive on load and refuses a nonzero-offset stream destination on write. Move
`prefix_data` and `3-fork-output/07` out of any "fork too lenient" framing. Note that
`3-fork-output/08` needs no separate fix — it was already refused, and the operation that produced
it is now refused outright.

Strike from "Rows that still need a human" the save-path items this plan closed, and keep the
non-writable-directory trade listed as a decision rather than a verdict.

**Constraints:**
- Do not alter any recorded Word verdict. Those are human observations; only the paper-docx column
  and the narrative change.
- Do not claim a verdict for anything unmeasured — the reachability gap stays open.
- Keep the "no verdict borrowed from another application" rule intact.

**Verification:**
```
grep -n "too lenient" ~/.claude/skills/verifying-against-word/WORD-VERDICTS.md
```
returns no row describing a prefixed archive as accepted.

### Step 2 — Update the generator and the spec

**Goal:** anything a human reads before opening the fixtures matches shipped behaviour, and the
spec records the decisions this plan made.

**Work:**
In `make_save_variants.py`, update the README text for files 07 and 08 to say the fork now
refuses the write outright, so the artifacts are a record of why rather than a live test.

In `spec.md`, set `status: planned` → `status: shipped`, and fold in the two resolutions: the
`patch_save` verbatim gate was dropped in favour of a coupling test because it would have needed a
loader bypass to test, and the truncate gate was kept as an explicit precondition. Remove the
resolved open question about `compare()` — it routes through `_read_zip` at `_compare.py:314`, so
that coverage is sufficient.

Record the corpus-scan path used for false-positive validation — 87 `.docx` under the worktree
root (45 in `tests/`, 41 in `features/`, 1 in `src/`) — so the number is reproducible and the
discrepancy with the parallel specs' 83 is traceable to scanned paths rather than a conflict.

**Constraints:**
- Do not restate phase plans in the spec. It records decisions, not steps.
- Leave the two parallel specs alone unless their ownership note contradicts what shipped.

**Verification:** both files render cleanly; `spec.md` frontmatter reads `status: shipped`.

### Step 3 — Stale-reference sweep

**Goal:** nothing in the repo or the skill points at removed code or disproved behaviour.

**Work:**
Search for `_copy_stream_prefix` and for prose asserting that a nonzero-offset save preserves a
caller's prefix or trailing data. Update or delete each hit.

**Constraints:**
- Sweep only for what this plan changed. `_snapshot_stream_tail`, `_restore_stream_tail`,
  `_has_stream_rollback_surface` and `_stream_position` all stay live — verified — so do not
  remove references to them.

**Verification:**
```
grep -rn "_copy_stream_prefix" . --exclude-dir=.git
grep -rniE "preserv(e|es|ing) (the )?(caller|prefix|trailing)" src/ tests/ docs/
```

## Files

| Action | Path |
|--------|------|
| Edit | `~/.claude/skills/verifying-against-word/WORD-VERDICTS.md` — shipped status for the byte-zero rule; outstanding-rows list |
| Edit | `~/.claude/skills/verifying-against-word/scripts/make_save_variants.py` — README expectations for files 07 and 08 |
| Edit | `docs/proposals/word-openable-packages/spec.md` — `status: shipped`, resolved decisions, corpus path |

## What this phase does NOT include

- New Word verdicts, or any change to a recorded one.
- Edits to `word-oracle-alignment.md` or `word-oracle-fixes.md` beyond an ownership correction.
- Documenting the reachability gap or the rendering axis as resolved — both remain open.
- Any source change. If this phase finds a code defect, it goes back to the owning phase.

## Tests this phase must include

None. This phase changes documentation only. The verification is the grep sweep and a clean read
of the ledger.

## Done when

1. `WORD-VERDICTS.md` records the byte-zero rule as shipped on both the read and the write side,
   with every human-recorded Word verdict unchanged.
2. No row in the ledger describes paper-docx as accepting a prefixed archive.
3. `make_save_variants.py`'s README text for files 07 and 08 matches shipped behaviour.
4. `spec.md` is `status: shipped` and records the verbatim-gate and truncate-gate decisions plus
   the corpus-scan path.
5. `grep -rn "_copy_stream_prefix" . --exclude-dir=.git` returns nothing.
6. A reviewer can read the ledger and spec cold and see the hole is closed on both sides without
   reading the phase plans.
