# Phase 2 — Refuse nonzero-offset stream writes

**Created:** 2026-08-18
**Status:** Ready

## Motivation

After Phase 1 a nonzero-offset save already fails, but with the wrong diagnosis: the caller gets
`PackageLimitError: package does not begin with a ZIP local file header`, raised from the
staged-output validator, describing a malformed *archive* when the actual fault is a bad
*destination argument*. An agent cannot act on that. This phase moves the refusal to the front of
the write path, where it can name the real cause, and removes the prefix-copying machinery that
only existed to support the capability being refused.

## Context

**What exists today:**
- `opc/package._atomic_stream_write` computes `start = _stream_position(stream)` and proceeds for
  any value. The write-only branch, `_write_staged_to_unrestorable_stream`, already refuses
  `start not in (None, 0)` with `OSError` — only the seekable branch permits it.
- `_copy_stream_prefix(stream, staged, start)` copies the caller's existing prefix into the
  staged buffer. One call site, at `opc/package.py:332`.
- A bare `stream.truncate()` at `opc/package.py:339` truncates after the package regardless of
  where the write began. Measured: a 312,000-byte tail is reduced to 0, where upstream
  python-docx preserves 274,992.
- Word refuses both artifacts of a nonzero-offset write: the whole container file (a prefixed
  archive) and the package sliced back out (offsets skewed by exactly the prefix length).

**What this phase delivers:**
- A seekable destination positioned past 0 is refused before anything is staged, with a message
  naming the position.
- `truncate()` is explicitly gated on `start == 0`.
- `_copy_stream_prefix` is gone.

**Reference files to study before starting:**
- `src/docx/opc/package.py` — `_atomic_stream_write`, `_stream_position`,
  `_has_stream_rollback_surface`, `_write_staged_to_unrestorable_stream`, `_copy_stream_prefix`
- `tests/paper/test_practical_opc_hardening.py` — `it_refuses_an_append_mode_stream_without_changing_it`
  is the pattern to follow: refuse, and assert the destination is untouched

## Steps

### Step 1 — Refuse a nonzero start before staging

**Goal:** `_atomic_stream_write` raises before it stages, snapshots, or writes anything when the
destination is seekable and positioned past 0.

**Work:**
Immediately after `start` is computed, and before `_snapshot_stream_tail` is called, refuse when
`start` is neither `None` nor `0`. This is the same category of fault as the append-mode refusal a
few lines above, so it reads as a pair with it. Raise `OSError`. The
message names the stream's position, states that a package written after existing bytes is not a
file Word can open, and tells the caller to pass a stream positioned at 0 or to write to a path.

`start is None` must stay permitted: that is the bare `write()`/`flush()` sink, which Word
confirmed OPENS.

**Constraints:**
- `OSError`, not `PackageLimitError`. A destination's cursor position is a property of the
  argument, not of the archive, and every other destination refusal in this module raises
  `OSError`. Do not change the append-mode refusal's type either — callers catch `OSError` there.
- Refuse before `_snapshot_stream_tail` is called, so nothing is read or copied.
- Do not touch `_write_staged_to_unrestorable_stream`'s existing refusal. It now duplicates this
  one for write-only streams; leaving it is harmless and keeps that branch independently correct.

**Verification:**
Saving into a stream seeked past 0 raises `OSError`, the message names the position, and the
destination's bytes are unchanged. Saving into a stream at 0, and into a write-only sink, both
still succeed.

### Step 2 — Gate the truncation and delete the dead helper

**Goal:** `truncate()` states its precondition, and no prefix-copying code remains.

**Work:**
Gate the `stream.truncate()` call on `start == 0`.

Then remove `_copy_stream_prefix` and its call. With nonzero offsets refused, `start` in the
snapshot branch is always 0, so the call copies nothing.

**Constraints:**
- The truncate gate is deliberately kept even though Step 1 makes it unreachable —
  `_has_stream_rollback_surface` requires `start is not None`, so a `None` position routes to the
  write-only branch and the snapshot branch is only reached with `start == 0`. It is one line and
  `truncate()`'s no-argument semantics (truncate at the current position) are subtle enough that
  stating the precondition documents intent. This is a different case from the `patch_save`
  verbatim gate dropped in Phase 1: that one would have needed a loader bypass to test at all,
  whereas this line sits inside a path every offset-0 save exercises.
- Delete only `_copy_stream_prefix`. `_snapshot_stream_tail`, `_restore_stream_tail`,
  `_has_stream_rollback_surface`, `_stream_position` and `_copy_stream` all keep live call sites —
  verified — and rollback at offset 0 depends on them.
- Do not restructure `_atomic_stream_write` beyond these two edits.

**Verification:**
```
grep -rn "_copy_stream_prefix" src/ tests/
```
returns nothing. Offset-0 saves over a longer prior document still come out byte-identical to a
fresh save with no stale tail.

### Step 3 — Guard test

**Goal:** the refusal and the untouched destination are pinned, and the offset-0 truncation
behaviour it replaces is pinned alongside it.

**Work:**
Add a test in `tests/paper/test_practical_opc_hardening.py` following
`it_refuses_an_append_mode_stream_without_changing_it`: seek a stream holding existing content
past 0, assert `OSError`, assert the destination is byte-for-byte unchanged, and assert the
position is untouched.

Pair it with an offset-0 save over a longer prior document asserting the result is byte-identical
to a fresh save — that is the behaviour truncation exists for, and it must not regress when the
gate is added.

**Constraints:**
- Assert on the destination's full bytes, not just its length.
- Do not duplicate the append-mode test; this is a distinct fault with a distinct message.

**Verification:**
```
uv run --with pytest --with 'pyparsing<3.3' python -m pytest -q tests/paper/test_practical_opc_hardening.py
uv run --with pytest --with 'pyparsing<3.3' python -m pytest -q
```
Then run the fixture-verdict replay gate from the build sequence.

## Files

| Action | Path |
|--------|------|
| Edit | `src/docx/opc/package.py` — nonzero-start refusal, `truncate()` gated on `start == 0`, `_copy_stream_prefix` and its call removed |
| Edit | `tests/paper/test_practical_opc_hardening.py` — refusal guard test plus the offset-0 truncation pair |

## What this phase does NOT include

- Offset rebasing to make embedding work. No output of that operation is Word-openable by either
  route; supporting it is a separate design with its own Word round.
- Any change to the append-mode refusal or its exception type.
- Any change to `_zipguard.py` or `diagnose()`.
- The non-writable-directory trade — inherent to atomic replacement, a documentation decision.

## Tests this phase must include

- Nonzero-offset save refuses with `OSError` and leaves the destination byte-for-byte unchanged.
- Offset-0 save over a longer prior document is byte-identical to a fresh save — the truncation
  behaviour the gate must not break.
- Write-only sink still succeeds, so `start is None` stays permitted.

**Does NOT need tests:** the removal of `_copy_stream_prefix`, and the exact wording beyond one
substring naming the position.

## Done when

1. A seekable destination positioned past 0 raises `OSError` naming the position, before any
   staging, and the destination is unchanged.
2. Offset-0 saves, including over a longer prior document, still succeed and still truncate.
3. A write-only `write()`/`flush()` sink still succeeds.
4. `grep -rn "_copy_stream_prefix" src/ tests/` returns nothing.
5. Full suite green per the baseline rule in the build sequence.
6. Replay gate shows no status change beyond Phase 1's two.
