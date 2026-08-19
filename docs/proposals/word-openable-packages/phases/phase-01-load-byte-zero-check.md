# Phase 1 — Load byte-zero check

**Created:** 2026-08-18
**Status:** Ready

## Motivation

Word refuses a package that does not begin at byte 0; paper-docx opens it and reports the
correct content, so nothing downstream can tell the document is unusable. This is the only
shape in the ledger where the fork is unsafely permissive. Adding one check to the shared
preflight closes it at every read entry point, and because `patch_save` and the save-path
validator both read back through the same preflight, it also closes prefix propagation and makes
a nonzero-offset save fail with the destination intact — which is what Phases 2 and 3 build on.

## Context

**What exists today:**
- `src/docx/_zipguard.py` `_preflight_zip_stream` validates the central directory is *internally
  consistent* (`central_offset + central_size == central_end`). A prefixed archive with rebased
  offsets satisfies that exactly, so it passes.
- `_LOCAL_HEADER_SIGNATURE` (`PK\x03\x04`) and `_END_RECORD_SIGNATURE` (`PK\x05\x06`) are already
  module constants.
- `preflight_zip` has two callers, covering every read path: `opc/phys_pkg.py:127` and
  `_paperpkg._read_zip` (`:226`), the latter reached by `patch_save` (original `:566`, own output
  `:604`), `diff_package` (`:291-292`), `diagnose` (`:481`) and `compare` (`_compare.py:314`).
- Two tests assert the behaviour Word disproved — see Step 2.

**What this phase delivers:**
- A prefixed archive is refused with `PackageLimitError` through every read entry point.
- The two disproved tests express the measured behaviour instead.
- Guard tests pinning the refusal, the adjacent legal shape, and the `patch_save` coupling.

**Reference files to study before starting:**
- `src/docx/_zipguard.py` — the function being changed and the constants to reuse
- `tests/paper/test_audit_zip_preflight.py` — the preflight guard tests a badly-placed check
  shadows; read `_zip64_count_only_archive` and the two tests asserting "too small for its member
  count"
- `tests/paper/test_practical_opc_hardening.py` — the two tests to rewrite
- `src/docx/errors.py` — `PackageLimitError`'s contract

## Steps

### Step 1 — Add the byte-zero requirement to the preflight

**Goal:** `_preflight_zip_stream` refuses any archive whose first four bytes are neither a local
file header signature nor an end-of-central-directory signature.

**Work:**
After the central directory has been fully validated, seek to position 0 and read four bytes.
Accept a local file header signature. Also accept an end-of-central-directory signature, so a
legal empty archive gets its own specific downstream error rather than a byte-zero message it
cannot act on. Anything else is a refusal.

The message must name what was found, why it cannot be interpreted, and what to do: that the
file does not begin with a ZIP local file header, that Word cannot open a `.docx` with bytes
ahead of the archive, and that the archive must be extracted to its own file. Do not report the
number of leading bytes — `_scan_central_directory` returns nothing and tracks no offsets, and
threading a value through it buys wording, not correctness.

**Constraints:**
- **Placement is load-bearing.** The check must run at the *end* of `_preflight_zip_stream`,
  after `_scan_central_directory`, inside the existing `try`. Placed immediately after the
  multi-disk check it shadows more specific diagnoses: measured, it turned two preflight
  refusals that correctly say "central directory is too small for its member count" into the
  generic byte-zero message, failing
  `tests/paper/test_audit_zip_preflight.py::DescribeCentralDirectoryPreflight::it_refuses_a_zip64_count_that_cannot_fit_in_the_central_directory`
  and `::it_preflights_an_ordinary_document_open_before_constructing_zipfile`.
- Reuse `_LOCAL_HEADER_SIGNATURE` and `_END_RECORD_SIGNATURE`. Do not introduce new constants.
- Raise `PackageLimitError`. Its contract covers a malformed archive; a prefixed archive is one.
- The existing `finally` restores the stream position — do not add a second restore.
- Touch only `_preflight_zip_stream`. `_zipguard.py` is being edited concurrently by two other
  specs; keep the diff to one function so the merge stays trivial.
- Do not alter `_find_end_record`, `_scan_central_directory`, or any existing refusal.

**Verification:**
```
uv run --with pytest --with 'pyparsing<3.3' python -m pytest -q tests/paper/test_audit_zip_preflight.py
```
must pass in full — that is the shadowing check. Then confirm the intended refusal fires and the
control does not, through all three entry points (`Document()`, `patch_save`, `diff_package`).

### Step 2 — Rewrite the two tests that encode the disproved assumption

**Goal:** neither test asserts behaviour Word has refuted, and both still test what they were for.

**Work:**
`tests/paper/test_practical_opc_hardening.py::it_validates_the_real_prefix_of_a_nonzero_position_stream`
asserts that a nonzero-offset save succeeds, that the prefix survives, and that the whole stream
reopens as a `Document`. All three are now wrong. Invert it: the save refuses, and the
destination is byte-for-byte unchanged. Keep its name meaningful — it is now about refusing a
nonzero-position stream, so rename accordingly.

`::it_restores_a_seekable_stream_after_commit_error` seeks to offset 7 only incidentally; its
subject is commit-failure rollback. Move it to offset 0 so it still exercises the snapshot and
restore path, and adjust the trailing position assertion.

**Constraints:**
- Do not delete either test. Both cover real behaviour; only their assumptions changed.
- The rollback test must still fail if rollback regresses — verify by confirming it exercises
  `_snapshot_stream_tail` and `_restore_stream_tail`, which stay live at offset 0.
- Do not weaken either test into a smoke test.

**Verification:**
```
uv run --with pytest --with 'pyparsing<3.3' python -m pytest -q tests/paper/test_practical_opc_hardening.py
```

### Step 3 — Guard tests for the new refusal

**Goal:** the refusal, its adjacent legal shape, and the `patch_save` coupling are pinned.

**Work:**
Add tests in `tests/paper/`, following the existing `it_...` function-level convention in
`test_audit_zip_preflight.py`:

- A self-consistent prefixed archive — rebased offsets, clean CRCs, internally correct central
  directory — is refused. Build it by prepending bytes to a real saved document and rebasing
  every central-directory local-header offset plus the end record's central-directory offset, so
  the fixture is refused for the prefix and not for inconsistency.
- The same document without a prefix still opens. This is the pair that makes the verdict
  attributable.
- A ZIP archive comment (bytes *after* the archive, declared correctly) still opens — it proves
  the new check is about the start of the file, not the end, and does not overlap the footer rule.
- `patch_save` refuses a prefixed original. This replaces the verbatim-copy gate that Phase 1
  makes unreachable: it asserts the coupling holds rather than guarding a dead branch.

**Constraints:**
- Name each test so the Word verdict it encodes is traceable.
- The prefixed fixture must be built in-test, not committed as a binary.
- Do not assert on message text beyond one substring that would change if the refusal were
  re-aimed at a different property.

**Verification:**
```
uv run --with pytest --with 'pyparsing<3.3' python -m pytest -q tests/paper/
uv run --with pytest --with 'pyparsing<3.3' python -m pytest -q
```
Then run the fixture-verdict replay gate from the build sequence.

## Files

| Action | Path |
|--------|------|
| Edit | `src/docx/_zipguard.py` — byte-zero check at the end of `_preflight_zip_stream` |
| Edit | `tests/paper/test_practical_opc_hardening.py` — invert the prefix test, move the rollback test to offset 0 |
| Edit | `tests/paper/test_audit_zip_preflight.py` — add the refusal/control/comment guard tests |
| Edit | `tests/paper/` — `patch_save` prefixed-original coupling test, in whichever module already covers `patch_save` |

## What this phase does NOT include

- Any change to `opc/package.py` — that is Phase 2.
- Any change to `diagnose()` — that is Phase 3.
- Prefix-length reporting in the message.
- A `patch_save` verbatim-copy gate. Unreachable after this phase, and testing it would need a
  loader bypass; the coupling test in Step 3 covers the real behaviour instead.
- Part-name rules, prefix-collision detection, content-type or protection changes — other specs.

## Tests this phase must include

- Prefixed archive refused; identical document without the prefix accepted (the attributable pair).
- Declared ZIP archive comment still accepted — separates this rule from the footer rule.
- `patch_save` refuses a prefixed original.
- The full `test_audit_zip_preflight.py` module still passes, which is the placement/shadowing
  guard.

**Does NOT need tests:** the constants reused, and the exact refusal wording beyond one
substring.

## Done when

1. A self-consistent prefixed archive raises `PackageLimitError` through `Document()`,
   `patch_save`, and `diff_package`.
2. The same document without a prefix opens through all three.
3. `tests/paper/test_audit_zip_preflight.py` passes in full — no shadowed diagnoses.
4. The two rewritten tests pass and still cover refusal and rollback respectively.
5. Full suite green with the 9 known `tests/opc/test_phys_pkg.py` collection errors and every
   count delta attributable to a test this phase adds or rewrites.
6. Replay gate shows exactly two status changes: `4-followups/06-prefix-data.docx` and
   `3-fork-output/07-container-WHOLE-FILE.docx`, both accepted → refused. Note the set names:
   `1-refused-inputs/06-` is `06-concatenated.docx`, a different shape that is already refused,
   and `3-fork-output/08` is already refused too — neither may move.
