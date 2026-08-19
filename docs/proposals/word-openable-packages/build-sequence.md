# Build Sequence: Refuse packages Word cannot open

**Created:** 2026-08-18
**Status:** Ready

## Dependency Graph

<!-- Machine-readable DAG. /execute-plan parses this block. -->
```yaml
phases:
  - id: 1
    name: Load byte-zero check
    depends_on: []
  - id: 2
    name: Refuse nonzero-offset stream writes
    depends_on: [1]
  - id: 3
    name: diagnose() label accuracy
    depends_on: [1]
  - id: 4
    name: Documentation hygiene
    depends_on: [1, 2, 3]
```

## Context

### What exists today

`_zipguard._preflight_zip_stream` validates that a package's central directory is *internally
consistent* — `central_offset + central_size == central_end` — which a prefixed archive with
rebased offsets satisfies exactly. Nothing checks where the archive starts. Word refuses such a
file; every Python reader opens it and reports the correct content.

`preflight_zip` has exactly two callers, and between them they cover every read entry point:

- `src/docx/opc/phys_pkg.py:127` — `Document()`, `PackageReader`
- `src/docx/_paperpkg.py:226` (`_read_zip`) — `patch_save` (original at `:566`, its own output at
  `:604`), `diff_package` (`:291-292`), `diagnose` (`:481`), `compare` (`_compare.py:314`)

`opc/package._atomic_stream_write` permits a seekable destination at any position. The
write-only branch (`_write_staged_to_unrestorable_stream`) already refuses a nonzero start with
`OSError`; only the seekable branch allows it. A bare `stream.truncate()` at
`opc/package.py:339` then destroys everything after the package.

`_paperpkg.diagnose` returns `readable=True, kind="docx"` for a prefixed archive whose prefix
begins with the bytes `PK`, and `kind="not-a-zip"` when it does not.

### What this feature requires

- A package must begin at byte 0 — a local file header signature, or an end-of-central-directory
  signature for a legal empty archive.
- Saving to a seekable stream positioned past 0 must refuse before staging anything, leaving the
  destination byte-for-byte unchanged.
- `diagnose()` must not label a prefixed archive `not-a-zip`.

### What this feature explicitly excludes

- **Part-name character rules, prefix-collision detection, image relationship content types, the
  protection gate, and message rewording.** Owned by `word-oracle-alignment` and
  `word-oracle-fixes`.
- **Embedding support.** No output of a nonzero-offset write is openable by either route, so
  supporting it needs offset rebasing plus an explicit statement that the container file is not
  itself openable. Separate design, separate Word round.
- **The reachability gap** (`undeclared_orphan_part`, `ds_store`, `macosx_sidecar`). No Word
  verdict yet.
- **Prefix-length reporting in the refusal message.** `_scan_central_directory` returns nothing
  and tracks no offsets; threading a value through it buys wording, not correctness.
- **A `patch_save` verbatim-copy gate.** Phase 1 makes it unreachable — `patch_save` reads the
  original through `_read_zip` at `:566`, long before the verbatim decision at `:577` — and
  testing it would require bypassing the loader. Replaced by a coupling test in Phase 1.

### Cross-cutting rules

- **Exception vocabulary.** Load-path refusals raise `PackageLimitError`; its contract is "a
  package archive is corrupt, encrypted, or malformed", which a prefixed archive is. Destination
  refusals raise `OSError`, matching append-mode and every other destination refusal in
  `opc/package.py`. A stream's cursor position is a property of the argument, not of the archive.
- **Refusal messages** name what was found, why it cannot be interpreted, and what to do.
- **No public signature changes**, and no new `PackageDiagnosis.kind` value — see Phase 3.
- **Smallest diff.** No drive-by refactors. `_zipguard.py` is also being edited by the two
  parallel specs; touch only `_preflight_zip_stream`.

### Reference files

- `src/docx/_zipguard.py` — `_preflight_zip_stream`, `_find_end_record`,
  `_scan_central_directory`, `_LOCAL_HEADER_SIGNATURE`, `_END_RECORD_SIGNATURE`
- `src/docx/opc/package.py` — `_atomic_stream_write`, `_stream_position`,
  `_has_stream_rollback_surface`, `_copy_stream_prefix`
- `src/docx/_paperpkg.py` — `_read_zip`, `diagnose`, `patch_save`
- `src/docx/errors.py` — `PaperRefusal` hierarchy and each subclass's contract
- `tests/paper/test_practical_opc_hardening.py` — save-path guard tests; the two that must be
  rewritten live here
- `tests/paper/test_audit_zip_preflight.py` — preflight guard tests; these are what a
  badly-placed check shadows

### Verification baseline

At the stack tip: **suite green, plus 9 collection errors** in `tests/opc/test_phys_pkg.py` from
a pytest deprecation about class-scoped fixtures declared as instance methods — pre-existing and
unrelated. Do not pin a passing count: this work adds tests and rewrites two. The gate is
"green, the 9 known collection errors, and every test-count delta attributable to a test this
plan intends".

```
uv run --with pytest --with 'pyparsing<3.3' python -m pytest -q
```

### Fixture-verdict replay gate

Phases 1 and 2 both change load or save behaviour for *every* entry point, so per-fix tests
cannot prove nothing else moved. At the end of each, replay the recorded-verdict fixture sets
and assert every fixture lands on its recorded accept/refuse status. Use the replay harness from
the `word-oracle-alignment` branch when available, or regenerate with
`verifying-against-word/scripts/build_review.sh` and compare. Sets `4-followups/` and
`5-encoding/` were built by other sessions and are not produced by `build_review.sh`; take them
from the existing review tree rather than regenerating.

The only status changes this plan intends, across all five sets — verified against the fork at
the stack tip:

| fixture | before | after | phase |
|---|---|---|---|
| `4-followups/06-prefix-data.docx` | accepted | refused | 1 |
| `3-fork-output/07-container-WHOLE-FILE.docx` | accepted | refused | 1 |

Any other movement is a regression.

**Two fixtures are easy to mistake for these and must NOT change status.** `06-` collides across
four sets, so check the set name, not the number:

| fixture | status today | why it must not move |
|---|---|---|
| `1-refused-inputs/06-concatenated.docx` | already refused | not `prefix_data`; refused for an ambiguous central-directory region |
| `3-fork-output/08-container-extracted-package.docx` | already refused | offsets skewed by the prefix length; refused by existing logic, not by this work |

## Phases

### Phase 1: Load byte-zero check

The keystone. Adds the byte-zero requirement to `_preflight_zip_stream`, which closes the
prefixed-archive hole at every read entry point at once, and — because both read back through
`preflight_zip` — also closes `patch_save` prefix propagation and makes a nonzero-offset save
fail with the destination intact.

**Steps:**
- Step 1 — add the check at the end of `_preflight_zip_stream`
- Step 2 — rewrite the two tests that encode the disproved assumption
- Step 3 — add guard tests, including the `patch_save` coupling test

**Touches:** `src/docx/_zipguard.py`, `tests/paper/test_practical_opc_hardening.py`,
`tests/paper/test_audit_zip_preflight.py`
**Done when:** a prefixed archive is refused through `Document()`, `patch_save` and
`diff_package`; the clean control still opens; suite green per the baseline rule; replay gate
shows only the two intended status changes
**Depends on:** nothing (foundation)

### Phase 2: Refuse nonzero-offset stream writes

Turns the diagnosis for a nonzero-offset save from a complaint about a malformed archive —
raised from the staged-output validator after Phase 1 — into an early, accurate statement that
the destination cursor is wrong, and removes the dead prefix-copying helper.

**Steps:**
- Step 1 — refuse a nonzero start in `_atomic_stream_write` before staging
- Step 2 — gate `truncate()` on `start == 0` and delete `_copy_stream_prefix`
- Step 3 — guard test for the refusal and the untouched destination

**Touches:** `src/docx/opc/package.py`, `tests/paper/test_practical_opc_hardening.py`
**Done when:** a nonzero-offset save raises `OSError` naming the position and leaves the
destination byte-identical; offset-0 saves over a longer document still truncate; suite green;
replay gate clean
**Depends on:** Phase 1

### Phase 3: diagnose() label accuracy

**Polish, not safety.** The dangerous half of the `diagnose()` bug — `readable=True, kind="docx"`
for a file Word refuses — is fixed for free by Phase 1, because `diagnose` reads through
`_read_zip` and already maps `PackageLimitError` to `kind="unsafe-archive"`. What remains is the
`not-a-zip` mislabel on a prefixed file whose prefix does not begin with `PK`.

**Steps:**
- Step 1 — distinguish "no ZIP structure at all" from "ZIP structure not at byte 0" in the early
  return, reusing `unsafe-archive` rather than adding a `kind` value
- Step 2 — guard test for both prefix shapes and for a genuine non-zip

**Touches:** `src/docx/_paperpkg.py`, `tests/paper/`
**Done when:** both prefix shapes report `readable=False, kind="unsafe-archive"` with a problem
string naming the defect; a genuine non-zip still reports `not-a-zip`; the clean control still
reports `docx`
**Depends on:** Phase 1

### Phase 4: Documentation hygiene (ALWAYS the final phase)

**Steps:**
- Update `verifying-against-word/WORD-VERDICTS.md`: move `prefix_data` and
  `3-fork-output/07` out of "fork too lenient" into shipped — **`08` was never in that framing;
  it is already refused** — and strike the save-path decisions this plan closes from "Rows that
  still need a human"
- Update `verifying-against-word/scripts/make_save_variants.py` — its README text for files 07
  and 08 states measured verdicts; add that the fork now refuses the write outright
- Update `docs/proposals/word-openable-packages/spec.md` — `status: planned` → `status: shipped`,
  and fold the Phase 3 `kind` decision and the truncate-gate resolution into it
- Note the byte-zero rule in whichever doc surface describes package validation, if one exists
- Grep-sweep for `_copy_stream_prefix` and for prose asserting that a nonzero-offset save
  preserves a prefix

**Verification:**
- `grep -rn "_copy_stream_prefix" .` returns nothing outside git history
- No path referenced in the skill or the proposals tree points at something that does not exist
- `WORD-VERDICTS.md` contains no row claiming the fork accepts a prefixed archive

**Done when:** a reviewer can read the ledger and the spec cold and see that the prefixed-archive
hole is closed on both the read and the write side, without reading the phase plans.
**Depends on:** Phases 1, 2, 3
