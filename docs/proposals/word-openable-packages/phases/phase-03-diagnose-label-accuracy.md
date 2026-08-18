# Phase 3 — diagnose() label accuracy

**Created:** 2026-08-18
**Status:** Ready

## Motivation

**This phase is diagnostic polish, not a safety fix.** The dangerous half of the `diagnose()`
defect is closed for free by Phase 1: a prefixed archive whose prefix begins with the bytes `PK`
currently returns `readable=True, kind="docx"` — the triage API vouching for a file Word refuses —
and once the loader refuses it, `diagnose` catches `PackageLimitError` and already reports
`kind="unsafe-archive"`. What remains is the other prefix shape, where `diagnose` returns
`kind="not-a-zip"`, which is false: it *is* a ZIP, just not one that starts at the beginning of
the file. `diagnose` exists so an agent can repair a file without a human, so a false label
defeats its only purpose.

## Context

**What exists today:**
- `src/docx/_paperpkg.py` `diagnose` reads 8 bytes, checks for a compound-file marker, then
  returns `kind="not-a-zip"` when the header does not start with `b"PK"` — before ever attempting
  `_read_zip` at `:481`.
- Measured at the stack tip:

  | input | result today |
  |---|---|
  | prefix not starting with `PK` | `readable=False, kind="not-a-zip"` — false label |
  | prefix starting with `PK` | `readable=True, kind="docx"` — fixed by Phase 1 |
  | genuine non-zip | `readable=False, kind="not-a-zip"` — correct |
  | clean document | `readable=True, kind="docx"` — correct |

- `PackageDiagnosis.kind` is a public field with a documented value set: `missing`,
  `encrypted-or-legacy-binary`, `not-a-zip`, `corrupt-zip`, `unsafe-archive`, `xlsx`, `pptx`,
  `docx`, `docm`, `dotm`, `dotx`, `opc-unknown`. `to_dict()` exposes it.

**What this phase delivers:**
- Both prefixed shapes report `readable=False, kind="unsafe-archive"` with a problem string naming
  the defect.
- A genuine non-zip still reports `not-a-zip`.

**Reference files to study before starting:**
- `src/docx/_paperpkg.py` — `diagnose`, its `result()` helper, and the `PackageDiagnosis` docstring
- `src/docx/_zipguard.py` — `_find_end_record` shows the tail-scan technique for locating an
  end-of-central-directory record

## Steps

### Step 1 — Distinguish "no ZIP structure" from "ZIP structure not at byte 0"

**Goal:** the early `not-a-zip` conclusion is only drawn when the file genuinely has no ZIP
structure.

**Work:**
Before concluding `not-a-zip`, determine whether the file contains an end-of-central-directory
record at all. If it does, the file is a ZIP that does not begin at the start of the file: report
`readable=False` with `kind="unsafe-archive"` and a problem string saying the archive does not
begin at the start of the file and must be extracted to its own file. If it does not, keep the
existing `not-a-zip` result unchanged.

Locate the end record by scanning the tail, the same way `_find_end_record` does — the record is
within the last 64 KiB plus the record's own size. Do not parse or validate it; its presence is
the only question being asked.

**Constraints:**
- **Reuse `unsafe-archive`; do not add a `kind` value.** `kind` is a public field with a
  documented value set, and callers may switch on it, so a new value is an additive public-surface
  change. Reusing `unsafe-archive` also makes `diagnose` consistent across both prefix shapes —
  the `PK`-prefixed one already lands there via Phase 1 — and the distinguishing detail belongs in
  `problems`, which is free text for exactly this.
- Do not remove or reorder the compound-file (`_CFB_MAGIC`) branch; an encrypted or legacy binary
  document must keep its own specific label.
- Do not make `diagnose` raise. It is the triage API; every input must produce a
  `PackageDiagnosis`.
- Keep the added work bounded — a tail read, not a whole-file scan.

**Verification:**
Both prefix shapes report `readable=False, kind="unsafe-archive"`; a genuine non-zip reports
`not-a-zip`; a clean document reports `docx`; an OLE compound file still reports
`encrypted-or-legacy-binary`.

### Step 2 — Guard test

**Goal:** all four `diagnose` outcomes above are pinned so the label cannot silently regress.

**Work:**
Add a test alongside the existing `diagnose` coverage asserting the four cases in the table above.
Build the two prefixed fixtures in-test by prepending bytes to a real saved document and rebasing
the central-directory offsets, so they differ only in whether the prefix begins with `PK`.

**Constraints:**
- Assert on `readable` and `kind`, plus one substring of the problem string. Do not assert the
  full message.
- The two prefixed fixtures must differ *only* in the prefix bytes, so the test isolates the
  label decision.

**Verification:**
```
uv run --with pytest --with 'pyparsing<3.3' python -m pytest -q tests/paper/
uv run --with pytest --with 'pyparsing<3.3' python -m pytest -q
```

## Files

| Action | Path |
|--------|------|
| Edit | `src/docx/_paperpkg.py` — `diagnose`: distinguish a prefixed archive from a non-zip before the early return |
| Edit | `tests/paper/` — four-case `diagnose` guard test, in whichever module already covers `diagnose` |

## What this phase does NOT include

- A new `PackageDiagnosis.kind` value.
- Any change to `_read_zip`, `_zipguard.py`, or `opc/package.py`.
- Repairing a prefixed file, or reporting how many bytes precede the archive.
- Any change to the other `diagnose` branches — compound files, `xlsx`/`pptx` detection,
  macro-enabled and template kinds.

## Tests this phase must include

- Prefixed archive, prefix not starting with `PK` → `readable=False, kind="unsafe-archive"`.
- Prefixed archive, prefix starting with `PK` → same result (via Phase 1's loader refusal).
- Genuine non-zip → `not-a-zip`.
- Clean document → `docx`.

**Does NOT need tests:** the tail-scan window size, and the exact problem wording.

## Done when

1. Both prefixed shapes report `readable=False`, `kind="unsafe-archive"`, and a problem string
   naming the defect.
2. A genuine non-zip still reports `not-a-zip`, and a clean document still reports `docx`.
3. No new value is present in the `kind` set.
4. `diagnose` raises on no input.
5. Full suite green per the baseline rule in the build sequence.
