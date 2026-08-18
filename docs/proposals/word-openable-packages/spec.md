---
title: Refuse packages Word cannot open
author: gavin
created: 2026-08-18
status: shipped
---

# Refuse packages Word cannot open

## Motivation

paper-docx exists because python-docx let agents produce `.docx` files that Word will not
open. A Word for Mac round on 2026-08-18 found the fork doing exactly that, in the one place
it is least excusable — its own save path.

Word refuses a package that does not begin at byte 0. Prefixed archives are valid ZIP; they
are not valid `.docx`. paper-docx currently **reads** them (reporting success), **writes**
them when a caller saves into a stream positioned past 0, and **byte-copies** them on a no-op
`patch_save`. Every Python reader — stdlib, upstream, and the fork — opens them, so nothing
below Word flags the problem.

All four behaviours were re-verified unfixed at `agent/missing-source-letterhead` (PR 22, the
stack tip, which contains PRs 13 and 15–25). Word has refused the shape twice, on two
independently built fixtures: `3-fork-output/07-container-WHOLE-FILE` (the fork's own output)
and `1-refused-inputs/06-prefix-data`.

## Relationship to the two parallel specs

Three specs were drafted independently from the same Word rounds: `word-oracle-alignment.md`,
`word-oracle-fixes.md`, and this one. They agree on every shared item. Ownership, as settled
across all three:

| item | owner |
|---|---|
| load-time byte-zero check | `word-oracle-alignment` (Fix 2) / `word-oracle-fixes` (L1) |
| part-name character rule, prefix-collision, image content type, protection gate, messages | those two specs |
| truncation gate (`package.py:339`) | folded into the write refusal; retired rather than fixed |
| **write-side refusal of nonzero-offset streams** | **this spec** |
| **`patch_save` verbatim-copy gate** | **this spec** |
| **`diagnose()` reporting a prefixed file as `not-a-zip`** | **this spec** |

Both parallel specs assign the last three here explicitly, so this spec narrows to them. The
load check is cited below because it is what makes the two gates defence-in-depth rather than
the fix — not because this branch implements it.

**Their #5 is the keystone, and it closes more than expected.** Applied alone as an experiment,
it: refuses prefixed reads; makes `patch_save`'s verbatim path unreachable for prefixed input
(because `patch_save` reads through `preflight_zip`); and makes a nonzero-offset save fail
*with the destination left byte-for-byte unchanged*. So the data-loss and bad-output outcomes of
issues 1, 2, and 3 are all closed by the load check.

What that leaves for this spec is **diagnosis quality and defence in depth**, not correctness:
after their #5, a caller who saves into a stream at offset 33 gets
`PackageLimitError: package does not begin with a ZIP local file header` — a complaint about an
archive, raised from the staged-output validator, for what is actually a bad destination
argument. An agent cannot act on that.

Two coordination facts:

- **Their `_zipguard.py` line numbers (280, 510, 947) are from the pre-PR-24 revision** (975
  lines; this branch has 793). PR 24 removed 182 lines from that file, so their patch will
  conflict. Their `package.py:339` and `pkgreader.py:417` refs are exact.
- **Both threads break the same two tests.** Their #5 alone fails exactly
  `it_validates_the_real_prefix_of_a_nonzero_position_stream` and
  `it_restores_a_seekable_stream_after_commit_error`, and nothing else. Whichever lands first
  owns those edits.

## Goals

- No sequence of paper-docx calls produces a `.docx` that Word refuses on envelope grounds.
- Opening a package Word cannot open raises a typed refusal instead of reporting success.
- `diagnose()` names the defect accurately enough for an agent to fix the file unaided.
- The seven save destinations Word confirmed **OPENS** keep working, byte-for-byte.

## Approach

One rule — *an OPC package begins at byte 0* — enforced at both boundaries, plus two
defence-in-depth gates behind it.

**Load (`_zipguard._preflight_zip_stream`).** Require the first four bytes of the archive to be
a local file header. This is the only new load-time check needed: `preflight_zip` is called from
exactly two places — `opc/phys_pkg.py` (`Document()`, `PackageReader`) and
`_paperpkg._read_zip` (`patch_save`, `diff_package`, `diagnose`, `compare`) — so one check
covers every read entry point.

**Placement is load-bearing.** The check must run *after* `_scan_central_directory`, at the end
of `_preflight_zip_stream` — not immediately after the multi-disk check. Placed early it
shadows more specific diagnoses: measured, it turned two preflight refusals that assert
"central directory is too small for its member count" into the generic byte-0 message. Placed
at the end, the full suite shows exactly the two expected failures and nothing else.

**Write (`opc/package._atomic_stream_write`).** Refuse a seekable destination whose position is
not 0, before anything is staged. The write-only branch already refuses this; only the seekable
branch permits it. After the load check this no longer changes *whether* the save fails — it
changes the caller's diagnosis from a complaint about a malformed archive to a statement that
the destination cursor is wrong, and it fails before doing the staging work.

**Two gates behind the rule.** `opc/package.py:339` truncates unconditionally, destroying a
caller's trailing data — gate it on `start == 0`. `_paperpkg.patch_save` takes its
verbatim-copy path on any no-op — gate it on the original beginning at byte 0. Fixing the load
check already makes both unreachable for prefixed input; they stay as the last line of defence
for any future envelope anomaly the loader tolerates.

**Exception types.** Load refusals raise `PackageLimitError` — its contract is "a package
archive is corrupt, encrypted, or malformed", which a prefixed archive is, and `diagnose()`
already maps it to `kind="unsafe-archive"`. Write refusals raise `OSError`, matching
append-mode and every other destination refusal in `opc/package.py`. A destination's cursor
position is not a property of the archive.

**Rejected alternatives:**
- *Rebase offsets so the embedded package is extractable* — fixes the slice, leaves the
  container file itself rejected by Word. Not a fix, and it keeps a capability that has no
  Word-valid output.
- *Also require `min(header_offset) == 0`* — redundant, and it false-positives on
  `concatenated`, which existing logic already refuses with a better-aimed message.
- *A new exception type* — `PackageLimitError` already means this.
- *Unify append-mode onto `PackageLimitError`* — would mis-type a destination-protocol failure
  and break callers catching `OSError`.
- *Repair prefixed files on load* — silently editing something an agent thinks is intact is
  the failure mode this package exists to prevent.

## Scope

### In scope
- Write-time refusal of nonzero-offset seekable stream destinations.
- Truncation gated on `start == 0`, as defence in depth.
- `patch_save` verbatim-copy gated on a clean original envelope.
- `diagnose()` distinguishing "no ZIP structure at all" from "ZIP structure not at byte 0";
  today a prefixed file whose prefix does not start with `PK` is reported as `not-a-zip`,
  which is false and unactionable.
- Deleting `_copy_stream_prefix`, dead once nonzero-offset writes are refused.

### Out of scope
- **The load-time byte-zero check.** Owned by `word-oracle-alignment` Fix 2. This spec depends
  on it and does not duplicate it. If it is placed at the end of `_preflight_zip_stream` as that
  spec requires, it should also accept an end-of-central-directory signature at offset 0 for an
  empty archive, and report the prefix length in its message.
- **The read-side leniency and over-strictness fixes** — part-name characters, prefix collisions,
  image relationship content types, the protection gate, message wording. All owned by the two
  parallel specs.
- **Embedding support.** Writing a package into a container has no Word-openable output by
  either route; supporting it needs a separate design and its own Word round.
- **The remaining ledger rows** — `undeclared_orphan_part`, `ds_store`, `macosx_sidecar`.
  Same class of question (the fork is lenient where Word may not be), no Word verdict yet.
- **The resource caps.** PR 24 removes them; merge it on its own merits.
- **The non-writable-directory trade.** Inherent to atomic replacement; a documentation
  decision, not a correctness bug.

## Requirements

### Functional requirements
- Opening a package whose first four bytes are not a ZIP local file header raises
  `PackageLimitError`, through every entry point that reaches `preflight_zip`.
- Saving to a seekable stream whose position is not 0 raises `OSError` and writes nothing.
  `tell()` returning `None`/unsupported stays permitted — that is the bare write/flush sink
  Word confirmed OPENS.
- Saving to a stream at position 0 still truncates, so no stale tail survives.
- Saving to a stream at a nonzero position leaves the destination byte-for-byte unchanged.
- `patch_save` takes the verbatim path only when the original begins at byte 0.
- `diagnose()` returns `readable=False` with a distinct `kind` for a prefixed archive, and
  does not raise.

### Interface contracts
Refusal messages must say what was found, why it cannot be interpreted, and what to do:

- Load: names that the file does not begin with a ZIP local file header, that Word cannot open
  a `.docx` with bytes ahead of the archive, and that the archive must be extracted to its own
  file.
- Write: names the stream's position, that a package written after existing bytes is not a file
  Word can open, and that the caller should pass a stream at position 0 or write to a path.

No public signature changes. `PatchSaveResult.verbatim_copy` keeps its meaning; it simply
reports `False` more often.

## Test strategy

### What earns its keep
- **Load refusal** on a self-consistent prefixed archive — internally correct offsets, clean
  CRCs — through `Document()`, `patch_save`, and `diff_package`. This is the Word-verified
  shape (`3-fork-output/07`), and every Python reader accepts it, so only a test pins it.
- **Write refusal** at a nonzero offset, asserting the destination is unchanged.
- **Truncation still happens at offset 0**, asserting output is byte-identical to a fresh save
  and carries no tail of a longer prior document.
- **`patch_save` verbatim gate** — clean original still byte-identical on no-op; prefixed
  original refused at load.
- **`diagnose()`** reports a prefixed archive without raising, with the distinct `kind`.
- **No-regression sweep**: the 28 fixtures in `1-refused-inputs/` and the seven Word-confirmed
  artifacts in `3-fork-output/` keep their current accept/refuse verdicts. Validated already —
  the load check flags exactly one of 37 fixtures.

### Two existing tests must change
- `tests/paper/test_practical_opc_hardening.py::it_validates_the_real_prefix_of_a_nonzero_position_stream`
  asserts the prefix is preserved *and* that the whole stream reopens as a `Document`. It
  encodes the bug. Invert it: expect `OSError` and unchanged original bytes.
- `tests/paper/test_practical_opc_hardening.py::it_restores_a_seekable_stream_after_commit_error`
  does `stream.seek(7)`, so it would be refused before it could exercise rollback. Move it to
  offset 0; rollback is still exercised via the snapshot path.

### What does not need tests
Constant definitions, the deleted `_copy_stream_prefix`, and re-proving that stdlib `zipfile`
accepts a prefixed archive.

## Risks and mitigations
- **Risk:** ~~the byte-0 rule rests on a single Word observation~~ **Retired.** Word has now
  refused the shape on two independently built fixtures — `3-fork-output/07` (fork output) and
  `1-refused-inputs/06-prefix-data` (a mutated input). The rule has two data points.
- **Risk:** both threads edit `_zipguard.py` and the same two tests. **Mitigation:** let the
  hardening thread own the load check and the truncation gate; this branch takes only the
  write-side refusal and the verbatim gate, and rebases after theirs lands.
- **Risk:** a caller legitimately embeds packages today and this breaks them.
  **Mitigation:** there is no Word-openable output from that path, so the operation was
  already broken; the refusal makes it loud. Message names the alternative.
- **Risk:** an exotic-but-valid `.docx` does not start with a local file header.
  **Mitigation:** measured across 37 fixtures — zero false positives. Archives starting with a
  multi-disk spanning marker are already refused elsewhere.

## Decisions made during the build

- **Exception types, as specced.** Load refusals raise `PackageLimitError`; the write refusal
  raises `OSError`, matching append-mode and every other destination refusal in
  `opc/package.py`. `word-oracle-fixes` L4 asked for "a typed refusal" on the write side, which
  would mean a `PaperRefusal` subclass; that was not adopted, because a destination's cursor
  position is a property of the argument, not of the archive.
- **The `patch_save` verbatim-copy gate was dropped**, replaced by a coupling test asserting
  `patch_save` refuses a prefixed original. Not because it is dead code — the truncate gate is
  equally dead — but because it would need a loader bypass to exercise at all, whereas the
  truncate gate sits in a path every offset-0 save runs.
- **The truncate gate was kept**, as an explicit `if start == 0`, because `truncate()`'s
  no-argument semantics are subtle enough to be worth stating.
- **No new `PackageDiagnosis.kind` value.** Both prefixed shapes report `unsafe-archive`, which
  keeps `kind` a closed set for callers switching on it and makes `diagnose` consistent across
  the two shapes; the distinguishing detail lives in `problems`.
- **`compare()` needs no separate regression test.** It routes through `_read_zip`
  (`_compare.py:314`), so preflight coverage reaches it, as it does `Document()`, `patch_save`
  (original and its own re-read), `diff_package` and `diagnose`.
- **Prefix length is not reported** in the refusal message. `_scan_central_directory` returns
  nothing and tracks no offsets; threading a value through it would buy wording, not correctness.

## Validation corpus

False-positive checking for the byte-zero rule used the **87 git-tracked `.docx` files** — 45 in
`tests/`, 41 in `features/`, 1 in `src/`. Every one still opens except the two under
`tests/paper/fixtures/generated/corrupt/`, which were already refused before this work and for
unrelated reasons. (`rglob` also finds `features/_scratch/test_out.docx`, a gitignored behave
artifact — not a corpus document.) The parallel specs cite 83 and 60 for the same check; the difference is the set of
scanned paths, not a disagreement about the rule.

The fixture replay covered **67 fixtures across five recorded-verdict sets**. Exactly two changed
status, both intended: `4-followups/06-prefix-data.docx` and
`3-fork-output/07-container-WHOLE-FILE.docx`, each accepted → refused. Note the set names:
`1-refused-inputs/06-` is `06-concatenated.docx`, a different shape that was already refused, and
`3-fork-output/08` was already refused too — neither moved.
