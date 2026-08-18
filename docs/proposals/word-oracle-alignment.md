# Word-oracle alignment

Base: `7e7cb51` (tip of the #23 → #24 → #25 → #13 → #22 stack).
Evidence: Word for Mac, 2026-08-18, 31 documents opened by hand. Ledger:
`verifying-against-word/WORD-VERDICTS.md`.

## The principle this spec applies

paper-docx exists because python-docx lets you build a document that looks fine in Python and
fails to open in Word. Closing that gap is the only thing that justifies diverging from
upstream. So each divergence has to fall into one of two buckets:

- **Justified** — upstream accepts or writes something Word rejects, and the fork refuses it.
- **Unjustified** — the fork refuses something Word opens. That is a regression against users
  with no purpose, and it should be deleted.

There is a third case the audit turned up that is worse than either: the fork *accepts*
something Word rejects, reports correct content for it, and will edit and re-save its own
misreading. That is the exact failure the package was built to prevent, occurring inside the
package. Those get priority.

## State at this commit — measured, not assumed

Every fixture set was replayed against this commit's code.

| # | Finding | Direction | Status here |
|---|---|---|---|
| 1 | `MAX_COMPRESSION_RATIO` | fork too strict | **fixed** by #24 |
| 2 | `MAX_XML_MEMBER_BYTES` | fork too strict | **fixed** by #24 |
| 3 | `MAX_MEMBER_COUNT` | fork too strict | **fixed** by #24 |
| 4 | image relationship content type | fork too strict | open |
| 5 | `prefix_data` accepted | **fork too lenient** | open |
| 6 | part-name character rule | **fork too lenient** (×4 spellings) | open |
| 7 | `prefix_collision_plain` accepted | **fork too lenient** | open |
| 8 | container-tail truncation | data loss the fork introduced | open |
| 9 | protection gate ignores mode | fork too strict | open |
| 10 | three refusal messages | wording | open |

#24 dropped all eight caps, not just the three measured. The other five were never Word rules
either, so that is the right direction; it is simply unmeasured.

Confirmed still correct and **not** to be touched: the ZIP footer rule (4 shapes), all three
relationship-id guards, core-part content-type strictness, `case_ambiguous_names`, and the
append-mode refusal — upstream silently writes a corrupt archive through an append handle and
cannot reopen its own output, so that `OSError` is load-bearing.

---

## Fix 1 — part-name characters (finding 6)

**File:** `src/docx/_zipguard.py`, `_validate_member_name`.

Word refused four of the five renamed-media spellings it was shown. paper-docx refuses exactly
one of them — the raw decomposed form — and opens the other three plus the literal space. So the
NFC check is aimed at the wrong property: **non-ASCII in a part name is the defect**, and
normalization form is not.

Replace the NFC check with an OPC part-name legality rule:

1. A member name may contain only ASCII characters.
2. Unescaped characters must be legal in a URI path (RFC 3986 `pchar` plus `/`).
3. A `%XX` escape is allowed only when it encodes a byte below `0x80`. Existing rejections for
   escapes of unreserved characters, `/`, `\`, NUL, DEL and controls stay.
4. **`[Content_Types].xml` is exempt.** It is a reserved ZIP item name, not an OPC part name.

Rule 4 is not optional. Across 83 real documents in this repo it is the *only* member name
outside the URI-path set — `[` and `]` are not `pchar`. Without the exemption every document
breaks.

Delete the NFC check entirely rather than keeping it alongside: a raw decomposed name is
non-ASCII, so rule 1 subsumes it. That removes the `unicodedata` import.

**Validation already run.** The rule reproduces all seven Word verdicts exactly and rejects
none of the 52 distinct member names across those 83 documents.

| name | rule | Word |
|---|---|---|
| `word/media/renamed1.png` | accept | OPENS |
| `word/media/my%20image.png` | accept | OPENS |
| `word/media/my image.png` | reject | REFUSES |
| `word/media/imagé1.png` (NFC and NFD) | reject | REFUSES |
| `word/media/imag%C3%A91.png` | reject | REFUSES |
| `word/media/image%CC%811.png` | reject | REFUSES |

**Message:** `"part name 'X' contains a character that is not legal in an OPC part name"`, naming
the character. The current text — "is not Unicode-normalized" — describes a form, not the defect.

---

## Fix 2 — the archive must begin at offset zero (finding 5)

**File:** `src/docx/_zipguard.py`, `_preflight_zip_stream`.

The most serious item. Word **refuses** a package with a stub in front of it; paper-docx opens it
and reports the *correct* content. Nothing downstream can tell the document is unusable.

The preflight checks that the central directory is internally consistent
(`central_offset + central_size == central_end`), which a rebased prefix satisfies. It never
checks where the archive starts.

Add: the first four bytes of the stream must be a local file header signature (`PK\x03\x04`), or
an end-of-central-directory signature (`PK\x05\x06`) for an empty archive.

**Validation already run.** Zero violations across 83 real documents; catches `prefix_data`;
the control passes. It does not disturb `stray_signature` (about the EOCD scan, not offset zero)
or `concatenated` (already refused for an ambiguous central-directory region).

**Message:** `"the package does not begin at the start of the file: N bytes precede the archive"`.

---

## Fix 3 — directory-prefix collisions (finding 7)

**File:** `src/docx/_zipguard.py`, `GuardedZipReader._validate_metadata`.

Word refuses all three collision shapes — a zero-length member named `word` beside members under
`word/` — **regardless of the member's attributes**. paper-docx refuses only the two carrying
directory attributes and accepts the plain one.

This inverts what the code currently assumes. The check fires on `info.is_dir() or external_attr
& 0x10`, i.e. on mode bits. Word says the mode bits are incidental and **the name is the defect**.

Add, for members whose names do not end in `/`: refuse when the name is a directory prefix of any
other member name. Keep the existing attribute check — a member claiming to be a directory is
still wrong — but it is no longer what carries this case.

**Validation already run.** Zero violations across 83 real documents; catches
`prefix_collision_plain`; the control passes. It must not apply to the `_directory_infos` set:
zero-length folder records like `word/` are confirmed **OPENS** in Word and must keep working.

**Message:** name the collision — `"member 'word' collides with the directory prefix used by
'word/document.xml'"` — which is what the current wording already claims and now actually means.

---

## Fix 4 — image relationship content types (finding 4)

**File:** `src/docx/opc/pkgreader.py`, `_validate_relationship_target_content_type`.

Word **opens** a displayed image declared `application/octet-stream`, and **refuses**
`styles.xml` declared `application/xml`. It cares what a core part claims to be; it does not care
what a media part claims to be.

Delete the `RT.IMAGE` branch. Images then fall through to the `_KNOWN_RELATIONSHIP_CONTENT_TYPES`
loop, which has no image entry, so no check applies. Keep every core-part entry unchanged — three
separate Word verdicts confirm that strictness is correct.

This is the one remaining unjustified regression on the read path, and the likeliest of all of
them to be hit by real third-party output, since a generic fallback media type is normal for a
non-Microsoft producer.

---

## Fix 5 — container-tail truncation (finding 8)

**File:** `src/docx/opc/package.py`, `_atomic_stream_write`, the bare `stream.truncate()`.

Not a Word question — the caller's own bytes are destroyed before any oracle sees the file.

| | tail preserved |
|---|---|
| upstream python-docx | 274,992 of 312,000 |
| paper-docx | **0** |

`truncate()` with no argument cuts at the current position, which is after the package, whatever
offset the write began at. Upstream never truncates, so it preserves the tail.

Condition truncation on the write having started at offset zero. Truncating is correct when
overwriting a whole file and destructive when writing into the middle of someone else's.

    if start == 0:
        stream.truncate()

Both behaviours are already covered by fixtures: `3-fork-output/03` and `/04` (overwrite a longer
document — must stay truncated, or they grow a stale tail that Word rejects) and `/07` (container
— must keep the tail).

---

## Fix 6 — protection gate: mode and operation aware (finding 9)

**File:** `src/docx/protection.py`, `_refuse_if_protected`.

The gate reads `enforced` and refuses. Its `operation` parameter is used only to build the
message and never affects the decision, so all five restriction modes behave identically across
all 23 call sites.

Word's rules are per-mode *and* per-operation-class. Measured:

| mode | comments | form fields | body |
|---|---|---|---|
| `readOnly` | blocked ✓measured | blocked | blocked |
| `comments` | **ALLOWED** ✓measured | — | blocked ✓measured |
| `trackedChanges` | **ALLOWED** ✓measured | — | permitted *(inferred)* |
| `forms` | blocked ✓measured | **ALLOWED** ✓measured | blocked |
| formatting-only | **ALLOWED** ✓measured | — | permitted *(inferred)* |

Make `operation` load-bearing by classifying the 23 call sites:

- **comments** — add/edit a comment, anchor a comment, delete a comment, reply to a comment,
  resolve a comment
- **form-field values** — set a control value
- **body content** — everything else, including resolve a revision / resolve revisions

`acknowledge_protection()` stays the escape hatch for genuinely blocked cases.

**Message:** every B-file in the protection set opened cleanly in Word with the comment visible,
so this refusal is a *policy* mirror of Word's UI, not a corruption guard. The message must say
Word's own UI would not permit the edit — not imply the document would break.

**Open decision, flagged rather than assumed.** Two body-content cells are inferred from Word's
documented semantics, not observed: body edits under `trackedChanges` and under formatting-only.
Under the principle at the top of this spec, an unmeasured cell has no evidence to justify
refusing, and protection is not a corruption guard — so permitting matches upstream and risks
nothing worse than doing something Word's UI would not let a human do. Either measure those two
cells first, or ship them permissive deliberately. **Do not ship them refusing by default**,
which is a regression with no evidence behind it.

---

## Fix 7 — three refusal messages (finding 10)

| shape | current | problem |
|---|---|---|
| `duplicate_default_ct` | "contains an ambiguous Default declaration" | two entries declaring the *same* type are not ambiguous. Say "duplicate Default declaration for extension 'png'". Word does refuse the file, so only the wording is wrong. |
| the four footer refusals | describe the end-of-central-directory record | say "the file has bytes after the end of the archive; re-save the document" |
| `duplicate_rel_ids` vs `conflicting_rel_ids` | identical text | one is a redundant declaration, one a contradiction. Both are refused; what the caller does next differs. |

---

## Not in this spec

**The reachability gap.** paper-docx opens `undeclared_orphan_part`, a top-level `.DS_Store`, and
a `__MACOSX/._word` sidecar — all unreferenced members that resolve to no content type. Reachable
undeclared parts are already refused correctly. Word has not been asked about the unreferenced
ones. Given that Word proved strict about both core-part types and part-name characters, and that
the fork has now been too lenient in three measured places, the prior is that these are a fourth.
Fixtures exist in `1-refused-inputs/`. **Measure before writing this fix**, and expect it to
interact with Fix 1 — one member-level legality rule may cover both.

**The rendering axis.** Numbering, tracked revisions, fields, `compare()`. Zero measurement, and
a different question: not "did it open" but "did Word draw what the library claims".

## Order

1. Fix 5 (truncation) — data loss, self-contained, no Word dependency.
2. Fixes 1, 2, 3 — the leniency bugs, all in `_zipguard.py`, all validated against 83 documents.
3. Fix 4 — one branch deleted.
4. Fix 7 — wording.
5. Fix 6 — largest surface (23 call sites) and carries the open decision above.

## Test plan

- Regression fixtures: every document in `1-refused-inputs/`, `4-followups/` and `5-encoding/`
  must land on its recorded Word verdict. That is the acceptance criterion for fixes 1–4.
- Real-document guard: the 83 `.docx` files in the repo must all still open. The three new rules
  were validated against them and produced zero false positives; keep that as a test.
- `3-fork-output/03`, `/04`, `/07` for fix 5.
- Baseline at this commit is 2,644 passing. The 9 collection errors in
  `tests/opc/test_phys_pkg.py` are a pytest deprecation about class-scoped fixtures declared as
  instance methods, unrelated to this work.
