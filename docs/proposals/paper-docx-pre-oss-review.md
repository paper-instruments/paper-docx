# paper-docx audit

Package: [paper-docx](https://github.com/paper-instruments/paper-docx) `a55be76` (0.1.2), vs tag `paper-base` (`e454546`). `git diff paper-base -- src`: 44 files, +15244/−64.

Stock python-docx is the create/edit baseline. This fork is for **agent Word work on real files**: inspect, edit, review, structure, combine, save, deliver. The agent should use the package, not unzip and patch XML.

Provenance for “was this asked for” is `paper-original-plans-and-specs-2026-08-13/paper-docx` (plans + reference harness). **Place?** is whether it belongs in the package, not whether that folder named it.

## Change order

Drop the zip-bomb caps. Keep `PackageLimitError` for a broken zip, and the rest of Place Yes. Fix the Defects table. Then the Missing table. Word fixtures stay parked.

**Drop**

- ~~Every zip-bomb size/count/ratio cap.~~ **Done.** Names gone: `MAX_COMPRESSED_BYTES`, `MAX_MEMBER_COUNT`, `MAX_CENTRAL_DIRECTORY_BYTES`, `MAX_XML_MEMBER_BYTES`, `MAX_BINARY_MEMBER_BYTES`, `MAX_TOTAL_EXPANDED_BYTES`, `MAX_COMPRESSION_RATIO`, `RATIO_ENFORCEMENT_FLOOR_BYTES`. A large but valid `.docx` opens.
- ~~Stop raising `PackageLimitError` for those caps.~~ **Done.** That class still means corrupt, encrypted, or malformed zip.

**Keep**

- `PackageLimitError` when the file is not a usable Word zip: corrupt, encrypted, or malformed.
- Typed refuse on that broken zip, atomic save, roll back a failed edit.
- Every **Place Yes** row.

**Fix**

- ~~`patch_save` writer (symlink + `fsync`)~~ **Done.**
- ~~Rollback on `blocks` / `fields` / bookmarks / `Comment.author` / `initials`~~ **Done.** Author/initials use cheap XML rollback.
- ~~Composition `assert ... or True`~~ **Done.**
- ~~Mixed-install identity check on Paper-only modules, plus a transitive-clobber CI install~~ **Done.** After upstream overwrites `docx/__init__.py`, `from docx import Document` still loads stock python-docx; run the doctor for that path.

**Missing**

- Delete one comment
- Modern comment identity parts (`commentsIds.xml`, `commentsExtensible.xml`)
- Lock a file for the recipient
- Replace one existing picture
- Create or retarget a hyperlink
- Auto-number caption field
- Add a footnote or endnote
- Fill a data-bound form control
- Append and keep the source letterhead

**Parked**

- Sample files authored in desktop Microsoft Word. Tests today use generated files, LibreOffice files, and one Google Docs export. Real Word markup is unproven until someone who has Word adds those files.

## Every Paper addition: does it have a place?

Read **What it does** first. The API column is only the name in the library.

| What it does | API | Use case | Place? | Why |
| --- | --- | --- | --- | --- |
| Read the whole file (body, headers, footers, footnotes, comments), including text that is still a tracked insertion or deletion. | `docx.story` | Edit existing, review | Yes | Stock only sees the main body and ignores pending revisions. Notes can be read, not created (see Missing). |
| Find a quoted phrase even when Word split it across bold/italic runs, and replace it without wiping that formatting. Can stamp the edit as a tracked change. | `docx.search`, `Span.replace` | Edit existing, review | Yes | Stock `paragraph.text` wipes run formatting. Agents otherwise splice XML. |
| Insert a heading and paragraphs, a real list, or a simple table after a known place. Mark whole paragraphs deleted or replaced as Word track-changes. | `docx.blocks` | Edit existing, review | Yes | Stock has no safe paragraph-level tracked edit. |
| List every tracked change and accept or reject it (inserts, deletes, moves, some formatting). | `Document.revisions` | Review | Yes | Stock cannot list or resolve tracked changes. Rare types (table-property, cell, section, numbering, custom XML) can be listed but not resolved. |
| Comment on an exact phrase, reply, and resolve the thread. | `Span.comment`, `commentops` | Review | Yes | Stock comments attach to whole runs, with no reply or resolve. The original plan asked for add, reply, and resolve. Deleting one comment was not asked for (see Missing). |
| Diff two `.docx` files into a Word redline of the text. | `docx.package.compare` | Review | Yes | Review job, not every session: two files in, a redline a person opens in Word. Images, fields, controls, and formatting-only diffs refuse. |
| Change one table cell, or insert/delete a row, without breaking a merged header. | `docx.tableops` | Structured | Yes | Stock cell/row edits scramble real tables. |
| Apply a real Word numbered or bulleted list, not fake Unicode bullets. Restart numbering when needed. | `docx.numbering` | Structured | Yes | Needed when inserting blocks into existing lists. |
| Fill a form control: text, checkbox, dropdown, or date. | `docx.controls` | Structured | Yes | Stock cannot fill content controls safely. Data-bound controls were asked to refuse, not to fill. Picture controls also refuse (see Missing). |
| Put a bookmark on a phrase, list bookmarks, or remove the markers (the text stays). | `docx.bookmarks` | Structured | Yes | Anchors for cross-references and TOC in existing files. |
| Insert a page number, date, cross-reference, or table of contents. The displayed value is computed when Word opens the file, not here. | `docx.fields` | Structured | Yes | Stock is weak; agents otherwise hand-edit field XML. Page numbers, dates, cross-references, and TOC were asked for. Auto-number captions (Figure 1, 2, …) were not (see Missing). |
| Copy a range, or a whole document, into another file, keeping styles, lists, and images. | `docx.composition` | Combine | Yes | Cross-file copy is where styles and relationships corrupt. Headers stay those of the destination file; keeping the source letterhead was left for later. Will not copy comments, pending revisions, or footnotes. |
| Save an edited file without rewriting ZIP parts that did not change. | `docx.package.patch_save` | Package | Yes | Stock save rewrites the whole ZIP, which noisily diffs and can break unrelated parts. |
| On open and save: refuse a broken ZIP with a typed error, write atomically, and roll back a failed edit instead of leaving a half-written file. | zipguard, atomic `save`, transactions, `PaperRefusal` | Package | Yes | Load/save plumbing. Keep. The zip-bomb *size caps* are gone (see Change order). A member that inflates past the size it declared still refuses. |
| Say why a file will not open: encrypted, template, macros, not Word, damaged ZIP. | `docx.package.diagnose` | Package | Yes | Stock raises untyped errors on bad input. |
| Show which ZIP parts or visible text changed between two files, or what pending revisions would change if accepted. | `diff_package`, `text_diff`, `pending_changes` | Package | Yes | Proof of what an edit did. Library belongs; not a ritual after every call. |
| Strip comments and author metadata for a clean outgoing copy. Optional: RSIDs and hidden text. Does not remove Restrict Editing. | `Document.scrub` | Deliver | Yes | Send-clean step, not an everyday edit: strip comments and author names before the file leaves. Keep; already shipped. |
| Accept or reject every tracked change, then check that none remain, or refuse and undo. | `Document.finalize` | Deliver | Yes | Accept All (or Reject All), then prove the file is not still redlined. If leftover markup remains, refuse and roll back. That check is not something you can do after `accept_all`. Comments and author names are `scrub`, a separate call. |
| See whether Restrict Editing is on, and refuse Paper edits until you override in memory for this session. Cannot turn protection on or strip it from the file. | `docx.protection` | Deliver | Yes | The original plan asked to notice a lock and refuse, never to strip it. A setter to lock a file for a recipient was not asked for (see Missing). |
| Point at text and ask what Word is actually showing: bold, italic, font, size, color, highlight, alignment, paragraph style. Those can live on the text, on a style, or on the document default; this walks all three and says which one won. If a phrase is half bold and half not, it says mixed. It does not answer spacing, indent, borders, table coloring, or list-number look; those come back as “we did not compute this” instead of a guess. | `docx.formatting` | Edit existing | Yes | The original plan asked for this inspect, and asked it to name what it cannot resolve rather than guess. The short list plus “not computed” is what was asked for, not a half-built inspect. A helper to match nearby formatting when inserting exists, but copy/insert do not call it yet. |
| Check that `import docx` is this fork, not a mixed install with python-docx. | `paper-docx-doctor` | Package | Yes | Both packages own the import name `docx`. Keep the doctor. Paper-only modules also refuse a mixed install. CI covers both install orders, leftover uninstall, and a dummy package that pulls `python-docx` as a dependency. |


---

## Missing

Jobs the agent still could not do through the package before this follow-up. This branch adds public APIs for each row.

| What the agent needs to do | Use case | Why the package does not cover it | Found in | Plans |
| --- | --- | --- | --- | --- |
| Delete one comment and leave the rest | Review | You can add a comment, reply, mark it resolved, or strip every comment when delivering. You cannot remove a single comment. That still means deleting the range markers and the comments.xml entry by hand. | Code: `commentops` has add/reply/resolve/thread, no delete. | **Not asked.** Original plan is add, reply, resolve only. |
| Add a comment that current Word will show and round-trip | Review | Word stores comment identity in extra package parts, not only the comment text. This fork writes the two older parts (`comments.xml` and `commentsExtended.xml`). Without the two newer ones (`commentsIds.xml`, `commentsExtensible.xml`), a comment can fail to appear or fail to round-trip after Word opens the file. | Anthropic `skills/docx/SKILL.md` (helper writes six parts). Code: we only create `commentsExtended`. | **Not asked.** Original plan is the older extra comments part only. The six-part list is from the Anthropic skill, not this folder. |
| Lock a file for the recipient (read-only, comments only, or forms only) | Deliver | You can see that Restrict Editing is already on, and Paper will refuse to edit (or you can override that refusal in memory for this session). You cannot turn protection *on* for delivery. Authoring a locked file still means writing the lock yourself. | Codex `set_protection.py`. Code: `protection.py` has no setter. | **Not asked.** Plan: report it, never strip it, refuse edits while it is on. No setter. |
| Replace one existing picture, keep size and position | Edit existing | You can insert a new picture (stock) or copy pictures along when combining documents. You cannot swap the bytes of one picture already in the file without sharing that image part with other shapes. Picture form controls also refuse to be filled. | Anthropic skill (insert/replace images). Codex images task. Code: no replace API; `controls` refuses `picture`. | **Later, if someone asks.** Image swap was parked until a real demand. Not in the v0.11 schedule. |
| Turn a phrase into a hyperlink, or change where an existing link goes | Edit existing, review | You can read a link's URL. You cannot create a link, retarget one, or record that as a tracked change. External links are copied only when you paste a range from another file. Editing a link in place still means XML. | Anthropic edit path is XML. Codex `hyperlinks_and_fields`. Code: `Hyperlink.address` is read-only. | **Later, if someone asks.** The plan asked to edit text *inside* an existing link, not to create or retarget one. |
| Caption a figure or table so numbers update (Figure 1, Figure 2, …) | Structured | You can bookmark the caption text and insert a cross-reference to that bookmark. Word captions are different: they use an auto-number field, so later figures renumber. This package cannot insert that field, so numbers stay as typed text. | Codex `captions_crossrefs`. Code: `fields.py` has PAGE, NUMPAGES, DATE, REF, TOC only. | **Not asked.** Fields in the plan are page numbers, date, cross-references, and TOC. “Captions” in the skill is writing advice, not an auto-number field. |
| Add a footnote or endnote | Edit existing | If a note already exists, you can find and edit its text. You cannot add a new note or the superscript mark in the body. | Codex `footnotes_endnotes`. Code: footnote/endnote parts load for reading; they are never created. | **Explicitly out.** Read is in. Creating notes was deferred until a legal or academic customer asked. |
| Fill a form field whose value lives in Word's hidden custom XML | Structured | Ordinary controls (text, checkbox, dropdown, date) can be filled. Some templates bind the control to a hidden part; Word overwrites the visible text from that part on open. Filling the surface would look like it worked and then vanish. There is no API to write the bound part. The package refuses instead. | Codex `forms_content_controls`. Code: `set_value` refuses data-bound controls. | **Asked as refuse.** Intentional no: refuse data-bound controls, do not fill the hidden store. |
| Append another document and keep *its* letterhead | Combine | You can append another document's body, keeping styles and images. Headers and footers stay those of the destination file. If both files have a letterhead, the source one is dropped. | Codex `multi_doc_merge`. Code: `append_document` docstring. | **Asked as future.** Destination headers on purpose; keeping the source letterhead was left for later. |

Not a package hole (stock, QA outside the library, or a different product): insert image (`Run.add_picture`), edit header/footer text (stock + story/search), page/section breaks (stock `add_section`), fill existing non-bound form controls, visual PNG render, flatten-runs, redaction, a11y audit, watermarks, Google Docs title sanitizer, `.doc` conversion. Public document-QA was asked to stay out of this library.

---

## Defects in what we already shipped

These are not new Word jobs. They are holes in APIs that already exist. Source: [#12](https://github.com/paper-instruments/paper-docx/pull/12), checked against `a55be76` where noted.


| Defect                                                                    | Why it matters                                                                                                                                                                                                                                                                                    | Found in                                                                                         |
| ------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| ~~`patch_save` does not use the same path writer as `Document.save`~~ **Done.** | Follows a destination symlink, `fsync`s the file and directory, refuses if the link changes during save. Same idea as `Document.save`. | [#12](https://github.com/paper-instruments/paper-docx/pull/12) P0-1                              |
| ~~Refusal/rollback is not a closed contract for every public mutator~~ **Done.** | `blocks.py` public mutators, `fields.py` inserts, and `create_bookmark` / `delete_bookmark` wrap `rollback_on_error`. `Comment.author` / `initials` wrap `rollback_xml_on_error` (one attribute set; a full-package snapshot was thousands of times slower and bought nothing). | [#12](https://github.com/paper-instruments/paper-docx/pull/12) P0-2                              |
| ~~Mixed `paper-docx` / `python-docx` install can leave a broken `docx` tree~~ **Done for Paper-only modules and CI.** | Doctor kept. Paper-only modules call `docx._guard.check_install`. CI covers both install orders, leftover uninstall, and a dummy package that pulls `python-docx`. After upstream overwrites `docx/__init__.py`, `from docx import Document` still succeeds as stock python-docx; the doctor is the check for that path. | [#12](https://github.com/paper-instruments/paper-docx/pull/12) P0-3                              |
| No desktop-Word sample files (parked) | Tests use generated files, LibreOffice files, and one Google Docs export. Word-authored tracked changes are still unproven. Needs someone with desktop Word; not this change. | [#12](https://github.com/paper-instruments/paper-docx/pull/12); `tests/paper/fixtures/README.md` |
| ~~Composition test always passes~~ **Done.**                              | `assert "Heading2" not in report.imported_styles` is a real assertion.                                                                                                                                                                                       | [#12](https://github.com/paper-instruments/paper-docx/pull/12)                |
| ~~Zip size/count/ratio caps refuse large but valid files~~ **Done.**      | Caps removed. `PackageLimitError` remains for corrupt, encrypted, or malformed zip. A member that inflates past the size it declared is refused while inflating, not after the whole blob is in memory. | `src/docx/_zipguard.py`                                                                          |


Left in [#12](https://github.com/paper-instruments/paper-docx/pull/12) and not copied here: README claim matrix, Ruff/Pyright CI, ZIP fuzzing, determinism hashes, upstream-sync process. Those are docs/process, not package capability.

---

## Constants, enums, and error classes vs python-docx

Not functions. Stock `WD_*` / `MSO_*` enums and stock error classes (`PythonDocxError`, `InvalidSpanError`, OPC and image errors) are unchanged. Nothing was removed. No new Enum types were added; callers get string lists instead.

**Change** is the decision for the work list: **Keep** or **Drop**. Dropped zip-bomb constants are gone from the code.

Provenance is `paper-original-plans-and-specs-2026-08-13/paper-docx` (plans + reference harness). **Asked for** = named there. **Supported** = follows a plan rule, but the name or extra values were not written down. **No support** = not in that folder.

| Name | Kind | Change | What it is | Why | Plans |
| --- | --- | --- | --- | --- | --- |
| `PaperRefusal` | Error | Keep | Base. Edit was refused; file and memory unchanged. | Agents catch “safe no” separately from a bug. | **Asked for.** `CONVENTIONS.md` pins this name. |
| `PackageLimitError` | Error | Keep | Typed error when the file is not a usable Word zip: corrupt, encrypted, or malformed. | Agents catch a broken package as a safe no. Size caps used to raise this class; they are gone. A member that inflates past the size it declared still raises this class. | **Supported, not named.** Plans asked for typed refuse on a corrupt or encrypted zip, not this class name, and not size caps. |
| `AmbiguousTargetError` | Error | Keep | Search matched more than one place. | Do not guess which hit to edit. | **Asked for.** `CONVENTIONS.md`; `PLAN-paper-docx.md`. |
| `TargetNotFoundError` | Error | Keep | Nothing matched, or the span went stale. | Do not invent a target. | **Asked for.** `CONVENTIONS.md`; `PLAN-v0.1.md` H7. |
| `UnsupportedStructureError` | Error | Keep | This edit is not supported on that Word structure. | Refuse instead of a quietly wrong file. | **Asked for.** `CONVENTIONS.md`; `PLAN-v0.1.md` H3/H4/H8. |
| `BoundaryViolationError` | Error | Keep | The edit would cross a paragraph, table, or control boundary. | Keep replacements inside one safe range. | **Asked for.** `CONVENTIONS.md`; `PLAN-paper-docx.md`. |
| `RelationshipPolicyError` | Error | Keep | The edit would create an unsafe package relationship. | Pinned in the error list for copy/rel work. Code defines it; nothing raises it yet. | **Asked for.** `CONVENTIONS.md` taxonomy. No docx plan names a call site. |
| `DocumentProtectedError` | Error | Keep | Restrict Editing is on. Acknowledge in memory to proceed. | Editing a locked template by accident looks successful and is wrong. | **Asked for.** `PLAN-v0.11` Phase 3 names this class. |
| `UnsupportedXmlError` | Error | Keep | XML we will not compare (for example a DOCTYPE). A `ValueError`, not a `PaperRefusal`. In `docx._paperpkg`. | Compare must not treat two parts as equal when a DTD could change the text. | **No support.** Not in the docx plans. |
| `DoctorError` | Error | Keep | Install is mixed or not paper-docx. A `RuntimeError`. In `paper_docx_doctor`. | Shared import name `docx` can hide a mixed install. Keep the doctor. | **No support** in the original plans (they asked for `__paper_version__`). **Keep anyway.** |
| `__paper_version__` | Constant | Keep | `"0.1.2"`. Stock still has `__version__ = "1.2.0"`. | Tell this fork apart from stock python-docx in the same import. | **Asked for.** `CONVENTIONS.md`; `SUMMARY-paper-docx-v0.11.md`. |
| `MAX_COMPRESSED_BYTES` | Constant | Dropped | Was 256 MiB zip-bomb cap. Gone. | Size limit on open. Refused large but valid files. | **No support.** Plans asked for typed errors on a *corrupt* zip, not a size cap. |
| `MAX_MEMBER_COUNT` | Constant | Dropped | Was 4096-member zip-bomb cap. Gone. | Same as above. | **No support.** |
| `MAX_CENTRAL_DIRECTORY_BYTES` | Constant | Dropped | Was 16 MiB zip-bomb cap. Gone. | Same as above. | **No support.** |
| `MAX_XML_MEMBER_BYTES` | Constant | Dropped | Was 64 MiB zip-bomb cap. Gone. | Same as above. | **No support.** |
| `MAX_BINARY_MEMBER_BYTES` | Constant | Dropped | Was 256 MiB zip-bomb cap. Gone. | Same as above. | **No support.** |
| `MAX_TOTAL_EXPANDED_BYTES` | Constant | Dropped | Was 512 MiB zip-bomb cap. Gone. | Same as above. | **No support.** |
| `MAX_COMPRESSION_RATIO` | Constant | Dropped | Was 100:1 zip-bomb cap. Gone. | Same as above. | **No support.** |
| `RATIO_ENFORCEMENT_FLOOR_BYTES` | Constant | Dropped | Was 16 MiB floor for the ratio cap. Gone. | Same as above. | **No support.** |
| `VIEWS` | Constant | Keep | `("current", "original", "all")`. Story read modes. | Search/replace must see what Word shows, what was deleted, or both. | **Asked for** `current` and `original` (`PLAN-v0.1.md` H1). **`all` not named.** |
| `RESOLVABLE_TYPES` | Constant | Keep | Revision types `accept`/`reject` can resolve. | H3: refuse the whole set if any selected type cannot be resolved. | **Asked for** as that rule (`PLAN-v0.1.md` H3). Constant name not written down. |
| `BLIND_REGION_KEYS` | Constant | Keep | Inspection keys for content story cannot fully read. | Do not claim “we saw the whole file” when math/OLE/hidden text is present. | **Asked for.** `PLAN-v0.1.md` H9; harness `BLIND_REGION_TAGS`. |
| `COMMENTS_EXTENDED_CONTENT_TYPE` | Constant | Keep | Word `commentsExtended` content type. Not in stock `CONTENT_TYPE`. | Reply/resolve live in that extra part. | **Asked for.** `PLAN-v0.1.md` V4 (`commentsExtended`). Constant name not written down. |
| `COMMENTS_EXTENDED_RELATIONSHIP_TYPE` | Constant | Keep | Word `commentsExtended` relationship. Not in stock `RELATIONSHIP_TYPE`. | Same as above. | **Asked for.** `PLAN-v0.1.md` V4. Constant name not written down. |
| Story view | String list | Keep | `current`, `original`, `all`. | Same as `VIEWS`. | **Asked for** `current`/`original`. **`all` not named.** |
| Resolvable revision type | String list | Keep | `insertion`, `deletion`, `format_change`, `row_insertion`, `row_deletion`, `move_from`, `move_to`. | Name what Word tracked so accept/reject can be honest. | **Asked for.** `PLAN-v0.1.md` H1–H3; `PLAN-v0.11` Phases 1–2 add rows and moves. |
| Listed but not resolvable | String list | Keep | `table_property_change`, `cell_revision`, `section_property_change`, `numbering_change`, `custom_xml_revision`. | Count them so “clean” is never a lie; refuse to resolve. | **Asked for.** `PLAN-v0.1.md` H2/H3. |
| Control type | String list | Keep | `text`, `rich_text`, `checkbox`, `dropdown`, `combo`, `date`, `picture`, `group`, `building_block`. | Fill the types Word templates use; refuse the rest. | **Asked for** text/rich_text/checkbox/dropdown/combo/date (`PLAN-v0.1.md` V1). **picture/group/building_block not named** (we enumerate and refuse). |
| Protection `edit` | String list | Keep | Word tokens: `readOnly`, `forms`, `comments`, `trackedChanges`, `none`. | Report the Restrict Editing mode the refusal named. | **Asked for.** `PLAN-v0.11` Phase 3. |
| Diagnose `kind` | String list | Keep | `missing`, `encrypted-or-legacy-binary`, `not-a-zip`, `unsafe-archive`, `corrupt-zip`, `docx`, `dotx`, `docm`, `dotm`, `xlsx`, `pptx`, `opc-unknown`. | `Document()` stays stock; this API says why open failed. | **Asked for** as `diagnose()` (`PLAN-v0.1.md` H10). The kind strings themselves are not listed. |
| Cross-reference kind | String list | Keep | `text`, `page`, `number` (REF / PAGEREF). | Insert a formula Word will renumber, not pasted digits. | **Asked for.** `PLAN-v0.11` Phase 6: text / number / page. |
