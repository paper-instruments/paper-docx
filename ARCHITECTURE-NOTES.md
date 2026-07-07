# Architecture Notes — Phase 0 orientation

Read-only deliverable for Phase 0 of `agent_docs/PLAN-paper-docx.md`. Fork point: upstream
`v1.2.0` (marker tag `paper-base`). Every path/line cited below was verified against this tree
on 2026-07-07. Both upstream suites were run **before** any exploration.

## Baseline verification

- `uv run pytest -q` → **1609 passed** — matches the `PAPER.md` baseline. Environment nuance:
  with the committed `uv.lock` environment (pyparsing 3.2.3, pytest 8.4.0) neither warning
  filter named in `PAPER.md` exists or is needed; the plain run is clean. `uv run behave -q` →
  **67 features / 650 scenarios / 1856 steps passed** — exact match to `PAPER.md`.

## The three layers (located by class name, not remembered paths)

- **opc layer** — `src/docx/opc/`: `OpcPackage` (`opc/package.py:24`), `Part`/`XmlPart`/
  `PartFactory` (`opc/part.py:21,207,165`), `Relationships`/`_Relationship` + `rId` allocation
  (`opc/rel.py:13,114,105`), reader (`opc/pkgreader.py`), writer (`opc/pkgwriter.py`), physical
  zip I/O (`opc/phys_pkg.py`). WML-specific `Package(OpcPackage)` lives in `src/docx/package.py`
  (see flag F1).
- **Save chain**: `Document.save` (`document.py:198`) → `DocumentPart.save` → `OpcPackage.save`
  (`opc/package.py:159`, calls each part's `before_marshal`) → `PackageWriter.write`
  (`opc/pkgwriter.py:31`) → `_ZipPkgWriter.writestr` (`opc/phys_pkg.py:116`, `ZIP_DEFLATED`).
  XML part blobs are produced lazily by `XmlPart.blob` → `serialize_part_xml`
  (`opc/oxml.py:53`): `etree.tostring(elm, encoding="UTF-8", standalone=True)`, no pretty-print.
- **Open chain**: `docx.api.Document` (`api.py:19`) → `Package.open` → `PackageReader.from_file`
  (`opc/pkgreader.py:20`) → `Unmarshaller.unmarshal` (`opc/package.py:186`) → `PartFactory`
  (selector fn, then `part_type_for[content_type]`, then generic `Part`) → `XmlPart.load` parses
  the blob and keeps only the element tree (original bytes are **not** retained — `patch_save`
  must re-read the original file for comparison, which its pinned signature already provides).
- **oxml layer** — `src/docx/oxml/`: `BaseOxmlElement` + metaclass `MetaOxmlElement`
  (`oxml/xmlchemy.py:643,93`); descriptors `ZeroOrOne`/`ZeroOrMore`/`OneOrMore`/`OneAndOnlyOne`/
  `Choice` and `RequiredAttribute`/`OptionalAttribute` all in `oxml/xmlchemy.py:111–583`. Each
  child descriptor auto-generates `{name}` / `{name}_lst`, `get_or_add_{name}`, `add_{name}`,
  `_insert_{name}`, `_remove_{name}`, `_new_{name}` methods on the element class.
- **`successors=` mechanism** (the schema-order encoder that replaces the reference helpers'
  ordered-insertion tables): `_insert_{name}` calls `insert_element_before(child,
  *self._successors)` (`oxml/xmlchemy.py:316–327`), which finds the first existing child whose
  tag appears in the successor list and inserts before it, else appends
  (`oxml/xmlchemy.py:664–670`).
- **Tag registration**: `register_element_cls` (`oxml/parser.py:32`) maps `w:tag` → class via
  lxml `ElementNamespaceClassLookup`; the single central registration hub is
  `src/docx/oxml/__init__.py` (~80+ registrations, grouped by subsystem). Namespaces/`qn()` in
  `oxml/ns.py`. The parser is built with `remove_blank_text=True, resolve_entities=False`
  (`oxml/parser.py:19`) — see flag F3.
- **api layer** — proxy bases in `src/docx/shared.py`: `ElementProxy` (:277), `Parented` (:319),
  `StoryChild` (:336); block-container mixin `BlockItemContainer` (`blkcntnr.py`). Proxies hold
  a live element ref + parent; `part` resolves through the parent chain. Public import surface
  is only `Document` (`__init__.py:20`); `__paper_version__ = "0.1.0"` at `__init__.py:17`;
  `src/docx/__init__.py` is also where `PartFactory.part_type_for[...]` part registration
  happens (:44–52).

## The comments feature (v1.2.0) — structural template for new vocabulary

- End-to-end chain: analysis doc `docs/dev/analysis/features/comments.rst` → oxml classes
  `CT_Comments`/`CT_Comment` with descriptor declarations (`oxml/comments.py`), registered in
  `oxml/__init__.py` → part `CommentsPart(StoryPart)` (`parts/comments.py:23`) with
  `default()` classmethod loading template `src/docx/templates/default-comments.xml`, wired
  lazily by `DocumentPart._comments_part` via `part_related_by(RT.COMMENTS)` /
  `relate_to(...)` (`parts/document.py:129–140`), loader-registered under `CT.WML_COMMENTS`
  (`__init__.py:46`) → proxies `Comments`/`Comment(BlockItemContainer)` (`src/docx/comments.py`)
  → exposure `Document.comments` / `Document.add_comment` (`document.py:161,41`) +
  `Run.mark_comment_range` (`text/run.py:176`) with the range-mark insertion done by CT_R
  methods (`oxml/text/run.py:91–102`).
- Commit discipline (from `git log v1.1.0..v1.2.0`): xfail behave acceptance test first
  (`5cb32d7`), then one micro-commit per property/method (`451747a` … `a809d6c`), analysis doc
  (`4fbe1f6`). Tests at all three layers: `tests/oxml/test_comments.py`,
  `tests/parts/test_comments.py`, `tests/test_comments.py`, plus 4 behave features with steps
  in `features/steps/comments.py`. This is the shape every organ PR here imitates.

## Where each v0 organ will live

- **Phase 2 kernel (`docx.package`)** — hooks needed already exist without touching opc
  internals: `OpcPackage.iter_parts()`/`iter_rels()` (`opc/package.py:69,46`), `Part.blob`/
  `partname`/`content_type` (`opc/part.py`). But the pinned module name collides with upstream
  `src/docx/package.py` — flag F1, human decision in PR-0.
- **Phase 3 traversal/inspection** — new proxy module(s) at the import root (pattern:
  `src/docx/comments.py`) + new part classes `parts/footnotes.py`/`parts/endnotes.py`
  (currently **no** `FootnotesPart`/`EndnotesPart` exists anywhere, though `CT.WML_FOOTNOTES`/
  `WML_ENDNOTES` and `RT.FOOTNOTES`/`ENDNOTES` constants already do — `opc/constants.py:138,141,221,231`)
  + registrations in `src/docx/__init__.py`. Headers/footers already reachable via
  `Section.header/.footer` (`section.py`); comments via `Document.comments`.
- **Phases 4–5 (find_text / Span / run-preserving replace)** — new modules at the import root,
  consuming Phase 3; run/text machinery to extend additively lives in `src/docx/text/run.py`,
  `text/paragraph.py` and `oxml/text/run.py` (`CT_R.inner_content_items`, run splitting will
  need new CT_R/CT_P oxml methods, never hand-built lxml in proxies).
- **Phase 6 (tracked replace)** — first new-vocabulary PR. `w:ins`/`w:del`/`w:delText` have
  **zero** oxml classes or registrations today (verified: no `CT_Ins`/`CT_Del` anywhere; no
  `w:ins` in `oxml/__init__.py`) → new `oxml/revision.py` (or similar) + registration +
  descriptor-declared children, imitating the comments commit shape. `w:rPrChange`-style
  successors already appear in existing `_tag_seq` tuples (e.g. `oxml/text/parfmt.py`), so
  ordering data is present where needed.
- **Phase 7 (block ops)** — anchor-relative insertion composes on `BlockItemContainer`
  (`blkcntnr.py`) and CT_P/CT_Tbl siblings; same-parent safety rule enforced at the oxml level.
- **Phase 8 (revisions)** — `doc.revisions` proxy module at import root over Phase 6 vocabulary,
  enumerating across Phase 3's story parts.
- **Phase 9 (tables/numbering)** — tables: additive methods on `Table`/`_Cell`/`_Row`
  (`src/docx/table.py`; `vMerge`/`gridSpan` guards at `oxml/table.py`). Numbering: oxml classes
  `CT_Numbering`/`CT_Num`/`CT_NumPr` exist (`oxml/numbering.py`), `NumberingPart` is a stub —
  `new()` raises `NotImplementedError` (`parts/numbering.py:14`) and `DocumentPart.numbering_part`
  auto-creates via that stub, so "apply existing definition" must go through the loaded part,
  never `new()`.
- **`docx.errors`** — name is free (upstream uses `exceptions.py` at both root and oxml level;
  no `errors` module exists), so the pinned `PaperRefusal` hierarchy can land as
  `src/docx/errors.py` with no collision.

## Confirmed blind regions (the gap Phase 3 must close)

- Block level: `CT_Body.inner_content_elements` is xpath `./w:p | ./w:tbl`
  (`oxml/document.py:88`) — its own docstring admits elements "shaded by nesting in a `w:ins`
  or other 'wrapper' element will not be included". Same pattern for `CT_Tc` (`oxml/table.py:493`)
  and `CT_HdrFtr`. Inline level: `CT_P.inner_content_elements` is `./w:r | ./w:hyperlink`
  (`oxml/text/paragraph.py:60`); `CT_R` text xpath covers 8 leaf tags, no field tags
  (`oxml/text/run.py:63–89`). Net: text inside `w:ins`, `w:del`/`w:delText`, `w:sdt` (block and
  inline), `w:txbxContent` (drawings are yielded as opaque `Drawing`), and field instructions is
  invisible to `.paragraphs`/`.runs`/`.text`/`iter_inner_content()`. Hyperlink text IS visible
  (since v1.1.0). Footnotes/endnotes are unreachable (no part class).
- Upstream has **zero** tests exercising any of these regions (all 65+ feature files and the
  pytest tree grepped; only schema-ordering references in `tests/oxml/test_table.py`) — so new
  opt-in traversal cannot regress upstream suites, and all safety coverage must come from
  `tests/paper/`.

## Test infrastructure facts

- pytest: RSpec-style discovery pinned in `pyproject.toml` — `python_classes = ["Test",
  "Describe"]`, `python_functions = ["it_", "its_", "they_", "and_", "but_"]`, `python_files =
  ["test_*.py"]`, warnings-as-errors (`filterwarnings = ["error", ...]`). `tests/paper/` will be
  auto-collected (not in `norecursedirs`) provided files/classes/functions follow those
  patterns. behave only discovers `features/*.feature`, so `tests/paper/fixtures/` can't leak
  into it.
- Fixture/tooling conventions: pytest fixtures in `tests/test_files/` via
  `tests/unitutil/file.py` helpers; oxml tests build elements from cxml strings
  (`tests/unitutil/cxml.py`, a local pyparsing DSL: `element("w:p/(w:r,w:r)")`); behave fixtures
  in `features/steps/test_files/`. CI (`.github/workflows/test.yml`) runs `pytest -q` and
  `behave -q` on CPython **3.9–3.13** — all new code must stay 3.9-compatible (no `match`, no
  `dataclass(kw_only=...)`).
- Reference material: `reference/office-transfer/` is present but **untracked** (ignored via
  the uncommitted `.gitignore` edit adding `/reference/`); it contains the helper specs mapped
  in the plan's mining table plus `skill/` user stories. Separately, tracked `ref/` holds the
  ISO/IEC 29500 spec PDFs and the full RELAX NG (`ref/rnc/`) + XSD (`ref/xsd/`) schema sets —
  the natural oracle source for fragment-scoped validation of emitted XML.

## Flags for PR-0 (human decision required or trap to encode in tests)

- **F1 — `docx.package` name collision.** CONVENTIONS §7 pins the kernel as "a new submodule
  named `package`", but `src/docx/package.py` already exists upstream (WML `Package(OpcPackage)`
  + `ImageParts`). Options: (a) implement the kernel in a new module and re-export
  `xml_equivalent`/`diff_package`/`patch_save` from `docx/package.py` via a ~3-line additive
  import block (keeps the pinned public path `docx.package.patch_save`, minimal upstream-file
  diff, shadows nothing); (b) convert `package.py` to a `package/` subpackage re-exporting
  `Package`/`ImageParts` (import-compatible but restructures an upstream file); (c) rename the
  kernel module (deviates from §7). Recommendation: (a). Decision belongs to PR-0.
- **F2 — zip writes are not currently deterministic.** `_ZipPkgWriter.write` uses
  `ZipFile.writestr(name, blob)` with no `ZipInfo`, so entries get **wall-clock timestamps**;
  entry order is iteration order (content-types, package rels, parts depth-first). Upstream
  `save()` stays untouched (§1.1); `patch_save` must own its zip writing with fixed timestamps
  and pinned entry order to satisfy §7's determinism requirement.
- **F3 — parse-time whitespace behavior.** The shared oxml parser sets `remove_blank_text=True`
  (`oxml/parser.py:19`): whitespace-only text nodes may be dropped at parse unless protected
  (e.g. `xml:space="preserve"` on `w:t`). The §7 kernel's `xml_equivalent` and the pinned
  trailing-space trap test must be designed against *parse → serialize* reality, not raw bytes;
  Phase 2 needs an explicit fixture pair proving a preserved trailing space in `w:t` survives
  round-trip and compares as **not** equivalent.
- **F4 — stale `uv.lock`.** The committed lockfile still declares the root package as
  `python-docx` (uv.lock:725, `revision = 1`), predating the `paper-docx` rename; any syncing
  `uv` command rewrites it wholesale. Until it's deliberately regenerated in a sanctioned PR,
  dev commands should use `uv run --no-sync` (the committed `.venv` provisioning is complete
  and both suites pass in it).
