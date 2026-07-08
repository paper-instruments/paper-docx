# paper-docx

**A drop-in fork of [`python-docx`](https://github.com/python-openxml/python-docx)
that adds the editing surface real document work needs: complete visibility
into a document, find-and-replace that survives Word's run fragmentation,
real tracked changes, and saves that touch only what you changed — all behind
typed, atomic refusals instead of silent corruption.**

```python
import docx                     # the import name never changes
doc = docx.Document("contract.docx")
```

Everything `python-docx` did, this package still does, unchanged — upstream's
own pytest and behave suites run green on every commit. What follows is what
it does *now* that stock `python-docx` cannot.

## Why this fork exists

`python-docx` has an excellent core: a lossless package layer that
round-trips content it doesn't understand, and a disciplined XML layer with a
decade of absorbed edge cases. That's why we forked it rather than rebuilt.

But its *editing* surface stalled years ago, and the gaps are exactly where
professional, agent-driven document work lives:

- **It can't see whole regions of the document.** Text inside tracked
  insertions, content controls, text boxes, footnotes, and endnotes is
  invisible to `.paragraphs`/`.runs` — the library literally cannot read
  parts of what Word displays.
- **It can't edit real-world text.** Word fragments most text across
  multiple runs (`$75`&#8203;`–`&#8203;`100/hr` is typically three runs), so there is no
  way to find or replace it without destroying formatting.
- **It has no tracked changes.** No way to propose a redline, and no way to
  accept or reject one.
- **Its saves churn the whole package**, so reviewing "what did this edit
  actually change?" is impossible at the file level.

Production systems have historically papered over these gaps with raw-XML
surgery, whose dominant failure mode is **silent corruption**: files that
open fine and are quietly wrong. This fork replaces that surgery with
first-class APIs whose failure mode is a loud, typed refusal.

## What was added

### See the whole document — `docx.story`

Opt-in traversal that covers every story part (body, headers, footers,
footnotes, endnotes, comments) and every region standard traversal is blind
to, with stable anchors and a choice of *views* — the document as it stands
(`"current"`), as it was before pending revisions (`"original"`), or
everything at once (`"all"`).

```python
from docx.story import outline

o = outline(doc)
o.blind_region_counts                # {"tracked_insertions": 2, "text_boxes": 1, ...}
[b.text for b in o.blocks if b.in_text_box]      # text upstream can't see
[b.text for b in o.blocks if b.story == "word/footnotes.xml"]
```

### Find and edit anything — `docx.search`

Matching is normalized (smart quotes, dashes, exotic spaces, case) and
assembles across fragmented runs, so you can quote text the way a person —
or a model — actually quotes it. The returned `Span` maps visible text back
to the exact XML that holds it, and `replace` preserves every untouched
run's formatting byte-for-byte.

```python
from docx.search import find_one

span = find_one(doc, 'rate: $75-100/hr')          # matches “rate: $75–100/hr”
span.replace("rate: $85–110/hr")                  # surgical; formatting intact
```

### Real tracked changes — `tracked=True` and `doc.revisions`

Redlines are genuine `w:ins`/`w:del` markup that Word renders natively. The
mark is minimal (common prefix/suffix trimmed), deleted text keeps each
source run's formatting so rejection restores the document exactly, and the
whole algebra is tested: *accept(tracked edit) ≡ plain edit* and
*reject(tracked edit) ≡ original*.

```python
find_one(doc, "forty-two units").replace(
    "forty-seven units", tracked=True, author="Alice Editor")

for rev in doc.revisions:                          # every story part
    print(rev.revision_type, rev.author, repr(rev.text))
doc.revisions.reject_all(author="Bob Reviewer")    # or accept_all()
```

### Clause-level edits — `docx.blocks`

Insert, delete, or replace whole paragraphs relative to a content anchor —
plain or as a tracked redline that preserves paragraph properties and marks
paragraph breaks correctly, so accepting/rejecting behaves exactly like Word.

```python
from docx.blocks import insert_section_after, tracked_replace_paragraphs

insert_section_after(doc, "Scope of Services",
                     heading="Confidentiality", paragraphs=["Each party shall…"])
tracked_replace_paragraphs(doc, "The term of this agreement",
                           ["The term is twenty-four (24) months."],
                           author="Alice Editor")
```

### Saves you can review — `docx.package`

A compare-based narrow save: parts you didn't semantically change keep their
**original bytes**, so a file-level diff of your edit shows your edit and
nothing else. A no-op round trip is byte-identical. Plus a semantic package
diff to prove it.

```python
from docx.package import patch_save, diff_package

patch_save("contract.docx", doc, "contract-redlined.docx")
diff_package("contract.docx", "contract-redlined.docx").semantic_changed_parts()
# ('word/document.xml',)
```

### Guarded structure ops — `docx.tableops`, `docx.numbering`

Table cell/row edits and list-numbering application that refuse loudly on
structures they can't handle safely (merged cells, nested tables, undefined
numbering) instead of guessing.

### Refusals you can catch — `docx.errors`

Every mutating API validates fully **before** touching anything. A refusal
means the document — in memory and on disk — is exactly as it was, and the
exception type tells you why:

```python
from docx.errors import PaperRefusal, AmbiguousTargetError

try:
    find_one(doc, "the")                # matches everywhere
except AmbiguousTargetError:
    ...                                 # disambiguate with nth=, near=, story=
except PaperRefusal:
    ...                                 # any safe refusal, distinct from bugs
```

A refused edit is a success mode. A quietly wrong file is the failure this
package exists to eliminate.

## Design principles

- **Strict superset.** v0 changes zero existing behavior; new capability is
  new, explicitly named API. Upstream's full test suites gate every change.
- **Refusal atomicity.** Validate fully, then mutate. Every documented
  refusal is tested to leave output bytes identical to input bytes.
- **Whitespace is content.** No comparison or rewrite path ever normalizes
  meaningful text whitespace — a trailing space in `w:t` is a real character.
- **Provenance-tested.** The frozen fixture corpus (`tests/paper/fixtures/`)
  spans generated and LibreOffice-authored files, with hand-verified ground
  truth, hash-frozen manifests, and a LibreOffice headless load oracle;
  desktop-Word fixtures are tracked in `FIXTURE-REQUESTS.md`.

## Naming

Four names to keep distinct — the mismatch is intentional (same pattern as
Pillow/PIL):

- GitHub repository / PyPI distribution: **`paper-docx`**
- Python import package: **`docx`** — frozen forever; millions of existing
  snippets and model priors say `from docx import Document`, and drop-in
  compatibility is the entire thesis of this fork
- Fork sentinel: `docx.__paper_version__`

Never rename `src/docx`, and never write `import paper_docx`.

## Installation

This repository is private and publication to PyPI is intentionally gated.
For now, install from Git:

```bash
pip install "paper-docx @ git+https://github.com/The-LLM-Data-Company/paper-docx.git@main"
```

Verify the fork sentinel:

```bash
python -c "import docx; print(docx.__paper_version__)"
```

## Going deeper

- [`API-PROPOSAL.md`](API-PROPOSAL.md) — the full approved v0 surface:
  signatures, return types, refusal conditions (mechanically enforced by
  `tests/paper/test_api_surface.py`).
- [`ARCHITECTURE-NOTES.md`](ARCHITECTURE-NOTES.md) — how the codebase is
  layered and where each organ lives.
- [`PAPER.md`](PAPER.md) — fork lineage, baseline test results, sanctioned
  deviations (none), upstream merge policy.
- Upstream [python-docx documentation](https://python-docx.readthedocs.io/)
  — everything inherited works as documented there.

`python-docx` is by Steve Canny and contributors (MIT); this fork exists
because that foundation was worth building on.
