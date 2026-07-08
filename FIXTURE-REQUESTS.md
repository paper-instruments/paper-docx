# Fixture requests — human authoring needed

Agents cannot run desktop Word or Google Docs. The corpus currently bootstraps
from `generated/` and `libreoffice/` buckets (see
`tests/paper/fixtures/README.md`); the fixtures below need a human with the
real applications. For each: author the file, fill in the sidecar (schema in
CONVENTIONS §4 — put your name in `verified_by`), drop both under the right
bucket, run `uv run --no-sync python tests/paper/fixtures_authoring/freeze_manifest.py`,
and PR the diff.

Conventions for all Word fixtures: US-English Word 365 (note exact version in
the sidecar), default Calibri template unless stated, author identity visible
in tracked changes/comments is fine and should be recorded in `ground_truth`.

## 0. Verify the 23 existing sidecars (no new files needed)

Every sidecar currently says
`"verified_by": "UNVERIFIED — agent-authored (Claude), machine-cross-checked; human sign-off pending"`.
A human should open each fixture (Word or LibreOffice), spot-check the
`ground_truth` values against what the document actually shows, then replace
`verified_by` with their name in `tests/paper/fixtures_authoring/write_sidecars.py`,
re-run it, and re-run `freeze_manifest.py` in the same PR.

## word/feature-isolated/ (authored in desktop Microsoft Word)

1. **tracked-ins-del.docx** — Type a short paragraph, turn on Track Changes
   (Review → Track Changes). As author A: insert a word mid-sentence. Change
   the user name (File → Options → General) to author B: delete a different
   word and insert a replacement. Do NOT accept/reject anything. Two authors
   must appear in `w:ins`/`w:del`.
2. **comments.docx** — Two paragraphs; select a phrase in each and Insert →
   Comment (different text each). One comment should contain two paragraphs
   (press Enter inside the comment).
3. **content-control.docx** — Developer tab → insert one *Rich Text* content
   control wrapping a whole paragraph and one *Plain Text* content control
   inline inside a sentence. Set Title/Tag on both (record values).
4. **textbox.docx** — Insert → Text Box (simple built-in style), type one
   sentence inside it, one body paragraph before and after. Word will emit
   `mc:AlternateContent` with a VML fallback — that duplicated
   `w:txbxContent` shape is exactly what we need on record.
5. **footnotes-endnotes.docx** — One paragraph with a real footnote
   (References → Insert Footnote), one with a real endnote.
6. **fragmented-runs.docx** — Type this sentence, letting Word autocorrect
   the quotes and dashes; then bold `$75–100/hr` and italicize
   `“full-service”` by selecting with the mouse in several passes (this
   produces the classic run fragmentation):
   `Consulting rate: $75–100/hr on a “full-service” basis — travel time billed at $37.50/hr.`
   Also add a paragraph containing a non-breaking space (Ctrl+Shift+Space),
   e.g. `Net 30 payment terms apply.` with the NBSP between "Net" and "30".
7. **numbering-list.docx** — Create a two-level multilevel list (Home →
   Multilevel List), 3 items with one demoted to level 2.
8. **table-merged-nested.docx** — 3×3 table; merge two cells horizontally and
   two vertically (Layout → Merge Cells); paste a small 2×2 table inside
   another cell.
9. **gauntlet.docx** — One document combining all of the above plus a second
   section (Layout → Breaks → Next Page) with a different header, and
   "Different First Page" enabled in section one.

## word/other/ — v0.1 honesty-recall priority requests

These four are the TOP requests: the v0.1 honesty work bootstraps them as
hand-built XML (`generated/` bucket), but LibreOffice cannot produce the real
shapes (it converts moves to plain ins/del and never emits `proofErr`), so
only desktop Word can confirm our model against reality.

9a. **tracked-moves.docx** — With Track Changes on, select a whole paragraph
    and drag it (or cut+paste it) two paragraphs down. Do not resolve. Word
    emits `w:moveFrom`/`w:moveTo` with range markers — record the moved text
    in `ground_truth`.
9b. **format-change-revisions.docx** — With Track Changes on: bold one word,
    re-center one paragraph, change one paragraph's style. Do not resolve
    (`w:rPrChange`/`w:pPrChange`).
9c. **toc-and-fields.docx** — Two heading paragraphs, References → Table of
    Contents, plus an Insert → Date field and a cross-reference (Insert →
    Cross-reference to a heading). Save WITHOUT updating fields.
9d. **word-noise.docx** — Type two paragraphs with a couple of deliberate
    typos, add one comment, save. (Captures `proofErr`, `_GoBack`, comment
    anchors, `rsid` attributes as Word actually scatters them.)

10. **spanning-revision.docx** — With Track Changes on, make one insertion
    that spans a paragraph boundary (select across two paragraphs, type
    replacement). This shape is hard to synthesize confidently and is needed
    for Phase 7/8 edge-case tests.
11. **field-codes.docx** — Insert a date field (Insert → Quick Parts → Field →
    Date) and a TOC (References → Table of Contents) over two heading
    paragraphs. Needed for field-instruction refusal tests (`w:instrText`,
    `w:fldChar`).

## google/feature-isolated/ (File → Download → Microsoft Word (.docx))

12. **tracked-suggestions.docx** — Google Doc in *Suggesting* mode: one
    suggested insertion, one suggested deletion, exported to .docx without
    resolving.
13. **comments.docx** — Google Doc with two comments, exported to .docx.
14. **basic-formatting.docx** — Headings, a bulleted and a numbered list, a
    small table; exported to .docx. (Google's exporter produces distinctive
    package shapes we must round-trip safely.)

## Sidecar ground truth to record (minimum, per file)

Visible body text (or the key sentences), counts of tracked
insertions/deletions with authors, comment authors + texts, control
titles/tags, note ids + texts, header/footer texts per section, table merge
geometry, list numIds actually referenced, and — for fragmented-runs — the
run count of the rate sentence as Word actually saved it.
