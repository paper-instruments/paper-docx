# Practical hardening proposal for paper-docx

## Purpose

This proposal captures the hardening work that would materially improve
`paper-docx` for professional document workflows. It is intentionally narrower
than the full research audit. The priority is preventing lost work, silent
document damage, incorrect edits, and unreliable review artifacts in contracts,
transaction documents, reports, and other long-lived Word files.

The `0.1.1` candidate remains the baseline. The work below should be delivered
as small, independently reviewed changes after that release rather than as one
large hardening patch.

## Priority findings

### 1. Guarantee atomic saves and compound edits

Some operations update multiple XML parts, relationships, or package-level
records. A late refusal or serialization failure must not leave the in-memory
document partly changed or overwrite a destination with an incomplete package.

This matters in practice because composition, review finalization, scrubbing,
and comment operations can touch several related parts. A lawyer or banker must
be able to retry or abandon a failed operation without reopening the original
file and without wondering which changes were applied.

Recommended contract:

- Fully validate an operation before its first mutation.
- Roll back every affected in-memory part when a typed refusal or unexpected
  error occurs after mutation begins.
- Stage and validate the complete OPC package before replacing a file or
  writing to a non-restorable output stream.
- Add forced-late-failure tests for each compound public operation.

### 2. Prevent stale or ambiguous edits from targeting the wrong text

Search results, blocks, spans, comments, and other editing handles can become
stale after an earlier mutation. Reusing an old handle must refuse rather than
apply an edit to content that happens to occupy the same approximate location.

This is a high-impact risk in automated contract and presentation-book editing:
an apparently successful replacement on the wrong clause, footnote, or table
cell is worse than an error.

Recommended contract:

- Validate content anchors against the current story and block before editing.
- Treat source location and normalized text identity as part of freshness, not
  just the visible matched string.
- Reject reversed, detached, cross-story, or otherwise ambiguous ranges.
- Keep replacement batches atomic when one result becomes stale during the
  operation.

### 3. Preserve fields, revisions, comments, bookmarks, and content controls

Visible text in a real Word document is often surrounded by structural markup.
A text edit that ignores those boundaries can orphan a field instruction,
invalidate a tracked-change range, detach a comment, hollow out a bookmark, or
desynchronize a content control from its stored value.

These structures are common in professional templates and negotiated documents.
They must be treated as document data, not incidental XML.

Recommended contract:

- Refuse edits that partially consume a complex field or cross incompatible
  revision scopes.
- Validate the complete selected range, including intervening runs, before
  adding or changing a comment.
- Reject stale comment proxies after comments are removed or rebuilt.
- Preserve bookmark pairing and revision scope, and refuse changes that leave a
  referenced bookmark empty.
- Permit generic text mutation only in content controls whose surface is
  actually text-editable; route typed controls through their dedicated APIs.
- Apply Restrict Editing checks consistently to inherited and fork-added
  mutators.

### 4. Make scrub and review finalization complete and trustworthy

`scrub`, accept/reject, and finalization APIs are used immediately before a
document is sent outside the organization. Their result must not retain hidden
review data or remove visible/structural content accidentally.

Recommended contract:

- Traverse every reachable story and review-related part, not only the main
  document body.
- Remove comment records, anchors, extended-comment data, and relationships as
  one atomic operation.
- Refuse hidden-text deletion when it would invalidate a bookmark, field,
  content control, or consumer-dependent compatibility branch.
- After finalization, verify that no targeted revision or review residue remains
  anywhere in the live package graph.
- Test the saved artifact with reopen checks and LibreOffice load smoke, not
  only the in-memory XML tree.

### 5. Preserve package relationships during composition

Cross-document composition must copy a complete, internally consistent graph of
styles, numbering definitions, media, hyperlinks, notes, and other referenced
parts. A composed document that opens but silently uses the wrong numbering,
style, image, or link is not a successful result.

Recommended contract:

- Validate source ownership and every transitive internal relationship before
  mutation.
- Reconcile style and numbering identifiers deterministically and return the
  resulting mapping to the caller.
- Refuse unsupported or ambiguous relationship graphs rather than dropping or
  sharing targets accidentally.
- Verify that source and destination remain independently editable after the
  import.

### 6. Address tables and lists by Word's structural model

Professional documents rely heavily on merged table cells, repeated headers,
numbered clauses, and nested lists. Physical XML position is not always the
visible grid column or list identity. Editing by the wrong coordinate can
change the wrong cell or reconnect a paragraph to the wrong numbering stream.

Recommended contract:

- Resolve table locations through `gridBefore`, `gridSpan`, vertical merges,
  and the visible layout grid.
- Guard the whole target cell for fields and pending revisions before replacing
  its contents.
- Validate numbering references and abstract-number definitions before list
  edits or composition.
- Return explicit remapping information whenever identifiers change.

### 7. Reject ambiguous package graphs before editing

Word files from document-management systems, conversion tools, and long editing
histories frequently contain optional parts and producer-specific package
layouts. The library should accept valid variation, but it must not guess when
duplicate names, conflicting content types, or ambiguous relationships make the
meaning of a part uncertain.

Recommended contract:

- Resolve internal part names case-insensitively, as OPC requires, while
  detecting case-colliding members.
- Validate relationship cardinality, role, target mode, target existence, and
  content type before exposing a part to a mutator.
- Preserve unknown but reachable parts byte-for-byte when they are outside the
  requested edit.
- Refuse ambiguous graphs before mutation and leave the source and destination
  unchanged.

## Delivery approach

Each priority should be implemented as a separate PR with a narrow public
contract and regression fixtures that demonstrate the user-visible failure.
Tests should include successful round trips, typed refusals, forced late
failures, package reopen checks, and LibreOffice load smoke where the saved file
is affected.

The work should not expand the public API unless a new return value is necessary
to report provenance, remapping, or refusal details. Existing import names,
signatures, exception messages, and drop-in behavior remain fixed unless a
separate compatibility decision explicitly changes them.

## Deliberately excluded

This proposal excludes hardening that only concerns contrived XML shapes,
resource limits with no plausible Office producer, cosmetic API consistency,
internal refactoring, speculative schema coverage, and checks that duplicate an
existing guarantee without preventing a practical failure. Those findings can
remain in the research branch and should not block `0.1.1`.
