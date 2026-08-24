# Lossy selection and replacement audit

**Audit completed:** 2026-08-23

**Scope:** Paper-added DOCX search, target resolution, formatting inspection,
replacement, table editing, and comparison behavior

This is a historical design audit, not an API reference. It records where a
Paper convenience layer discarded identity, boundary, formatting, or topology
evidence and what the selection-integrity changes retained or removed. Current
contracts live in the `docx.search`, `docx.blocks`, and `docx.formatting` API
documentation.

## Evidence and provenance

The code review compared the fork with upstream `python-docx` 1.2.0 and traced
each Paper-added behavior to Git history. None of LS-01 through LS-11 came from
upstream. LS-01 through LS-05 and LS-07 through LS-11 entered in the Paper
bootstrap commit `a55be769`. LS-06 entered later in the unmerged draft PR #41
patch identified by the audit as commit `4cb35aa`.

The measured evaluation evidence is narrower than the code audit. A 2026-08-22
trace review covered 75 focused Harbor trajectories and 60 Knowledge48 DOCX
trajectories. It distinguished three events: an API was called, the lossy input
condition was present, and a harmful output or score effect was observed. Only
LS-05 reached all three. The other entries remain real code-level risks or open
design questions, but this corpus does not show that they cost points.

| ID | Evaluated condition | Observed result | Final disposition |
| --- | --- | --- | --- |
| LS-01 | Resolver called; no stale equivalent block | No adverse effect | Live block identity retained; serialized locator removed |
| LS-02 | Span-to-block path called; no cross-paragraph target | No adverse effect | One-paragraph operations refuse cross-paragraph spans |
| LS-03 | `near` called in four Harbor trajectories | One safe refusal; no final loss | Ranking remains inspection-only on `find_text()` |
| LS-04 | Normalized case-folding observed | No wrong mutation | Exact is default; normalized matching is explicit |
| LS-05 | Heterogeneous fragmented replacement | **Structure loss and score loss** | Ordinary replacement proves one safe changed region or refuses |
| LS-06 | Not present in evaluated build | Not executable | Capacity allocator and separate mode removed |
| LS-07 | No calls | Not exercised | Convenience wrapper removed |
| LS-08 | Row copy called; template cells were uniform | Lossy choice not activated | Open; outside this stack |
| LS-09 | Table lookup called; queries stayed within one cell | Lossy choice not activated | Open; outside this stack |
| LS-10 | Comparison pairing called | No observed mispair | Open; outside this stack |
| LS-11 | Tracked replacement called on simple regions | No observed formatting loss | Open; outside this stack |

## Closed issues

### LS-01 — serialized block aliases were treated as identity

**Lossy behavior.** A serialized block anchor combined a story name, block
index, and the first eight hexadecimal SHA-256 characters of normalized text.
Re-resolution treated that tuple as the block. Normalization erased case,
whitespace, and punctuation, while the 32-bit digest omitted element identity,
structure, formatting, revisions, controls, and table topology.

**Impact and failure shape.** If an edit moved a normalization-equivalent
paragraph into the recorded index, an old anchor could authorize a mutation of
the new paragraph. For example, an anchor captured for `Payment Terms` could
later resolve to a different `payment  terms` paragraph at the same index and
insert or delete content in the wrong clause while reporting success.

**Evidence.** Fresh anchors were used successfully in several eval workflows,
but no trace mutated the document between capture and reuse in a way that put
an equivalent block at the old index. There was no measured adverse effect.

**Disposition.** A live `Block` now retains its owning document and exact OOXML
element and validates attachment, story, containing structure, and ownership.
It survives index shifts but refuses detached, replaced, reparented, or foreign
elements. `BlockLocator` was removed rather than replaced with another
persistence heuristic. Legacy serialized `Anchor` data remains inert evidence
and cannot authorize mutation; callers reacquire a live target after reload.

### LS-02 — cross-paragraph spans collapsed to the first paragraph

**Lossy behavior.** Search intentionally permits a span to cross a paragraph
boundary, but the shared block resolver converted such a span to the paragraph
of its first selected atom. It discarded the other endpoint without requiring
the caller to choose one.

**Impact and failure shape.** A caller selecting `end of clause A\nstart of
clause B` and asking to insert after the selection could insert after clause A,
even though the supplied range included clause B. Field, composition, and block
operations could therefore act at an invented endpoint.

**Evidence.** Span-to-block resolution ran in the evals, but all reported spans
were confined to one paragraph. The risky condition and adverse effect were not
observed.

**Disposition.** Operations that require one paragraph revalidate both the
span summary and its concrete selected atoms and raise a boundary refusal when
they cross paragraphs. Cross-paragraph inspection remains valid. Ordinary
replacement may narrow unchanged exact affixes only when its actual changed
interval is proved to remain within one paragraph.

### LS-03 — contextual distance looked like unique authority

**Lossy behavior.** `near` ranked targets by absolute character distance to
context. Missing context gave all candidates the same non-result, and tied
distances fell back to document order. Combining the ranking with a positional
selector could make that fallback look like a context-proved identity.

**Impact and failure shape.** If `Renewal` is absent or exactly between two
`Payment terms` occurrences, choosing the first ranked result can mutate the
wrong clause even though the caller supplied context specifically to
disambiguate it. Character distance is useful evidence, but it is not semantic
document identity.

**Evidence.** Four Harbor trajectories used `near`. Three reached the intended
target and one multi-match exploration refused safely before recovery. The
corpus contained no missing-or-tied context that caused a wrong final edit or
score loss.

**Disposition.** `find_text(..., near=...)` remains an inspection tool: it
orders and returns every candidate, preserves stable document order for ties
or missing context, and makes no unique-authority claim. `find_one()` has no
`near` keyword and retains ordinary zero/one/many refusal semantics. A caller
that wants one result must supply a uniquely identifying target or explicitly
choose from inspected live spans.

### LS-04 — normalized discovery was the mutation default

**Lossy behavior.** The original search path always case-folded text, mapped
smart quotes and dashes, collapsed whitespace, mapped exotic spaces, and
removed soft hyphens. The resulting span was mutation-capable even when the
document text was only normalization-equivalent to the requested string.

**Impact and failure shape.** `Company` can select `company`; an en dash can be
treated like a hyphen; and a non-breaking space can be treated like a normal
space. Those distinctions may identify names, terms, typography, or protected
labels. A successful mutation can therefore target text the caller never
literally named.

**Evidence.** One `fm02` exploration proved case-folding was active by finding
lowercase occurrences from an uppercase query. The actual mutations used more
specific targets, and a typography-only echo was preserved. No harmful
overmatch was observed.

**Disposition.** Exact codepoint matching is the default for search and
replacement. Normalized matching remains available only through an explicit
policy. Both policies use the same live span machinery and preserve the actual
selected document text as evidence.

### LS-05 — ordinary replacement adopted the first selected run

**Lossy behavior.** Ordinary untracked replacement wrote all new text into the
first selected text node and emptied later selected nodes. The first run's
formatting and semantic scope silently became the formatting policy for the
whole replacement.

**Impact and failure shape.** Replacing text spanning regular, bold, italic, or
underlined runs could flatten the entire replacement into the first style.
Crossing a hyperlink, control, revision, or marker boundary could move text to
the wrong scope or hollow a bookmark while leaving an openable but incorrect
document.

**Evidence.** This was the only issue with direct adverse eval evidence. The
fragmented `fm02` target contained bold, italic, and underlined regions plus
bookmark/proofing boundaries. Public replacement consolidated text into the
first node and produced structure-check scores of `.962963`, `.925926`, and
`.962963` in the skill-05 attempts. A manual topology-aware edit scored
`1.000`. The trace and grader both identified the lost structure.

**Disposition.** PR #48 first introduced one ordinary planner that could leave
exact common prefix and suffix text in place and change only a residual interval,
fixing the observed formatting loss. A follow-up audit found that choosing one
greedy affix split could still manufacture certainty for repeated affixes or
insertion boundaries. The planner now considers every maximal exact alignment
and proceeds only when they agree on both one changed interval and one writable
formatting/inline-ancestry destination. Otherwise it refuses with guidance to
re-find the intended substring. It preserves safe markers and required
`xml:space`/placeholder behavior and refuses mixed, marker-crossing, or unresolved
intent. Direct, batch, and cell-update paths share the planner. Every successful
non-no-op direct replacement consumes its supplied span so stale live state
cannot be reconstructed approximately; callers re-find before another operation.
No-op and atomically refused or rolled-back direct operations leave the supplied
span reusable. The original greedy planner was part of the selection-integrity
repair, not the Paper bootstrap behavior that caused LS-05, and the follow-up is
a code-level correctness correction rather than a new eval-attributed loss.

### LS-06 — old character counts were treated as formatting intent

**Lossy behavior.** Draft PR #41 added `preserve_structure` and divided new
text among existing text nodes according to their former character capacity.
Node lengths often reflect arbitrary editing history, not boundaries between
meaningful formatting regions.

**Impact and failure shape.** Nodes that previously held 6, 4, and 8
characters could split a replacement at those same counts, including through a
new word. The call could report preserved structure while assigning the word's
halves to unrelated run formats.

**Evidence.** The evaluated build predated the draft allocator and no trace
called the mode, so LS-06 cannot explain an eval outcome.

**Disposition.** The public `preserve_structure` option, associated result
field, option matrix, and inherited capacity allocator were deleted. Safe
topology-aware text assignment belongs to the one ordinary planner described
under LS-05; an edit that planner cannot prove safe refuses. There is no
replacement mode for exact topology and no substitute allocator.

### LS-07 — formatting inspection sampled an unrelated first run

**Lossy behavior.** `surrounding_format(document, target)` resolved a target to
a paragraph and sampled that paragraph's first run, even when the target was a
live span over text later in the paragraph.

**Impact and failure shape.** In a paragraph beginning bold and ending regular,
asking about the regular clause could report bold. An agent could then use that
incorrect evidence to format an insertion or replacement.

**Evidence.** The skill material mentioned the wrapper, but none of the 135
reviewed trajectories called it. No score can be attributed to LS-07.

**Disposition.** The redundant wrapper was removed. Callers explicitly locate
text and pass the live span to `format_of()`, which resolves every touched run
and reports mixed values and provenance; callers that already hold a run or
paragraph pass that object directly.

## Open issues outside this stack

These behaviors were also introduced by Paper bootstrap commit `a55be769`.
They were not changed by the selection-integrity stack, and this audit does not
present the absence of an eval loss as proof that they are safe.

### LS-08 — copied table cells choose one representative run format

**Lossy behavior and impact.** Row-copy population finds the first run with
explicit run properties, clears the cell, and applies that one property set to
the replacement. A template cell whose first styled run is red and whose later
run is bold can become entirely red, losing the later region and any semantic
relationship between text and style.

**Evidence and status.** Row copying ran in three `fm13` trajectories, but each
template cell had one ordinary unformatted run, so no competing formats were
discarded. The observed score loss came from row order, not this heuristic.
The issue remains open and is outside this stack.

### LS-09 — table search can synthesize text across cell boundaries

**Lossy behavior and impact.** `find_table()` concatenates cell text with
separators and normalizes the result before substring search. Because
normalization collapses whitespace, a query can be formed from the end of one
cell and the start of another even though no cell contains it. In a document
with repeated tables, `Account` in one cell and `Balance` in the next can make
`Account Balance` appear to be a real within-cell marker and select the wrong
table.

**Evidence and status.** Public table lookup ran in four `fm13` trajectories,
but every query was a unique marker wholly inside one cell. Cross-cell
synthesis was not activated. The issue remains open and is outside this stack.

### LS-10 — comparison pairs blocks above one global text threshold

**Lossy behavior and impact.** Word-level comparison uses order-preserving
text-similarity pairing and a global `0.5` threshold. Similar boilerplate or
repeated clauses can cross the threshold for the wrong counterpart, producing
a plausible but misleading redline instead of an explicit deletion and
insertion.

**Evidence and status.** All five valid `fm21` trajectories exercised block
pairing and produced correct accept/reject projections. Their `.250` scores
came from an unstated exact author/date grader filter, not an observed mispair.
That clean fixture does not validate the threshold for repeated or similarly
worded paragraphs. The issue remains open and is outside this stack.

### LS-11 — tracked replacement infers revision formatting from one run

**Lossy behavior and impact.** Tracked replacement narrows a common textual
prefix and suffix, then uses run properties from the first changed region for
inserted redline text. A replacement spanning differently formatted changed
runs can therefore attribute all inserted text to one run's formatting and
present a semantically misleading revision even when accept/reject text is
correct.

**Evidence and status.** The path ran in `fm06`, `fm02` experiments, and
`fm21`, but those changed regions did not force a choice among competing run
formats. Accept/reject algebra was correct and no grader reported this loss.
The issue remains open and is outside this stack.

## Resulting boundary

The closed changes remove authority that was based on normalization,
position, missing context, a representative run, or old character capacity.
They retain only evidence the package can validate directly: exact or
explicitly normalized text policy, complete candidate inspection, live
owner-bound objects, concrete paragraph boundaries, span-local formatting,
and one proved uniform replacement region. When those facts do not determine a
safe mutation, the public contract is refusal rather than a plausible guess.
