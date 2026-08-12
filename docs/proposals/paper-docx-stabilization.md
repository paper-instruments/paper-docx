# paper-docx stabilization implementation proposal

Status: Proposed

Baseline: `main` at `a55be76` / `paper-docx` 0.1.2

Documentation candidate: `origin/gavin/docs-readme` at `a6ffb9f`

Last updated: 2026-08-11

## Summary

Keep `paper-docx` as a hard fork, but narrow the stability promise and complete
the work needed to make that promise true.

The fork is justified by behavior that must sit inside or immediately beside
python-docx internals:

- bounded OPC and ZIP loading;
- staged, validated, atomic package writes;
- revision-aware text perception;
- deterministic story traversal and cross-run search;
- tracked edits and revision resolution;
- transactional multi-part mutations;
- integrated comments, notes, and protection checks; and
- package-level inspection and minimal-save workflows.

The current package exposes a much larger surface than the evidence supports as
stable. `compare`, cross-document composition, and delivery-oriented scrubbing
are valuable, but they depend on complex Word semantics and have not been
validated against a representative Microsoft Word corpus. `patch_save` is a
core capability in intent, but it currently has weaker destination semantics
than `Document.save()`.

This proposal therefore makes four decisions:

1. Preserve the hard fork and its agent-oriented core.
2. Divide the public surface into stable, provisional, and experimental tiers.
3. Fix correctness and packaging failures before expanding the API.
4. Replace universal safety claims with exact, testable contracts.

## Evidence behind the proposal

The fork is a material implementation, not a documentation wrapper:

- 42 `src/docx` files changed or added relative to `paper-base`;
- 15,064 source lines added and 64 removed;
- 18,134 lines added under `tests/paper`;
- 2,597 pytest tests pass when run with the warning filters used by CI; and
- 650 inherited Behave scenarios pass.

The green inherited suites are strong evidence that ordinary python-docx 1.2.0
behavior has largely been preserved. They do not establish compatibility with
all Word-produced documents or prove the package's broad safety claims.

The audit also confirmed the following gaps:

- `patch_save()` replaces a destination symlink instead of updating its target,
  while `Document.save()` preserves the symlink and updates the target;
- `patch_save()` does not fsync the staged file or destination directory;
- installing `python-docx` after `paper-docx` overwrites the import guard but
  leaves Paper-only modules behind, producing a mixed package;
- the Microsoft Word and Google Docs fixture buckets are empty;
- the LibreOffice gate proves openability, not visual or Word-semantic fidelity;
- `tests/paper/test_composition.py` contains an always-true assertion;
- Ruff and Pyright are advertised as checks but are not CI gates; and
- the draft README contains incorrect or overly broad comparisons and safety
  guarantees.

One production task successfully used Paper's search, tracked deletion,
tracked replacement, comment, revision-view, and minimal-save APIs. The output
contained native revisions and comments and preserved unrelated package
content. This proves that the mechanisms can support a real workflow. It does
not prove package-only superiority because the package, skill, tools, and agent
workflow were not isolated from one another.

## Goals

- Make every stable-tier contract precise enough to test.
- Give `patch_save()` the same destination-safety semantics as `Document.save()`.
- Detect mixed Paper/upstream installations wherever Paper code executes.
- Validate Word-facing behavior against documents created and processed by
  Microsoft Word.
- Make the repository's advertised quality checks real and reproducible.
- Correct the README and RST documentation before publishing broader claims.
- Establish an explicit upstream-maintenance strategy for the fork.
- Preserve the useful same-namespace agent experience only if its operational
  risks are accepted and controlled.

## Non-goals

- Reimplement every Microsoft Word feature.
- Promise that a document which opens is visually or semantically correct.
- Treat protection metadata as an access-control or DRM boundary.
- Make `compare()` equivalent to Microsoft Word's general Compare command.
- Support arbitrary composition graphs by silently dropping unsupported data.
- Rewrite the published bootstrap commit or reconstruct historical commits.
- Use the production eval as causal proof that Paper beats stock python-docx.

## Architectural decisions

### AD-1: Retain the hard fork

The project should remain a hard fork rather than become only a wrapper around
stock python-docx.

Revision-aware stock accessors, load preflight, save replacement, custom part
registration, and inherited comment/protection behavior require changes inside
the `docx` implementation. A wrapper could expose separate helpers, but stock
accessors would remain revision-blind and the inherited save path would remain
outside the safety contract.

The hard fork does not imply that every workflow belongs in the core package.
Higher-level features may be provisional, experimental, or eventually moved to
a companion distribution.

### AD-2: Record a separate decision for the shared `docx` import name

Before the next feature release, add an architecture decision record that
compares:

1. retaining the `docx` import namespace;
2. moving to a distinct `paper_docx` namespace; and
3. providing a compatibility shim during a namespace migration.

The proposed near-term decision is to retain `docx`, because compatibility with
existing code and model priors is a real product benefit. This is conditional
on accepting these deployment invariants:

- `paper-docx` must be the sole installed owner of the `docx` tree;
- deployments must run `paper-docx-doctor` before document work;
- controlled environments must block `python-docx` through constraints or
  dependency replacement;
- applications with unavoidable transitive `python-docx` dependencies must use
  a separate environment; and
- documentation must state that `import docx` cannot detect every overwrite
  direction.

The ADR must name an owner, reevaluation date, migration cost, and trigger for
reconsidering the namespace. A trigger should include any production mixed-
install incident or a material upstream release that becomes difficult to
reconcile.

### AD-3: Adopt stability tiers

#### Stable after the required fixes in this proposal

- OPC and ZIP preflight;
- hardened `Document.save()`;
- transaction and rollback infrastructure;
- story traversal and blind-spot reporting;
- normalized cross-run search and anchored spans;
- tracked text and block edits;
- revision enumeration and supported accept/reject operations;
- native comment anchoring and Paper comment operations; and
- protection checks, explicitly limited to Paper mutators.

#### Provisional

- `patch_save`, until destination parity is complete;
- content controls;
- table operations;
- numbering;
- bookmarks;
- fields; and
- effective-formatting inspection.

These APIs solve relevant problems and may remain public. Their documentation
must identify supported structures and refusal boundaries. They move to stable
only after Word-authored fixtures cover their principal success and refusal
paths.

#### Experimental

- `compare` as a general user-facing workflow;
- cross-document composition;
- `scrub()` and delivery-oriented finalization; and
- cross-package human-readable diff workflows.

Experimental APIs must be labeled in the README, RST pages, docstrings, and
release notes. They must not carry the stable tier's universal guarantees.
Nothing in this proposal requires immediate removal. If composition has no
demonstrated product use after two release cycles, evaluate moving it to a
companion `paper-docx-workflows` distribution.

## Required implementation changes

### P0-1: Unify path-based write semantics

Replace the separate path-writing implementation in `docx._paperpkg` with one
shared internal primitive used by both `Document.save()` and `patch_save()`.

The primitive must:

- require an existing parent directory, or explicitly document and test any
  directory-creation side effect;
- resolve a destination symlink once;
- record the symlink device, inode, and resolved target;
- stage the complete output beside the resolved destination;
- preserve the existing destination's permission bits;
- validate the complete staged OPC package;
- fsync the staged file before replacement;
- revalidate that the symlink and resolved target did not change;
- atomically replace the resolved destination;
- fsync the destination directory where supported; and
- remove the temporary file on every failure path.

`patch_save(link, document, link)` must preserve `link` as a symlink and update
the same target that `Document.save(link)` would update.

Tests must cover:

- a new regular destination;
- an existing regular destination;
- an in-place save;
- a relative symlink;
- an absolute symlink;
- a dangling symlink;
- a symlink replaced between staging and commit;
- a symlink whose target changes between staging and commit;
- destination permission preservation;
- validation failure before replacement;
- write, fsync, and replace failures; and
- cleanup of staged files after each failure.

No public documentation may call `patch_save` durable or equivalent to
`Document.save()` until these tests pass.

Define its package-preservation contract separately from its destination-write
contract:

- a true no-op must copy the complete original archive verbatim;
- for a changed output, "unchanged part bytes" means the uncompressed payload
  bytes for that OPC part, not the original compressed ZIP member;
- central-directory order, compression stream, timestamps, permission bits,
  comments, and other ZIP container metadata must either be preserved or named
  as normalized output in the result and documentation;
- deterministic normalization must never be described as complete container
  preservation; and
- `PatchSaveResult` must distinguish semantic part preservation from archive-
  container normalization.

### P0-2: Define the exact refusal and rollback contract

Replace the current universal language in `docx.errors`, the README, and the RST
pages with this scoped contract:

> A documented Paper mutator either completes and returns its declared outcome,
> or raises `PaperRefusal` without changing the live document package or an
> existing path destination. Programmer errors and operating-system failures
> are not `PaperRefusal` conditions. Stream destinations have the explicitly
> documented rollback guarantees of their stream type.

Implementation requirements:

- every stable compound Paper mutator must validate before its first mutation
  or use `rollback_on_error` for all later phases;
- rollback must restore XML trees, parts, relationships, content-type state,
  binary blobs, and package-owned registries reachable by the operation;
- rollback must cover unexpected exceptions after mutation begins, not only
  `PaperRefusal`;
- path writes must not replace an existing destination until serialization and
  package validation complete;
- file-like outputs must distinguish readable/seekable, write-only seekable,
  and write-only non-seekable behavior; and
- inherited upstream mutators must be described as outside the Paper refusal
  contract unless they are explicitly wrapped and tested.

Add a contract-test registry mapping every stable public mutator to:

- at least one success test;
- at least one pre-mutation refusal test;
- at least one forced late-failure rollback test when it can mutate more than
  one structure; and
- its documented protection behavior.

### P0-3: Harden mixed-distribution detection

The package cannot guarantee that `import docx` refuses after upstream has
overwritten Paper's `docx/__init__.py`. The implementation and documentation
must stop claiming otherwise.

Required changes:

- keep `assert_distribution_identity()` at the Paper package entry point;
- invoke the identity check from a small shared Paper-only dependency imported
  by every Paper-only public module, so `import docx.search` or
  `import docx.revision` fails even when upstream overwrote `__init__.py`;
- keep `paper-docx-doctor` as the authoritative environment check;
- make the doctor verify distribution metadata, installed paths, Paper wheel
  `RECORD` hashes, the fork sentinel, and ownership of every Paper public module;
- add a dummy-package CI fixture whose dependency on `python-docx` simulates a
  transitive clobber;
- retain both install-order tests and the unsafe-uninstall test; and
- print a remediation that reconstructs a clean environment rather than
  promising an in-place repair.

The docs must recommend a resolver constraint such as `python-docx<0` only for
controlled deployments and explain that it intentionally makes any direct or
transitive upstream requirement unsatisfiable.

### P0-4: Correct the documentation contract before merging Gavin's branch

The documentation branch contains no runtime implementation changes. Do not
merge its broadened claims until the wording changes below are applied.

#### Required replacements

| Current claim | Required meaning |
| --- | --- |
| Every fork addition succeeds exactly or refuses with memory and disk byte-for-byte unchanged | Scope the guarantee to documented Paper mutators, existing path destinations, and documented stream behavior. |
| Stock traversal is blind to headers, footers, and comments | Stock exposes headers and footers per section and exposes comments. Paper adds one traversal over supported story parts, revision views, notes, text boxes, and a blind-spot census. |
| Stock comments are "add only" | Stock can add, read, and edit comment content and can anchor comments on run boundaries. Paper adds arbitrary-span convenience, replies, resolution, traversal, and guarded mutation. |
| Stock table edits require raw XML indexes | Stock has cells, rows, columns, row insertion, and merging. Paper adds explicit visual-grid interpretation and conservative refusals for ambiguous structures. |
| `import docx` always raises when both distributions exist | The guard raises only while Paper's entry-point files remain installed. The doctor detects both overwrite directions. |
| Atomic, deterministic save on every save | State `Document.save()` guarantees separately from `patch_save`; document stream exceptions and deterministic-output scope must be explicit. |
| `compare` proves literal reproduction | State that it checks Paper accept/reject self-consistency using canonical story comparison and semantic package equivalence. It does not prove Microsoft Word behavior or literal ZIP equality. |
| Every README example was executed and verified | Remove the claim or run the examples in CI as a maintained executable corpus. |
| `tests/paper` passes 1,020 tests | Report the count generated by the actual command, or omit volatile counts. The direct observed count is 986. |

Additional documentation requirements:

- replace "strict superset" or unconditional "drop-in" wording with a precise
  compatibility statement: ordinary python-docx 1.2.0 code remains supported,
  while hostile, ambiguous, or newly bounded inputs may now refuse;
- label provisional and experimental APIs next to their first mention;
- keep the existing limitations for protection, composition, compare, blind
  regions, and non-seekable streams;
- describe deterministic ZIP output as stable only within the tested runtime
  matrix unless cross-zlib equality is proven;
- distinguish package validity, Word compatibility, visual fidelity, and
  business correctness; and
- do not cite the production eval as package-only superiority.

Change the package classifier from `Development Status :: 5 -
Production/Stable` to an appropriate beta classifier until the stable release
gates in this proposal pass. Restore the stable classifier only as an explicit
release decision backed by current Word compatibility results.

### P1-1: Build a Microsoft Word compatibility corpus

Add hash-frozen fixture buckets for desktop Microsoft Word. Google Docs exports
remain useful but are secondary to Word for the first stabilization release.

The initial Word corpus must include:

- fragmented runs with direct and style-derived formatting;
- headers and footers across multiple sections;
- footnotes and endnotes;
- basic, threaded, and resolved comments where the Word version supports them;
- insertions, deletions, moves, run-format changes, paragraph-format changes,
  and table-row revisions;
- merged and omitted table cells;
- lists with restarts and custom numbering;
- plain, rich-text, checkbox, dropdown, date, locked, and data-bound controls;
- bookmarks referenced by REF, PAGEREF, TOC, formulas, and hyperlinks;
- fields and stale cached field results;
- hidden text, RSIDs, custom properties, and review metadata;
- images, external hyperlinks, and internal relationships; and
- a small set of long-lived real-world documents with redacted content.

Each fixture needs producer version, provenance, expected structures, supported
operations, known blind regions, and a human-reviewed ground-truth sidecar.

Compatibility gates must distinguish:

1. Paper can load the source.
2. Paper can save and reopen the output.
3. LibreOffice can convert the output.
4. Microsoft Word can open without repair.
5. Word displays expected revisions, comments, fields, and controls.
6. Accepting and rejecting Paper revisions in Word yield the expected states.
7. Untargeted content remains structurally and visually equivalent.

The Word gates may run in a scheduled or release workflow if desktop Word
automation is unavailable on every pull request. Their latest passing result
must be visible before a stable release.

### P1-2: Validate `compare` independently

Keep the current accept/reject algebra as a valuable internal self-consistency
check, but test the canonicalizer as an independent trust boundary.

Add adversarial tests for:

- run splitting and coalescing;
- whitespace and `xml:space` preservation;
- namespace-prefix changes and QName-valued attributes;
- proofing markers;
- relationship ordering and duplicate relationship semantics;
- content-type defaults and overrides;
- revision IDs, authors, and dates;
- fields, controls, images, hyperlinks, and merged tables that must refuse; and
- package-part addition and removal.

For supported comparisons, automate Microsoft Word accepting and rejecting the
Paper-produced redline and compare those saved outputs with the expected source
states. Until this passes, document `compare` as experimental and its algebra as
a Paper self-consistency check.

### P1-3: Validate scrub and finalization as delivery operations

Treat `scrub()` and delivery-oriented finalization as experimental until they
are tested on Word-authored documents.

Create an explicit inventory of review and privacy-bearing structures,
including:

- core and custom document properties;
- comments, people, extended comments, replies, and resolution state;
- all supported revision forms across every story;
- RSIDs and proofing/session metadata;
- hidden text;
- custom XML and data-bound controls;
- field instructions and cached results;
- document variables and producer-specific review parts; and
- external relationships that may disclose local or remote targets.

The API must report what it removed, what it preserved, what it could not
inspect, and whether delivery certification succeeded. It must refuse to claim
cleanliness if a relevant structure is unsupported or a blind region remains.
Do not automatically delete custom XML, external relationships, or hidden text
without an explicit caller policy.

### P1-4: Bound cross-document composition

Composition remains body-only and experimental. Strengthen it around an
explicit support matrix rather than broadening it opportunistically.

Required work:

- fix the always-true imported-style assertion;
- add invariant tests for styles, numbering, media, hyperlinks, bookmarks,
  content-control IDs, and relationship ownership;
- prove that source and destination remain independently editable;
- add Word-produced fixtures for supported content;
- return every identifier and part remapping in `CompositionReport`;
- refuse unsupported notes, sections, fields, controls, revision scopes, and
  relationship graphs before mutation; and
- clarify that `section="new_page"` inserts a page break, not a Word section.

If product evidence does not justify this maintenance surface after two release
cycles, propose extracting it to `paper-docx-workflows`.

### P1-5: Fuzz and harden package preflight

`docx._zipguard` parses hostile archive metadata and belongs to the security
boundary. Add continuous fuzzing or a reproducible local fuzz target covering:

- local-header and central-directory disagreements;
- duplicate, case-colliding, Unicode-normalization-colliding, and traversal
  names;
- ZIP64 boundaries;
- data descriptors;
- encrypted and unsupported compression entries;
- forged compressed and expanded sizes;
- overlapping entries;
- truncated records;
- excessive member counts, expansion ratios, and total sizes; and
- malformed relationships and content-type declarations after extraction.

Seed the corpus with every existing corrupt fixture. Preserve a minimized
regression input for every discovered crash or incorrect acceptance. Run the
bounded regression corpus in normal CI and the full fuzzer continuously or on
a schedule.

### P1-6: Make quality gates truthful

#### Ruff

- change the target from Python 3.8 to the package minimum, Python 3.9;
- fix the source violations;
- explicitly annotate intentional public reexports;
- fix or suppress test findings only with a stated reason; and
- run `ruff check src tests` in required CI.

#### Pyright

Do not treat the present strict whole-fork result as a useful gate. Establish an
upstream baseline and select one of these approaches:

1. type-check Paper-owned modules strictly and treat inherited upstream errors
   as a frozen baseline; or
2. maintain a decreasing allowed-error baseline with ownership and expiry.

The first approach is preferred. CI must fail on new errors within the selected
boundary. If neither approach is adopted, remove Pyright from the advertised
check suite.

#### Tests

- remove the `or True` assertion and search for other vacuous assertions;
- add mutation testing for `search.py` and `revision.py` first;
- make test-count claims generated rather than hand-maintained;
- run README examples as tests if they claim observed output;
- keep the inherited pytest and Behave suites required; and
- keep build, install-order, docs, and LibreOffice jobs required.

### P1-7: Define deterministic-output scope

Determine whether "deterministic" means:

- repeated saves in one environment;
- identical bytes across the supported Python and operating-system matrix; or
- identical bytes across zlib implementations.

Add a CI job that builds the same non-noop package across the supported matrix
and compares hashes. If bytes differ, document the narrower guarantee. Do not
make cross-environment content-addressing claims without cross-environment
proof.

### P2-1: Establish upstream maintenance

Do not rewrite the bootstrap commit. Instead, add an upstream-maintenance
document containing:

- the exact upstream base commit and release;
- a machine-generated inventory of inherited files modified by Paper;
- the reason for each inherited-file divergence;
- subsystem ownership;
- expected conflict hotspots, especially comments, package loading, and saving;
- a quarterly upstream-fetch and comparison cadence;
- the process for evaluating upstream feature overlap;
- compatibility tests required after an upstream rebase; and
- a changelog section separating upstream merges from Paper behavior changes.

Future implementation commits must remain narrow enough to review and bisect.
New public behavior should not land as another undifferentiated bootstrap.

### P2-2: Run matched package evaluations

Create a small immutable control evaluation over tasks selected for revision,
comment, fragmented-text, table, control, and package-preservation requirements.

At minimum compare:

1. stock python-docx with stock instructions;
2. Paper with the same stock instructions; and
3. Paper with the Paper skill and workflow.

Hold model, prompt, source files, sandbox, rubric, sampling, and tool access
fixed. Run more than one sample per task. Measure:

- task score;
- successful completion and refusal rate;
- document openability and Word repair prompts;
- revision and comment validity;
- changed-part budgets;
- unrelated-content equivalence;
- visual differences where relevant;
- number of raw-XML fallbacks; and
- time and tool-call cost.

Use the result to prioritize APIs and decide whether experimental modules stay
in the core distribution. Do not block correctness fixes on this evaluation.

## Public API and compatibility rules

- Do not remove a public API solely because it is experimental.
- Experimental status must be visible at import documentation and first use.
- Any behavior change to inherited python-docx APIs requires a compatibility
  note and an inherited-suite regression test.
- `PaperRefusal` remains reserved for safe, anticipated inability to perform a
  documented operation. `TypeError` and `ValueError` remain programmer errors;
  `OSError` remains an environmental failure.
- Typed reports must use explicit schema and version fields when intended for
  agent or JSON consumption.
- Stable APIs must return or expose enough information to prove target identity,
  changed structures, refusals, and blind regions.
- Protection remains advisory. Documentation must never describe it as a
  security boundary.

## Delivery sequence

### Phase 0: Stop overclaiming

- Correct Gavin's README and RST language.
- Mark the stability tiers.
- Correct test counts and stock-package comparisons.
- Publish the same-namespace ADR.

Phase 0 may merge before implementation only if the documentation describes
current behavior honestly. It must not claim that later phases are complete.

### Phase 1: Repair core correctness

- Complete shared destination writing and `patch_save` parity.
- Scope and test the refusal contract.
- Improve mixed-install detection.
- Fix the vacuous test and source Ruff findings.

No new public workflow APIs should land during this phase.

### Phase 2: Establish external compatibility

- Add the Word fixture corpus and Word automation.
- Validate revision accept/reject, comments, controls, fields, tables, and lists.
- Add canonicalizer adversarial tests and package fuzzing.

### Phase 3: Promote or isolate workflows

- Promote provisional APIs that satisfy their Word fixture gates.
- Reassess `compare`, scrub/finalize, and composition.
- Extract low-signal workflow modules if their maintenance cost exceeds their
  demonstrated use.

### Phase 4: Measure causal value

- Run the matched evaluation.
- Use failure modes and mechanism metrics, not only aggregate score, to decide
  the next package investments.

## Release gates

The next release presented as stable must satisfy all of the following:

- no known path-save inconsistency between `Document.save()` and `patch_save()`;
- all stable mutators appear in the contract-test registry;
- both mixed-install orders and a transitive clobber fail the doctor;
- README and RST claims match current implementation;
- inherited pytest and Behave suites pass;
- Paper-specific tests pass with no vacuous assertions;
- Ruff is clean and required;
- the selected Pyright boundary is clean or Pyright is no longer advertised;
- the initial Microsoft Word fixture bucket passes its release workflow;
- supported Paper redlines open in Word and accept/reject to expected states;
- experimental APIs are labeled and excluded from stable guarantees;
- package-preflight regression corpus passes; and
- the upstream-maintenance document names the current base and conflict owners.

## Work breakdown

| ID | Priority | Area | Deliverable | Completion evidence |
| --- | --- | --- | --- | --- |
| P0-1 | Blocker | Saving | Shared durable destination writer | Symlink, failure-injection, permission, and fsync tests |
| P0-2 | Blocker | Safety contract | Scoped refusal contract and registry | Every stable mutator mapped to success/refusal/rollback tests |
| P0-3 | Blocker | Packaging | Mixed-install hardening and ADR | Both orders plus transitive-clobber CI fixture |
| P0-4 | Blocker | Documentation | Correct Gavin README/RST | Review checklist contains no known false claim |
| P1-1 | High | Compatibility | Word-authored fixture corpus | Word open, display, accept, reject, and preservation results |
| P1-2 | High | Compare | Independent validation | Adversarial canonicalizer tests and Word resolution |
| P1-3 | High | Delivery safety | Scrub/finalize inventory and certification | Word-authored privacy fixtures and explicit blind-region refusal |
| P1-4 | High | Composition | Bounded experimental contract | Fixed assertion, invariant tests, Word fixtures, remapping report |
| P1-5 | High | Security | ZIP/OPC fuzzing | Reproducible target, corpus, scheduled run, minimized regressions |
| P1-6 | High | Quality | Required Ruff and scoped typing | Clean required CI gates |
| P1-7 | Medium | Reproducibility | Determinism definition | Cross-environment hash matrix and scoped docs |
| P2-1 | Medium | Maintenance | Upstream sync policy | Base inventory, owners, cadence, and compatibility process |
| P2-2 | Medium | Evaluation | Matched stock/Paper experiment | Immutable results with treatment compliance and mechanism metrics |

## Definition of done

This proposal is complete when the blocker items are implemented, the stable
surface has exact contract coverage, Microsoft Word provides an external
compatibility oracle for the supported revision workflow, the documentation no
longer claims more than the code proves, and every remaining public workflow is
clearly classified as stable, provisional, or experimental.
