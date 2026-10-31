# DOCX Redline MCP

Local `stdio` MCP server that lets an agent review and edit Microsoft `.docx` files as real Word tracked changes and
comments. It is a thin layer over [`@ansonlai/docx-redline-js`](https://github.com/AnsonLai/docx-redline-js)
(pinned to `0.11.1`), which does the editing, validation and rollback.

It is designed for agent workflows that need to:
- find the right clause with a focused search, then edit it by a short handle
- apply tracked changes, comments and replies that Word shows as a normal review
- check a batch before applying it, and get actionable recovery advice when something is refused
- resolve other reviewers' changes and comments only when the user has said so
- save a reviewed copy without overwriting the original

## Requirements

- Node.js 20 or newer (required by the library)

## Install and run

```bash
npm install
node src/index.js
```

Package binary:

```bash
npx docx-redline-mcp
```

## The ordinary workflow

Most edits need four calls:

1. **`open_docx`** returns a `handle` for the document.
2. **`inspect_document`** with a focused `search` (and `around` for context). Every paragraph comes back with a
   `target` handle (`T1`, `T2`, ...), its exact text and a human reference such as the clause heading.
3. **`edit_document`** once, with every change you intend. Use `replacements` for small literal changes and
   `desiredText` for a rewrite.
4. **`save_docx`** writes `<name>.reviewed.docx` next to the original.

```json
{
  "handle": "doc_...",
  "author": "Roo",
  "edits": [
    {
      "target": "T3",
      "replacements": [{ "find": "may terminate", "replace": "may terminate on 30 days' written notice" }]
    },
    { "target": "T5", "comment": { "content": "Confirm the cap with finance", "anchorText": "$50,000" } }
  ]
}
```

A successful result looks like this (abridged):

```json
{
  "status": "ok",
  "completion": true,
  "hasChanges": true,
  "edits": [
    {
      "index": 1,
      "target": "T3",
      "humanReference": "12.1 Termination",
      "intent": "replacements",
      "status": "applied",
      "replacements": [{ "find": "may terminate", "replace": "...", "afterExcerpt": "The Supplier may terminate on 30 days' ..." }]
    },
    { "index": 2, "target": "T5", "intent": "comment", "status": "applied" }
  ],
  "refreshedTargets": [{ "previousHandle": "T3", "handle": "T9", "exactText": "...", "humanReference": "12.1 Termination" }],
  "receipts": [{ "operationIndex": 1, "committed": true, "revisionItems": [{ "id": "1000", "kind": "del" }] }],
  "packageRevision": { "algorithm": "sha256", "scope": "package", "value": "..." }
}
```

Treat only `completion: true` as success.

### Handles

A handle is bound to the version of the document it was inspected at. Any change to the document (an edit,
`resolve_review`, `write_part`) retires every handle, and using one fails with `STALE_TARGET_HANDLE` instead of
editing a paragraph that has moved or changed. `edit_document` returns `refreshedTargets` for the paragraphs it
edited so the next edit does not need another inspection; otherwise inspect again.

### When something fails

Failures come back as MCP error results (`isError: true`) with a structured body:

```json
{
  "status": "error",
  "completion": false,
  "rolledBack": true,
  "error": {
    "code": "COMMENTED_CONTENT_DELETE",
    "message": "Refusing to delete a paragraph with existing comments. ...",
    "operationIndex": 1,
    "recovery": {
      "action": "resolve_comments",
      "requiresReinspection": false,
      "requiresUserAuthorization": true,
      "sameArgumentsSafe": false
    }
  },
  "edits": [{ "index": 1, "target": "T4", "status": "error", "error": { "...": "..." } }],
  "retryPlan": { "base": "original", "failedIndexes": [1], "unattemptedIndexes": [], "replayWholeBatch": true }
}
```

- `error` is the failing operation's own error (not the generic "batch failed"). When several operations fail,
  one that needs the user's authorization is reported first.
- Follow `error.recovery.action`; never resend unchanged arguments. If `requiresUserAuthorization` is true, ask the user.
- A failed request never changes the open document.
- `retryPlan.base` is `original` (replay the whole batch) or `output` (only with `atomic: false`: keep what
  succeeded and retry `failedIndexes`).

## Tools

| Tool | Use it to | Changes the document |
|---|---|---|
| `open_docx` | Open a file; returns a handle, `packageRevision`, a summary and library compatibility | no |
| `inspect_document` | Find and read paragraphs; issues target handles | no |
| `preflight_operations` | Dry-run canonical operations | no |
| `edit_document` | Apply edits by handle (the ordinary path) | yes |
| `apply_operations` / `apply_operation` | Apply any canonical operation (lists, tables, formatting, headers/footers, comment replies) | yes |
| `resolve_review` | Accept/reject tracked changes, delete comments, resolve threads | yes |
| `save_docx` | Write the document to disk | writes a file |
| `close_docx` | Release a handle | no |
| `list_parts`, `read_part`, `write_part` | Raw package access (`write_part` is an escape hatch) | `write_part` only |

Tool annotations are set honestly (`readOnlyHint`, `destructiveHint`), so clients can prompt for the tools that matter.

### `open_docx`

```json
{ "path": "C:/contracts/master-services-agreement.docx" }
```

The response includes `packageRevision`, a `summary` (paragraph, revised-paragraph and comment counts, existing
tracked-change and comment authors, headers/footers) and `library` (`tested`, `installed`, `matches`).

### `inspect_document`

```json
{ "handle": "doc_...", "search": "termination", "around": 2 }
```

Filters: `search` (case-insensitive), `around`, `indexes`, `range`, `revisedOnly`, `inTable`, `skipEmpty`,
`revisionView`. Results are paged (default `limit` 100): follow `selection.nextAfter` as `after` while
`selection.truncated` is true. Use `revisionView: "rejected"` to see text another reviewer deleted.
`compact: false` returns the library's full records, including `paragraphId` and `fingerprint`.

Paragraphs carry `humanReference`, `provision` and `nearestHeading`; cite those to the user rather than
paragraph numbers. The response also lists comments with their anchors and the headers/footers that can be edited.

### `edit_document`

Each edit has a `target` handle and exactly one intent:

| Intent | Meaning |
|---|---|
| `replacements` | Small literal changes inside the paragraph (`find`, `replace`, optional `occurrence`) |
| `desiredText` | The complete new paragraph text, for a broad rewrite |
| `delete` | Delete the paragraph as a tracked change |
| `comment` | Add a comment (`content`, optional `anchorText`) |
| `restoreText` | Counterpropose text for a paragraph another reviewer deleted. The handle must come from `inspect_document` with `revisionView: "rejected"` |

Edits are translated into canonical operations and run through the same apply path as `apply_operations`.
Targets bind to the document as inspected, so list every edit in one call. The batch is atomic by default.

### `apply_operations`

For anything beyond text edits. Operations are the library's canonical operations
(`docx-redline://schema/document-operations`): `redline`, `insert`, `delete`, `restore`, `comment`, `comment_reply`,
`comment_resolve`, `format`, `character-format`, `highlight`, `paragraph-format`, `list-format`, `list-change`,
`table-reconciliation`. A target can be a descriptor, free text, or `targetHandle` / `targetEndHandle` from
`inspect_document`.

```json
{
  "handle": "doc_...",
  "author": "Roo",
  "expectedRevision": { "algorithm": "sha256", "version": 1, "scope": "package", "value": "..." },
  "operations": [
    { "type": "list-format", "targetHandle": "T4", "list": { "mode": "convert-plain", "format": "decimal" } },
    {
      "type": "redline",
      "part": { "kind": "footer", "type": "default" },
      "target": { "exactText": "Draft v1" },
      "modified": "Draft v2"
    }
  ]
}
```

- `expectedRevision` (a `packageRevision` from `open_docx`, `inspect_document` or an earlier apply) refuses the batch
  with `REVISION_MISMATCH` if the document changed since.
- `part` edits a header or footer listed by `open_docx` / `inspect_document`. Word does not allow comments there.
- Use `preflight_operations` first for a large or unusual batch: it resolves every target and reports conflicts
  without changing anything.

### Options

`edit_document`, `apply_operations`, `apply_operation`, `preflight_operations` and `resolve_review` accept `options`.
The defaults are the safe ones; set a field only for a stated reason.

| Option | Default | Effect |
|---|---|---|
| `atomic` | `true` | Any failing operation rolls the whole batch back. `false` keeps what succeeded (`status: "partial"`) |
| `strictTargets` | `true` | Targets must match the inspected text exactly |
| `existingRevisions` | `merge-same-author` | How to edit paragraphs that already have tracked changes. Another reviewer's pending changes are protected (`EXISTING_REVISIONS`); `slice-cross-author` edits inside them, `auto` chooses per paragraph |
| `generateRedlines` | `true` | `false` writes clean direct edits with no revision markup |
| `structuredContent` | `true` | Markdown tables, headings and lists in new text become Word structures; `false` keeps text literal |
| `verifyRevisionFidelity` | `enforce` | Refuse an edit whose Reject All would not restore the original. `warn` applies it and reports it |
| `strictListConversion` | `atomic` | Make unsupported list conversions errors instead of warnings |

Every result echoes `effectiveOptions`.

### `resolve_review`

Accepts or rejects tracked changes, deletes comments, and resolves or reopens comment threads. Call it only when the user
asked for it, and give the scope they named:

```json
{ "handle": "doc_...", "action": "accept", "author": "Alice" }
```

```json
{ "handle": "doc_...", "action": "delete_comments", "commentIds": ["3", "4"] }
```

| `action` | Scope |
|---|---|
| `accept`, `reject` | `author` or `allAuthors: true` |
| `delete_comments` | `author`, `commentIds` or `allAuthors: true` (deleting a thread root also removes its replies) |
| `resolve_comment`, `reopen_comment` | `commentIds` |

A request without exactly one explicit scope is refused with `REVIEW_SCOPE_REQUIRED` and
`requiresUserAuthorization: true`, so the question goes to the user instead of defaulting to "everyone".

### `save_docx`

```json
{ "handle": "doc_..." }
```

- By default writes `<name>.reviewed.docx` beside the source, never replacing an existing file
  (`.reviewed-2.docx`, ...). The source is untouched and `dirty` stays `true`.
- `outputPath` writes to a specific `.docx` path.
- `inPlace: true` overwrites the source. Only when the user asked for it.

The write is atomic (temp file, then rename). The result has `completion: true`, `savedPath` and `replacedSource`.

### `close_docx`

Releases a handle. Refused with `UNSAVED_CHANGES` if edits were never saved, unless `discardChanges: true`.

### Raw package access

`list_parts`, `read_part` and `write_part` expose the OOXML package. `write_part` replaces a writable XML part
(`word/document.xml`, `word/comments.xml`, `word/numbering.xml`, `word/header*.xml`, `word/footer*.xml`,
`[Content_Types].xml`) after validating that it is well-formed XML. It bypasses the library's revision tracking and
retires all handles; prefer `edit_document`.

## MCP resources and instructions

The server sends a short launch card as its MCP `instructions`, and serves reference material from the pinned
library release:

| URI | Content |
|---|---|
| `docx-redline://schema/document-operations` | JSON schema for every operation |
| `docx-redline://docs/agent-fast-start` | The inspect-once, apply-once workflow and recovery rules |
| `docx-redline://docs/agent-knowledge-base` | Detailed operation shapes, list and table handling, revision policies, error codes |

## Behavior notes

- Each `open_docx` call creates an in-memory working session; nothing reaches disk until `save_docx`.
- Edits are computed on a copy and committed only if the resulting package validates, so a refused or failed request
  leaves the session exactly as it was.
- The library verifies that Word's Reject All would restore the original; an edit that would not is refused
  (`REVISION_FIDELITY_VIOLATION`) with advice on how to narrow it.
- Comment anchors that cross tracked insertion/deletion boundaries are rejected because Word commonly refuses to open
  those packages.
- Redline targets have hidden field scaffolding (`w:fldChar`, `w:instrText`) and proofing markers (`w:proofErr`)
  removed before diffing, keeping the visible field result such as `9.1`. With `atomic: false`, a redline that later
  fails can still leave its paragraph with that scaffolding removed.
- Non-XML parts are read-only.

## Logging

- Default MCP log level is `warn`
- Logs are written to `stderr`, never `stdout`, which carries the protocol
- To enable engine diagnostics:

```bash
DOCX_REDLINE_MCP_LOG_LEVEL=info node src/index.js
```

A warning is written to `stderr` at startup if the installed library is not the release this server was tested with.

## Troubleshooting

- `Transport closed` during a tool call: check for any dependency or local logging writing to `stdout`.
- `STALE_TARGET_HANDLE`: the document changed since the handle was issued. Use `refreshedTargets` or inspect again.
- `REVISION_MISMATCH`: the document changed since the `expectedRevision` you sent. Inspect again.
- `EXISTING_REVISIONS`: the paragraph has another reviewer's pending changes. Ask the user before using
  `existingRevisions: "slice-cross-author"`.
- Word rejects a saved file even though the XML looks valid: prefer `edit_document` over `write_part` so the library's
  validation and the field normalization run.
- `COMMENT_ANCHOR_CROSSES_REVISION`: a comment's range crosses a tracked change (the default anchor is the whole
  paragraph, which includes any change you just made). Set `comment.anchorText` to unchanged text.

## Upgrading from 0.1.x

- The library is pinned to `@ansonlai/docx-redline-js@0.11.1` and requires Node.js 20 or newer.
- `save_docx` no longer overwrites the source by default; it writes `<name>.reviewed.docx`. Pass `inPlace: true` for the old behavior.
- Batches are atomic and targets strict by default. Pass `options: { atomic: false }` or `{ strictTargets: false }` to loosen them.
- Failures are MCP error results with `error.code`, `error.recovery` and `retryPlan` instead of a bare message.
- Apply results report `status`, `completion`, per-operation `results`, `receipts`, `warnings` and `packageRevision`.
- Accepting, rejecting and deleting comments moved to `resolve_review`, which requires an explicit scope. The old
  `accept_all_tracked_changes`, `reject_all_tracked_changes` and `delete_all_comments` operations still work inside
  `apply_operations` (with optional `author` / `ids`) but are deprecated.
- New: `inspect_document`, `edit_document`, `preflight_operations`, `resolve_review`, `close_docx`, `expectedRevision`,
  header/footer editing, library resources, and the MCP `instructions` launch card.

## Test

```bash
npm test
```

The suite covers tool handlers, the handle registry and edit translation, inspection, atomic and partial apply,
revision-token checks, header/footer edits, review resolution and its authorization scope, safe saving, library
compatibility, and the stdio transport (instructions, annotations, structured errors, clean stdout).

`tests/real-docx-diagnostics.test.js` needs a local `simulation/saas-agreement-2023-02-16.docx` (gitignored) and fails
without it. See [AGENTS.md](AGENTS.md) for the design invariants and how to upgrade the library.

## License

MIT. See [LICENSE](./LICENSE).
