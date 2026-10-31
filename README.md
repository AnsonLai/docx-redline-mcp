# DOCX Redline MCP

Local `stdio` MCP server for inspecting and editing Microsoft `.docx` files through OOXML package parts.

It is intended for agent workflows that need to:
- open a local `.docx`
- inspect OOXML parts such as `word/document.xml`
- inspect paragraphs, tracked changes, comments and headers/footers with stable target descriptors
- apply real Word tracked changes through `@ansonlai/docx-redline-js` (pinned to `0.11.1`)
- dry-run a batch of edits before touching the document
- add comments, reply to and resolve threads, and run review maintenance operations safely
- save back in place or to a new output path

## Status

All document edits go through the library's `openDocx` facade, which performs the edit, package wiring (comments,
numbering, content types), validation and rollback as one transaction. The server adds its own checks on top.

Current behavior is hardened around the failure modes discovered during end-to-end Word testing:
- MCP logging is routed to `stderr` so stdio transport is not corrupted
- package updates are validated before the in-memory session is committed
- invalid comment/revision combinations are rejected and rolled back
- `delete_all_comments` clears both inline anchors and `word/comments.xml`
- redline operations normalize hidden field scaffolding (`w:fldChar`, `w:instrText`) and proofing markers (`w:proofErr`) in the matched target paragraph before diffing
- batch redline application is supported through `apply_operations`
- Reject All fidelity is verified by the library: an edit whose Reject All would not restore the original is refused
- footnote/endnote references, list structure, tables and existing revisions from other authors are handled by the library

## Requirements

- Node.js 20 or newer (required by `@ansonlai/docx-redline-js`)

## Install

```bash
npm install
```

## Run

```bash
node src/index.js
```

Package binary:

```bash
npx docx-redline-mcp
```

## MCP Tools

### `open_docx`

Open a `.docx` file and create a working handle. The response also carries `packageRevision` and a `summary`: paragraph,
revised-paragraph and comment counts, existing tracked-change and comment authors, and the headers/footers present.

```json
{
  "path": "C:/contracts/master-services-agreement.docx"
}
```

### `list_parts`

List package parts for an open handle.

```json
{
  "handle": "doc_..."
}
```

### `read_part`

Read an OOXML part as XML text or get a binary summary.

```json
{
  "handle": "doc_...",
  "partPath": "word/document.xml",
  "maxChars": 12000
}
```

Binary export:

```json
{
  "handle": "doc_...",
  "partPath": "word/media/image1.png",
  "includeBase64": true
}
```

### `write_part`

Replace a writable XML part in memory.

```json
{
  "handle": "doc_...",
  "partPath": "word/document.xml",
  "xml": "<w:document>...</w:document>"
}
```

Writable XML parts:
- `word/document.xml`
- `word/comments.xml`
- `word/numbering.xml`
- `word/header*.xml`
- `word/footer*.xml`
- `[Content_Types].xml`

### `inspect_document`

Inspect the open document. This is the recommended first step: it returns the exact paragraph text and the target
descriptors (`paragraphId`, `fingerprint`) that operations should use, plus headings, table and list context,
tracked-change authors, comments with their anchors, and the headers/footers that can be targeted with `part`.

```json
{
  "handle": "doc_...",
  "search": "termination",
  "around": 2
}
```

Filters: `search` (case-insensitive), `around`, `indexes`, `range`, `revisedOnly`, `inTable`, `skipEmpty`,
`revisionView` (`accepted`, `rejected` or `current`; use `rejected` to find deleted text to restore).
Results are paged: the default `limit` is 100 paragraphs, and `selection.nextAfter` is the `after` cursor for the next page.
`compact: false` returns the library's full paragraph records.

The response includes `packageRevision`, a token for the current state of the package. Pass it back as
`expectedRevision` when applying so the edit is refused if the document changed in between.

### `preflight_operations`

Dry-run a batch against the open document. It resolves every target, reports conflicts between operations, the
author(s) that would be used and which comment/numbering artifacts the batch needs. Nothing is changed.

```json
{
  "handle": "doc_...",
  "author": "Roo",
  "operations": [
    { "type": "redline", "target": { "exactText": "Original text" }, "modified": "Updated text" }
  ]
}
```

### `apply_operation`

Apply one structured operation. Targets can be free text (`"target": "Original text"`) or, preferably, the descriptor
returned by `inspect_document`.

Redline example:

```json
{
  "handle": "doc_...",
  "author": "Roo",
  "operation": {
    "type": "redline",
    "target": { "exactText": "The Supplier may terminate.", "fingerprint": "fnv1a32:..." },
    "replacements": [{ "find": "may terminate", "replace": "may terminate on 30 days' notice" }]
  }
}
```

Use `modified` (the complete new paragraph text) for a broad rewrite, or `replacements` for small literal changes.

Comment example:

```json
{
  "handle": "doc_...",
  "author": "Roo",
  "operation": {
    "type": "comment",
    "target": "Original paragraph text",
    "textToComment": "Original",
    "commentContent": "Explain the change"
  }
}
```

Any operation type supported by the library can be sent, including `insert`, `delete`, `restore`, `comment_reply`,
`comment_resolve`, `format`, `character-format`, `highlight`, `paragraph-format`, `list-format` (convert typed markers to
a real list, change a list's format, join lists) and `table-reconciliation`. The canonical schema ships with the library
and is available as the `docx-redline://schema/document-operations` MCP resource.

Add `"part": { "kind": "footer", "type": "default" }` (or a part path such as `word/header2.xml`) to edit a header or
footer. `open_docx` and `inspect_document` list the available parts. Comments are not supported there (Word does not allow them).

Optional arguments on `apply_operation` and `apply_operations`:

| Argument | Purpose |
|----------|---------|
| `author` | Reviewer name stamped on tracked changes and comments (default `AI Agent`). |
| `expectedRevision` | `packageRevision` from `open_docx`, `inspect_document` or a previous apply. The edit is refused with `REVISION_MISMATCH` if the document changed since. |
| `options.atomic` | Default `true`: any failing operation rolls the whole batch back. `false` keeps the operations that succeeded (`status: "partial"`). |
| `options.strictTargets` | Default `true`: targets must match exactly. `false` allows looser text matching. |
| `options.existingRevisions` | How to edit paragraphs that already contain tracked changes: `merge-same-author` (default; another author's pending changes are protected), `slice-cross-author`, `auto`, `accept-all-first`, `reject-input`. |
| `options.generateRedlines` | `false` applies clean direct edits without revision markup. |
| `options.structuredContent` | `false` keeps Markdown in replacement text literal instead of building tables, headings and lists. |
| `options.verifyRevisionFidelity` | `enforce` (default) refuses edits whose Reject All would not restore the source; `warn` applies them and reports it. |
| `options.strictListConversion` | Treat unsupported list conversions as errors instead of warnings. |

Maintenance operations (send only when the user asked for them):

| Operation | Effect |
|-----------|--------|
| `{ "type": "accept_all_tracked_changes" }` | Accept tracked changes in the body, headers and footers. Add `"author": "Name"` to accept only that author's. |
| `{ "type": "reject_all_tracked_changes" }` | Reject tracked changes, optionally for one `author`. |
| `{ "type": "delete_all_comments" }` | Delete comments and their anchors. Add `"author"` or `"ids": ["3", "4"]` to delete a subset (deleting a thread root also removes its replies). |
| `{ "type": "resolve_comment", "commentId": "3" }` | Resolve the whole thread containing the comment. `"resolved": false` reopens it. |

Successful result (abridged):

```json
{
  "handle": "doc_...",
  "dirty": true,
  "status": "ok",
  "complete": true,
  "hasChanges": true,
  "results": [{ "index": 1, "type": "redline", "status": "applied", "resolvedTarget": { "...": "..." } }],
  "receipts": [{ "operationIndex": 1, "committed": true, "revisionItems": [{ "id": "1000", "kind": "del" }] }],
  "warnings": [],
  "authorsUsed": ["Roo"],
  "artifactsChanged": ["word/document.xml"],
  "packageRevision": { "algorithm": "sha256", "scope": "package", "value": "..." },
  "parts": ["..."]
}
```

`status` is `ok`, `partial` (only with `atomic: false`) or `error`. Treat only `complete: true` as full success.

Failures are returned as MCP error results (`isError: true`) whose structured content explains how to recover:

```json
{
  "status": "error",
  "rolledBack": true,
  "error": {
    "code": "TARGET_NOT_FOUND",
    "message": "...",
    "recovery": { "action": "reinspect", "requiresReinspection": true, "sameArgumentsSafe": false }
  },
  "results": [{ "index": 2, "status": "error", "error": { "...": "..." } }],
  "retryPlan": { "base": "original", "failedIndexes": [2], "unattemptedIndexes": [], "replayWholeBatch": true }
}
```

Follow `error.recovery.action` and never retry unchanged arguments. A failed apply never modifies the session.

### `apply_operations`

Apply several operations in a single batch. This is the preferred path when one edit session needs several
redlines/comments against the same open document. Targets bind at the start of the batch, so every target should come
from a single `inspect_document` call. Operations are atomic by default.

```json
{
  "handle": "doc_...",
  "author": "Roo",
  "expectedRevision": { "algorithm": "sha256", "version": 1, "scope": "package", "value": "..." },
  "operations": [
    {
      "type": "redline",
      "target": { "exactText": "Original text", "fingerprint": "fnv1a32:..." },
      "modified": "Updated text"
    },
    {
      "type": "comment",
      "target": "Other paragraph",
      "textToComment": "Other",
      "commentContent": "Reasoning"
    }
  ]
}
```

### `save_docx`

Save the working package.

In place:

```json
{
  "handle": "doc_..."
}
```

To a separate file:

```json
{
  "handle": "doc_...",
  "outputPath": "C:/contracts/master-services-agreement-copy.docx"
}
```

## Behavior Notes

- Each `open_docx` call creates an in-memory working session.
- `save_docx` writes back to the original file unless `outputPath` is supplied.
- Non-XML parts are read-only.
- `write_part` validates XML before accepting it.
- `apply_operation` and `apply_operations` validate the resulting package before updating the session state.
- A failed apply never changes the session: the result is computed on a copy and only committed once it validates.
- `write_part` bypasses the library; use it for last-resort XML fixes. A `packageRevision` taken before a `write_part` no longer matches afterwards.
- For redlines, the server preserves visible Word cross-reference output such as `9.1` or `10.6` but removes hidden field instruction scaffolding before diffing.
- Comment anchors that cross tracked insertion/deletion boundaries are rejected because Word commonly refuses to open those packages.
- With `atomic: false`, the field/proofing normalization above is applied to the target paragraph of every redline up front, so a redline that later fails can still leave its paragraph's hidden field scaffolding removed (the visible field text is kept).

## Upgrading from 0.1.x

- The library is pinned to `@ansonlai/docx-redline-js@0.11.1`; it requires Node.js 20 or newer.
- Batches are now **atomic by default** and targets are **strict** by default. Pass `options: { atomic: false }` or `{ strictTargets: false }` to restore the previous leniency.
- Failures are MCP error results with a structured payload (`error.code`, `error.recovery`, `retryPlan`) instead of a bare message.
- Apply results now include `status`, `complete`, per-operation `results`, `receipts`, `warnings` and `packageRevision`.
- `accept_all_tracked_changes`, `reject_all_tracked_changes` and `delete_all_comments` accept an optional `author` (and `ids` for comments); without them they still apply to everyone.
- New tools: `inspect_document` and `preflight_operations`. New resources expose the library's operation schema and agent guides.

## MCP Resources

The server exposes reference material shipped with the pinned library release:

| URI | Content |
|-----|---------|
| `docx-redline://schema/document-operations` | JSON schema for every operation |
| `docx-redline://docs/agent-fast-start` | Recommended inspect-once, apply-once workflow and recovery rules |
| `docx-redline://docs/agent-knowledge-base` | Detailed operation shapes, list/table handling and revision policies |

## Logging

- Default MCP log level is `warn`
- Logs are written to `stderr`, not `stdout`
- To enable engine diagnostics safely:

```bash
$env:DOCX_REDLINE_MCP_LOG_LEVEL = "info"
node src/index.js
```

## Troubleshooting

- `Transport closed` during a tool call:
  - check for any dependency or local logging writing to `stdout`
- Word rejects the saved file even though XML looks valid:
  - inspect the revised paragraph for hidden field runs (`w:fldChar`, `w:instrText`) or proofing markers (`w:proofErr`)
  - prefer the server redline path over direct XML surgery so the normalization step runs
- Comment operation rejected because of a tracked-change boundary:
  - re-anchor the comment to stable text outside the revised span

## Test

```bash
npm test
```

The test suite covers:
- basic tool handler behavior
- inspection, preflight, atomic and partial apply, revision-token checks, header/footer edits and review maintenance
- the stdio transport (tool and resource listing, structured error results, clean stdout)
- package read/write/save
- rollback on invalid comment/revision markup
- batch operation handling
- real-DOCX diagnostic generation
- field/proofing normalization regressions

## License

MIT. See [LICENSE](./LICENSE).
