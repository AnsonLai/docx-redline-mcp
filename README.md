# DOCX Redline MCP

Local `stdio` MCP server for inspecting and editing Microsoft `.docx` files through OOXML package parts.

It is intended for agent workflows that need to:
- open a local `.docx`
- inspect OOXML parts such as `word/document.xml`
- apply real Word tracked changes through `@ansonlai/docx-redline-js`
- add comments and maintenance operations safely
- save back in place or to a new output path

## Status

Current behavior is hardened around the failure modes discovered during end-to-end Word testing:
- MCP logging is routed to `stderr` so stdio transport is not corrupted
- package updates are validated before the in-memory session is committed
- invalid comment/revision combinations are rejected and rolled back
- `delete_all_comments` clears both inline anchors and `word/comments.xml`
- redline operations normalize hidden field scaffolding (`w:fldChar`, `w:instrText`) and proofing markers (`w:proofErr`) in the matched target paragraph before diffing
- batch redline application is supported through `apply_operations`

## Requirements

- Node.js 18 or newer

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

Open a `.docx` file and create a working handle.

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

### `apply_operation`

Apply one structured operation to `word/document.xml`.

Redline example:

```json
{
  "handle": "doc_...",
  "author": "Roo",
  "operation": {
    "type": "redline",
    "target": "Original text",
    "modified": "Updated text"
  }
}
```

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

Maintenance operations:
- `accept_all_tracked_changes`
- `reject_all_tracked_changes`
- `delete_all_comments`

### `apply_operations`

Apply multiple structured operations in a single batch. This is the preferred path when one edit session needs several redlines/comments against the same open document.

```json
{
  "handle": "doc_...",
  "author": "Roo",
  "operations": [
    {
      "type": "redline",
      "target": "Original text",
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
- On failure, the server rolls back the affected `word/document.xml` and `word/comments.xml` state.
- For redlines, the server preserves visible Word cross-reference output such as `9.1` or `10.6` but removes hidden field instruction scaffolding before diffing.
- Comment anchors that cross tracked insertion/deletion boundaries are rejected because Word commonly refuses to open those packages.

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
- package read/write/save
- rollback on invalid comment/revision markup
- batch operation handling
- real-DOCX diagnostic generation
- field/proofing normalization regressions

## License

MIT. See [LICENSE](./LICENSE).
