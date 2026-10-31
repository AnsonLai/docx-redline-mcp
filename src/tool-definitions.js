/**
 * Agent-facing text for the server and its tools.
 *
 * Each description states the use case, side effects, retry safety, success criterion and the
 * errors worth knowing about, so a model can choose a tool and recover from a failure without
 * reading anything else. Optional capabilities (list conversion, restoring deleted text, headers
 * and footers, revision policies) live in the library guides exposed as MCP resources rather than
 * in every tool description.
 */

export const SERVER_INSTRUCTIONS = [
  "Edit Word .docx files as real tracked changes and comments (powered by @ansonlai/docx-redline-js).",
  "",
  "Ordinary edit: open_docx -> inspect_document (focused search) -> edit_document -> save_docx.",
  "1. inspect_document with a search and `around` context. Each paragraph has a `target` handle (T1, T2, ...).",
  "2. edit_document once with every change you intend. Use `replacements` for small literal changes and `desiredText` for a rewrite.",
  "3. save_docx. It writes a new <name>.reviewed.docx beside the source unless you pass outputPath or inPlace: true.",
  "",
  "Rules:",
  "- Handles are bound to the document version. Any edit retires them; use `refreshedTargets` from the edit result or inspect again.",
  "- Require `completion: true`. On failure follow `error.recovery.action` and `retryPlan`; never resend unchanged arguments.",
  "- If `error.recovery.requiresUserAuthorization` is true, stop and ask the user.",
  "- Accepting, rejecting or deleting other reviewers' changes or comments needs the user's explicit request: use resolve_review.",
  "- Cite paragraphs to the user by `humanReference`, `provision` or `nearestHeading`, never by index.",
  "- Keep relative wording relative unless the user asks for a concrete value.",
  "- Do not set options (atomic, existingRevisions, generateRedlines, ...) unless the user asked for that behavior.",
  "",
  "Other cases: apply_operations accepts any canonical operation (lists, tables, formatting, headers/footers, comment replies) and",
  "preflight_operations checks a batch without changing anything. The operation schema and guides are MCP resources (docx-redline://...)."
].join("\n");

const READ_ONLY = { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false };

export const TOOL_DEFINITIONS = Object.freeze({
  open_docx: {
    title: "Open a Word document",
    description: [
      "Open a local .docx and return a handle for the other tools.",
      "Also returns packageRevision, a summary (paragraph, revision and comment counts, existing reviewers, headers/footers) and the library compatibility.",
      "Effects: reads the file into memory; nothing is written. Retry-safe: each call creates a new handle.",
      "Errors: the path is missing or is not a .docx."
    ].join(" "),
    annotations: { title: "Open a Word document", ...READ_ONLY }
  },

  close_docx: {
    title: "Close a document session",
    description: [
      "Release a handle's memory.",
      "Effects: discards the in-memory document. Refused with UNSAVED_CHANGES if edits were never saved, unless discardChanges is true.",
      "Retry-safe: closing an unknown handle fails."
    ].join(" "),
    annotations: { title: "Close a document session", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false }
  },

  list_parts: {
    title: "List package parts",
    description: "List the files inside the open .docx package (XML parts and media). Read-only and retry-safe. Ordinary edits never need this.",
    annotations: { title: "List package parts", ...READ_ONLY }
  },

  read_part: {
    title: "Read a package part",
    description: "Read an OOXML part as XML text (maxChars truncates) or get a binary summary. Read-only and retry-safe. Prefer inspect_document for document content.",
    annotations: { title: "Read a package part", ...READ_ONLY }
  },

  write_part: {
    title: "Replace an XML part",
    description: [
      "Last-resort escape hatch: replace a writable XML part (document, comments, numbering, headers, footers, content types) with raw XML.",
      "Bypasses targeting, revision tracking and the library's validation of the edit, and retires all target handles.",
      "Effects: changes the in-memory document. Errors: malformed XML or a non-writable part. Prefer edit_document."
    ].join(" "),
    annotations: { title: "Replace an XML part", readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false }
  },

  inspect_document: {
    title: "Inspect the document",
    description: [
      "Find and read paragraphs. Use a focused `search` with `around` context rather than reading everything.",
      "Each paragraph carries a `target` handle for edit_document, its exact text, a human reference, heading/table/list context and tracked-change authors; comments and headers/footers are listed too.",
      "revisionView \"rejected\" shows text another reviewer deleted (to restore it).",
      "Results are paged: follow selection.nextAfter while selection.truncated is true.",
      "Effects: none beyond issuing handles, which are valid until the document next changes. Retry-safe.",
      "Returns packageRevision, usable as expectedRevision."
    ].join(" "),
    annotations: { title: "Inspect the document", ...READ_ONLY }
  },

  preflight_operations: {
    title: "Check operations without applying",
    description: [
      "Dry-run canonical operations: resolves every target, reports conflicts, the author(s) that would be used and the comment/numbering artifacts required.",
      "Effects: none. Retry-safe. Success: valid is true.",
      "Use before a large or unusual batch (lists, tables, formatting, headers/footers). Maintenance operations are not accepted."
    ].join(" "),
    annotations: { title: "Check operations without applying", ...READ_ONLY }
  },

  edit_document: {
    title: "Edit the document with tracked changes",
    description: [
      "Apply edits to inspected paragraphs as Word tracked changes. Preferred way to change a document.",
      "Each edit names a `target` handle and exactly one of: replacements (small literal changes), desiredText (full rewrite), delete, comment, restoreText.",
      "Send all intended edits in one call; targets bind to the document as it was inspected, so do not reorder around structural edits.",
      "Effects: changes the in-memory document and retires existing handles (new ones come back in refreshedTargets). Atomic: on any failure nothing is applied. Not retry-safe: re-inspect after a failure.",
      "Success: completion is true. Common errors: STALE_TARGET_HANDLE and TARGET_TEXT_MISMATCH (inspect again), EXISTING_REVISIONS (another reviewer's pending change; ask the user before changing existingRevisions), COMMENTED_CONTENT_DELETE (needs user authorization), REVISION_FIDELITY_VIOLATION (retry narrower, never disable the check)."
    ].join(" "),
    annotations: { title: "Edit the document with tracked changes", readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false }
  },

  apply_operations: {
    title: "Apply canonical operations",
    description: [
      "Apply canonical document operations as one batch: any type in the operation schema (lists, tables, formatting, comments and replies, headers/footers via `part`).",
      "Use edit_document for ordinary text edits. Targets may be a descriptor, free text, or `targetHandle` / `targetEndHandle` from inspect_document.",
      "Effects, retry safety, success criterion and errors are as for edit_document. expectedRevision refuses the batch if the document changed since that packageRevision.",
      "Legacy maintenance types (accept_all_tracked_changes, reject_all_tracked_changes, delete_all_comments) still work here but use resolve_review, which requires an explicit scope."
    ].join(" "),
    annotations: { title: "Apply canonical operations", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false }
  },

  apply_operation: {
    title: "Apply one canonical operation",
    description: "Same as apply_operations with a single operation.",
    annotations: { title: "Apply one canonical operation", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false }
  },

  resolve_review: {
    title: "Accept, reject or delete review markup",
    description: [
      "Resolve other people's review state: accept or reject tracked changes, delete comments, or resolve/reopen comment threads.",
      "Only call this when the user explicitly asked, and name the scope they gave: `author`, `commentIds`, or `allAuthors: true`. A request without an explicit scope is refused (REVIEW_SCOPE_REQUIRED) so it can be put to the user.",
      "Effects: accepted/rejected changes and deleted comments cannot be recovered except by not saving. Retires all target handles. Atomic.",
      "Success: completion is true; results lists what was done."
    ].join(" "),
    annotations: { title: "Accept, reject or delete review markup", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false }
  },

  save_docx: {
    title: "Save the document",
    description: [
      "Write the edited document to disk.",
      "By default it creates a new <name>.reviewed.docx beside the source (never replacing an existing file); pass outputPath for a specific .docx path, or inPlace: true to overwrite the source. Overwrite only when the user asked for it.",
      "Effects: writes one file atomically. Retry-safe for an explicit outputPath. Success: completion is true and savedPath is where it went."
    ].join(" "),
    annotations: { title: "Save the document", readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false }
  }
});
