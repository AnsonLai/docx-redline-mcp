# AGENTS.md — Launch Card

Use this file to route work on `docx-redline-mcp`, a thin stdio MCP server over
[`@ansonlai/docx-redline-js`](https://github.com/AnsonLai/docx-redline-js). Do not explore the whole
repository before acting, and never read `node_modules/@ansonlai/docx-redline-js/dist/` to infer library
behavior: use the library's documentation (exposed by this server as `docx-redline://` resources, or under
`node_modules/@ansonlai/docx-redline-js/docs/`).

## Pick the route

| Task | Start here |
|---|---|
| Use the server to edit a `.docx` | The server `instructions` (`src/tool-definitions.js`) and [README](README.md) |
| Change what an agent sees (tool text, annotations, input schemas) | `src/tool-definitions.js`, `src/schemas.js`, then register in `src/server.js` |
| Change how an edit is applied or committed | `src/lib/operation-service.js` → `src/lib/session-apply.js` |
| Change handle-based editing | `src/lib/target-registry.js`, `src/lib/edit-translator.js`, `src/tools/edit-document.js` |
| Change review resolution (accept/reject/comments) | `src/tools/resolve-review.js`, maintenance operations in `src/lib/operation-service.js` |
| Change saving, closing or the session | `src/tools/save-docx.js`, `src/tools/close-docx.js`, `src/lib/docx-package.js`, `src/lib/session-store.js` |
| Upgrade the library | [Upgrading the library](#upgrading-the-library) |
| Add or fix tests | The closest `tests/*.test.js`; fixtures come from `tests/helpers/fixtures.js` |

## Design invariants

These are the properties the tests defend. Keep them when changing anything.

1. **Delegate, don't reimplement.** Mutation, targeting, revision allocation, comments, numbering, ZIP handling,
   validation and rollback belong to the library's `openDocx` facade. A tool translates a narrow request into
   canonical operations (see the library's `docs/schemas/document-operations.schema.json`) and calls the same
   apply path. Only server concerns live here: sessions, handles, transport, output paths, authorization scope.
2. **A failed request never changes the session.** `applyOperationsToPackage` works on a copy and the result is
   committed (`session-apply.js`) only after `validateSessionPackage` passes.
3. **Handles are revision-bound.** `inspect_document` issues `T1`, `T2`, ... bound to the package revision they
   were inspected at. Any change to the package retires them (`STALE_TARGET_HANDLE`); `edit_document` returns
   `refreshedTargets` for the paragraphs it edited. Never accept a handle without checking its revision.
4. **The source file is never overwritten implicitly.** `save_docx` writes `<name>.reviewed.docx` unless
   `outputPath` or `inPlace: true` is given. `close_docx` refuses to drop unsaved edits.
5. **Other people's review work needs an explicit scope.** `resolve_review` refuses accept/reject/delete without
   `author`, `commentIds` or `allAuthors: true` (`REVIEW_SCOPE_REQUIRED`, `requiresUserAuthorization`). Library
   refusals that need the user (`COMMENTED_CONTENT_DELETE`, ...) must reach the agent unchanged.
6. **Report the actionable error.** A batch failure is surfaced with the failing operation's own error at
   `error` (`chooseActionableError`), never the generic `BATCH_OPERATION_FAILED`. Failures are `isError` results
   carrying `error.recovery` and `retryPlan`.
7. **`completion: true` is the only success signal.** Safe defaults (`atomic`, `strictTargets`, fidelity
   `enforce`) are set in `resolveApplyOptions`; tools must not expose `verifyRevisionFidelity: "off"`.
8. **stdout belongs to the protocol.** All logging goes to stderr (`src/lib/mcp-logger.js`).
9. **Tool text is part of the interface.** Each description states use case, effects, retry safety, success
   criterion and common errors. Optional capabilities belong in the library guides (MCP resources), not in every
   description. Annotations (`readOnlyHint`, `destructiveHint`) must stay honest; `tests/stdio-server.test.js`
   checks them.
10. **The library is pinned and declared.** `package.json` pins the exact release and
    `src/lib/compatibility.js` declares the tested one; a test keeps them equal.

## Code map

```text
src/index.js              stdio entry point
src/server.js             tool and resource registration, structured error results
src/tool-definitions.js   server instructions, tool titles, descriptions, annotations
src/schemas.js            shared zod input schemas
src/tools/                one handler per tool
src/lib/
  operation-service.js    opens the library facade, runs a batch on a copy, builds the result payload
  session-apply.js        commits a validated result to the session
  edit-translator.js      edit_document edits and targetHandle -> canonical operations
  target-registry.js      revision-bound target handles
  errors.js               OperationFailedError and the recovery-envelope request errors
  compatibility.js        tested library release and mismatch warning
  library-docs.js         library schema and guides served as MCP resources
  redline-preprocess.js   strips hidden field scaffolding / proofing marks from a redline's target paragraph
  docx-package.js         read/write package parts, derived output paths, atomic save
  package-validation.js   comment-boundary and package checks run before committing
tests/*.test.js           node:test suites; tests/helpers/fixtures.js builds small .docx files
```

## Upgrading the library

The library is pre-1.0, so each release is unverified until proven:

1. Read the library `CHANGELOG.md` entries between the tested release and the new one. Look for changes to
   result fields, error codes, `existingRevisions` and fidelity behavior, and anything touching
   `DocxDocument`.
2. Change `TESTED_LIBRARY_RELEASE` in `src/lib/compatibility.js` and the pinned version in `package.json`
   (`npm install`), then run `npm test`. Update the prose in the README if behavior changed.
3. Do not add a feature just because the library has it: add it when an ordinary edit workflow needs it.

## Verification

Run the closest suite first, for example `node --test tests/agent-workflow.test.js`, then `npm test`.
`tests/real-docx-diagnostics.test.js` needs `simulation/saas-agreement-2023-02-16.docx`, which is gitignored; those
four tests fail without it and are unrelated to code changes. Real Word behavior (opening a saved file) is not
covered by these tests.
