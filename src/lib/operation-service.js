import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import {
  areRevisionTokensEqual,
  configureXmlProvider,
  validateRevisionToken
} from "@ansonlai/docx-redline-js";
import { openDocx } from "@ansonlai/docx-redline-js/node";
import { validateSessionPackage } from "./package-validation.js";
import { preprocessDocumentXmlForOperation } from "./redline-preprocess.js";

configureXmlProvider({ DOMParser, XMLSerializer });

const textEncoder = new TextEncoder();

export const DEFAULT_AUTHOR = "AI Agent";

/**
 * Operations the MCP server handles through dedicated DocxDocument methods instead of the
 * operation runner. They act on the whole package (or a subset selected by author or id).
 */
export const MAINTENANCE_OPERATION_TYPES = Object.freeze([
  "accept_all_tracked_changes",
  "reject_all_tracked_changes",
  "delete_all_comments",
  "resolve_comment"
]);

const MAINTENANCE_TYPE_SET = new Set(MAINTENANCE_OPERATION_TYPES);

/**
 * Apply options forwarded to `DocxDocument.applyOperations`. `atomic` and `strictTargets` default
 * to the safe side (all-or-nothing, exact targets); everything else follows the library defaults.
 */
const FORWARDED_APPLY_OPTIONS = Object.freeze([
  "continueOnError",
  "existingRevisions",
  "generateRedlines",
  "pairReplacements",
  "removeFormatting",
  "sanitizeInput",
  "structuredContent",
  "strictListConversion",
  "verifyRevisionFidelity"
]);

/**
 * Thrown when an apply request fails. `payload` is the structured failure (error code, recovery
 * guidance, per-operation results, retry plan) so MCP clients can follow `error.recovery`.
 */
export class OperationFailedError extends Error {
  constructor(message, payload) {
    super(message);
    this.name = "OperationFailedError";
    this.payload = payload;
  }
}

export function isMaintenanceOperation(operation) {
  return MAINTENANCE_TYPE_SET.has(operationType(operation));
}

function operationType(operation) {
  return String(operation?.type || "").toLowerCase();
}

export function resolveApplyOptions(options = {}) {
  const resolved = {
    atomic: options.atomic !== false,
    strictTargets: options.strictTargets !== false
  };
  for (const key of FORWARDED_APPLY_OPTIONS) {
    if (options[key] !== undefined) {
      resolved[key] = options[key];
    }
  }
  return resolved;
}

/** Opens the session's current package as a library `DocxDocument` (the session is not modified). */
export async function openSessionDocument(session) {
  const bytes = await session.zip.generateAsync({ type: "uint8array" });
  return openDocx(bytes);
}

export function describeError(error) {
  if (!error) return "Operation failed";
  const code = error.code ? `${error.code}: ` : "";
  return `${code}${error.message || "Operation failed"}`;
}

function failureMessage(payload) {
  const failedResult = (payload.results || []).find(result => result?.status === "error" && result.error);
  const primary = payload.error || failedResult?.error;
  const base = describeError(primary);
  if (failedResult && payload.error && failedResult.error !== payload.error) {
    return `${base} (operation ${failedResult.index}: ${describeError(failedResult.error)})`;
  }
  return base;
}

function summarizeReceipt(receipt, offset) {
  if (!receipt) return receipt;
  return {
    operationIndex: (receipt.operationIndex ?? 0) + offset,
    ...(receipt.operationId ? { operationId: receipt.operationId } : {}),
    finalDisposition: receipt.finalDisposition,
    committed: receipt.committed,
    revisionItems: receipt.revisionItems,
    commentIds: receipt.commentIds,
    numberingIds: receipt.numberingIds,
    warnings: receipt.warnings
  };
}

function summarizeResult(result, offset) {
  // Receipts are returned once at the top level; match diagnostics are bulky and only useful for
  // debugging target resolution.
  const { receipt: _receipt, matchDiagnostics: _matchDiagnostics, ...rest } = result;
  return { ...rest, index: (result.index ?? 0) + offset };
}

function shiftIndexes(indexes, offset) {
  return Array.isArray(indexes) ? indexes.map(index => index + offset) : indexes;
}

function mergeRetryPlans(plans, operationCount, atomic) {
  if (plans.length === 0) return undefined;
  if (plans.length === 1 && plans[0].offset === 0 && operationCount === plans[0].count) {
    return plans[0].plan;
  }
  const merged = {
    base: atomic ? "original" : "output",
    committedIndexes: [],
    failedIndexes: [],
    unattemptedIndexes: [],
    replayWholeBatch: atomic,
    sameArgumentsSafe: false
  };
  for (const { plan, offset } of plans) {
    merged.committedIndexes.push(...shiftIndexes(plan.committedIndexes || [], offset));
    merged.failedIndexes.push(...shiftIndexes(plan.failedIndexes || [], offset));
    merged.unattemptedIndexes.push(...shiftIndexes(plan.unattemptedIndexes || [], offset));
  }
  return merged;
}

/** Splits a batch into runs of runner operations and individual maintenance operations. */
function segmentOperations(operations) {
  const segments = [];
  operations.forEach((operation, position) => {
    if (isMaintenanceOperation(operation)) {
      segments.push({ kind: "maintenance", operation, offset: position, count: 1 });
      return;
    }
    const last = segments[segments.length - 1];
    if (last && last.kind === "runner") {
      last.operations.push(operation);
      last.count += 1;
    } else {
      segments.push({ kind: "runner", operations: [operation], offset: position, count: 1 });
    }
  });
  return segments;
}

async function runMaintenanceOperation(document, operation, options) {
  const type = operationType(operation);
  const validate = options.validate !== false;

  if (type === "accept_all_tracked_changes" || type === "reject_all_tracked_changes") {
    const action = type === "accept_all_tracked_changes" ? "accept" : "reject";
    const author = typeof operation.author === "string" && operation.author.trim() ? operation.author : null;
    const result = await document.resolveRevisions(action, author
      ? { author, validate }
      : { allAuthors: true, validate });
    return { result, details: author ? { author } : { allAuthors: true } };
  }

  if (type === "delete_all_comments") {
    const ids = Array.isArray(operation.ids) ? operation.ids : null;
    const author = typeof operation.author === "string" && operation.author.trim() ? operation.author : null;
    const filter = ids ? { ids } : author ? { author } : { allAuthors: true };
    const result = await document.deleteComments({ ...filter, validate });
    return {
      result,
      details: {
        ...filter,
        commentsRemoved: result.commentsRemoved,
        referencesRemoved: result.referencesRemoved
      }
    };
  }

  // resolve_comment
  const commentId = operation.commentId ?? operation.id;
  if (commentId === undefined || commentId === null || commentId === "") {
    return {
      result: {
        status: "error",
        hasChanges: false,
        written: false,
        error: { code: "INVALID_OPERATION", field: "commentId", message: "resolve_comment requires commentId." }
      },
      details: {}
    };
  }
  const result = await document.resolveComment(commentId, {
    resolved: operation.resolved !== false,
    validate
  });
  return {
    result,
    details: { commentId: String(commentId), resolved: result.resolved, commentIds: result.commentIds }
  };
}

function maintenanceResult(index, operation, outcome) {
  const { result, details } = outcome;
  const failed = result.status === "error" || Boolean(result.error);
  return {
    index,
    type: operationType(operation),
    status: failed ? "error" : result.hasChanges ? "applied" : "no-op",
    ...(failed ? { error: result.error } : {}),
    ...details
  };
}

function revisionMismatchPayload(code, message, extra = {}) {
  return {
    status: "error",
    complete: false,
    hasChanges: false,
    rolledBack: true,
    error: {
      code,
      message,
      recovery: {
        action: "reinspect",
        requiresReinspection: true,
        requiresUserAuthorization: false,
        sameArgumentsSafe: false
      }
    },
    ...extra
  };
}

function checkExpectedRevision(document, expectedRevision) {
  if (!expectedRevision) return null;
  const validation = validateRevisionToken(expectedRevision);
  if (!validation.valid) {
    return revisionMismatchPayload(
      validation.error?.code || "INVALID_REVISION_TOKEN",
      validation.error?.message || "Invalid revision token."
    );
  }
  const current = document.getRevisionToken();
  if (expectedRevision.scope !== current.scope) {
    return revisionMismatchPayload(
      "REVISION_TOKEN_SCOPE_MISMATCH",
      `Revision token scope mismatch: expected '${current.scope}', got '${expectedRevision.scope}'. Use the packageRevision returned by open_docx, inspect_document or apply_operations.`,
      { currentRevision: current }
    );
  }
  if (!areRevisionTokensEqual(current.value, expectedRevision.value)) {
    return revisionMismatchPayload(
      "REVISION_MISMATCH",
      `Document revision mismatch: expected '${expectedRevision.value}', current is '${current.value}'. The document changed since it was inspected; inspect it again.`,
      { expectedRevision, currentRevision: current }
    );
  }
  return null;
}

/**
 * Hidden field scaffolding and proofing markers in a redline's target paragraph confuse the diff
 * and have produced packages Word refuses to open, so they are normalized first. The visible
 * field result text is preserved.
 */
function preprocessOperations(document, operations) {
  const documentBytes = document.entries.get("word/document.xml");
  if (!documentBytes) return;

  const original = new TextDecoder().decode(documentBytes);
  let xml = original;
  for (const operation of operations) {
    if (isMaintenanceOperation(operation)) continue;
    xml = preprocessDocumentXmlForOperation(xml, operation);
  }
  if (xml !== original) {
    document.entries.set("word/document.xml", textEncoder.encode(xml));
  }
}

/**
 * Applies a batch of operations to an in-memory package without touching the session.
 *
 * Returns `{ payload, bytes }`. `bytes` is the resulting package when the batch changed it and is
 * safe to commit, otherwise `null`. A payload with `status: "error"` must be surfaced as a failure.
 */
export async function applyOperationsToPackage(session, operations, settings = {}) {
  const author = settings.author || DEFAULT_AUTHOR;
  const options = resolveApplyOptions(settings.options);
  const document = await openSessionDocument(session);
  const startingRevision = document.getRevisionToken();

  const mismatch = checkExpectedRevision(document, settings.expectedRevision);
  if (mismatch) {
    return { payload: { ...mismatch, effectiveOptions: options }, bytes: null };
  }

  preprocessOperations(document, operations);

  const results = [];
  const receipts = [];
  const warnings = [];
  const retryPlans = [];
  const authorsUsed = new Set();
  const artifactsChanged = new Set();
  let hasChanges = false;
  let failed = false;
  let rolledBack = false;
  let topLevelError = null;
  let validation = null;

  for (const segment of segmentOperations(operations)) {
    const { offset } = segment;

    if (segment.kind === "maintenance") {
      const outcome = await runMaintenanceOperation(document, segment.operation, options);
      const result = maintenanceResult(offset + 1, segment.operation, outcome);
      results.push(result);
      if (result.status === "error") {
        failed = true;
        topLevelError = topLevelError || result.error;
      }
      hasChanges = hasChanges || Boolean(outcome.result.hasChanges);
      for (const name of outcome.result.artifactsChanged || []) artifactsChanged.add(name);
      if (failed && options.atomic) break;
      continue;
    }

    const runnerOptions = { ...options, author };
    const outcome = await document.applyOperations(segment.operations, runnerOptions);

    for (const result of outcome.results || []) results.push(summarizeResult(result, offset));
    for (const receipt of outcome.receipts || []) receipts.push(summarizeReceipt(receipt, offset));
    for (const warning of outcome.warnings || []) warnings.push(warning);
    for (const name of outcome.authorsUsed || []) authorsUsed.add(name);
    for (const name of outcome.artifactsChanged || []) artifactsChanged.add(name);
    validation = outcome.validation || validation;
    if (outcome.retryPlan) retryPlans.push({ plan: outcome.retryPlan, offset, count: segment.count });

    const segmentFailed = outcome.status === "error"
      || outcome.status === "partial"
      || (outcome.results || []).some(result => result?.status === "error");
    if (outcome.rolledBack) rolledBack = true;
    hasChanges = hasChanges || Boolean(outcome.hasChanges && outcome.written !== false);

    if (segmentFailed) {
      failed = true;
      topLevelError = topLevelError || outcome.error || null;
      if (options.atomic) break;
    }
  }

  const completedAll = !failed;
  const status = !failed ? "ok" : hasChanges && !options.atomic ? "partial" : "error";
  const finalBytes = (status === "ok" || status === "partial") && hasChanges
    ? document.toUint8Array()
    : null;

  const payload = {
    status,
    complete: completedAll,
    hasChanges: status === "error" ? false : hasChanges,
    ...(status === "error" ? { rolledBack: true } : rolledBack ? { rolledBack } : {}),
    effectiveOptions: { author, ...options },
    results,
    receipts,
    authorsUsed: [...authorsUsed],
    artifactsChanged: [...artifactsChanged],
    ...(warnings.length ? { warnings } : {}),
    ...(topLevelError ? { error: topLevelError } : {}),
    ...(validation ? { validation } : {}),
    ...(failed
      ? { retryPlan: mergeRetryPlans(retryPlans, operations.length, options.atomic) }
      : {}),
    packageRevision: status === "error" ? startingRevision : document.getRevisionToken()
  };

  if (status === "error" && !payload.error) {
    const failedResult = results.find(result => result.status === "error");
    if (failedResult?.error) payload.error = failedResult.error;
  }

  return { payload, bytes: finalBytes };
}

/**
 * Applies operations to the session, committing only if the whole result validates. Throws
 * `OperationFailedError` (with a structured `payload`) on failure; the session is left untouched.
 */
export async function applyOperationsToSession({ session, store, operations, author, options, expectedRevision }) {
  const { payload, bytes } = await applyOperationsToPackage(session, operations, {
    author,
    options,
    expectedRevision
  });

  if (payload.status === "error") {
    throw new OperationFailedError(failureMessage(payload), payload);
  }

  let candidate = null;
  if (bytes) {
    candidate = await JSZip.loadAsync(bytes);
    try {
      await validateSessionPackage(candidate);
    } catch (error) {
      throw new OperationFailedError(error.message, {
        ...payload,
        status: "error",
        complete: false,
        hasChanges: false,
        rolledBack: true,
        error: {
          code: "PACKAGE_VALIDATION_FAILED",
          message: error.message,
          recovery: {
            action: "change_request",
            requiresReinspection: true,
            requiresUserAuthorization: false,
            sameArgumentsSafe: false
          }
        }
      });
    }
  }

  return { payload, candidate };
}
