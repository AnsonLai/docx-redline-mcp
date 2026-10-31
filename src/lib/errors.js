import { normalizeErrorWithRecovery } from "@ansonlai/docx-redline-js/services/error-recovery.js";

/**
 * Thrown when a request fails. `payload` is the structured failure (error code, recovery guidance,
 * per-operation results, retry plan) so MCP clients can follow `error.recovery` instead of parsing
 * a message.
 */
export class OperationFailedError extends Error {
  constructor(message, payload) {
    super(message);
    this.name = "OperationFailedError";
    this.payload = payload;
  }
}

export function describeError(error) {
  if (!error) return "Operation failed";
  const code = error.code ? `${error.code}: ` : "";
  return `${code}${error.message || "Operation failed"}`;
}

/**
 * A request the server refuses before reading or changing anything. Follows the library's
 * recovery envelope so it is handled exactly like a library error.
 */
export function requestError({
  code,
  message,
  field = null,
  action = "change_request",
  requiresReinspection = false,
  requiresUserAuthorization = false,
  failedIndexes = [],
  operationCount = 0
}) {
  const error = normalizeErrorWithRecovery({
    code,
    message,
    ...(field ? { field } : {}),
    recovery: {
      action,
      sameArgumentsSafe: false,
      requiresReinspection,
      requiresUserAuthorization
    }
  });
  const attempted = new Set(failedIndexes);
  const unattempted = [];
  for (let index = 1; index <= operationCount; index += 1) {
    if (!attempted.has(index)) unattempted.push(index);
  }
  return new OperationFailedError(describeError(error), {
    status: "error",
    completion: false,
    hasChanges: false,
    rolledBack: true,
    error,
    retryPlan: {
      base: "original",
      committedIndexes: [],
      failedIndexes,
      unattemptedIndexes: unattempted,
      replayWholeBatch: true,
      sameArgumentsSafe: false
    }
  });
}
