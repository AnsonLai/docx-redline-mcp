import { expandOperationHandles } from "../lib/edit-translator.js";
import { resolveHandle } from "../lib/target-registry.js";
import {
  DEFAULT_AUTHOR,
  isMaintenanceOperation,
  openSessionDocument,
  resolveApplyOptions
} from "../lib/operation-service.js";

/**
 * Dry-runs a batch against the open document: resolves every target, reports conflicts and the
 * artifacts (comments, numbering) the batch would need, and changes nothing.
 */
export function createPreflightOperationsHandler({ getSession }) {
  return async function preflightOperations({ handle, operations, author, options }) {
    const list = Array.isArray(operations) ? operations.filter(Boolean) : [];
    if (list.length === 0) {
      throw new Error("At least one operation is required");
    }

    const maintenance = list.findIndex(isMaintenanceOperation);
    if (maintenance >= 0) {
      throw new Error(
        `Operation ${maintenance + 1} is a maintenance operation (${list[maintenance].type}); `
        + "preflight_operations only checks document edits. Apply maintenance operations directly."
      );
    }

    const session = getSession(handle);
    const document = await openSessionDocument(session);
    const effectiveOptions = resolveApplyOptions(options);
    const resolvedAuthor = author || DEFAULT_AUTHOR;
    const revision = document.getRevisionToken().value;
    const expanded = expandOperationHandles(list, name => resolveHandle(session.targets, name, revision));
    const preflight = document.preflight(expanded, resolvedAuthor, effectiveOptions);

    return {
      handle: session.handle,
      packageRevision: document.getRevisionToken(),
      effectiveOptions: { author: resolvedAuthor, ...effectiveOptions },
      ...preflight
    };
  };
}
