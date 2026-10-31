/**
 * Revision-bound target handles.
 *
 * Inspection returns a short handle (`T1`, `T2`, ...) per paragraph instead of making the agent
 * copy and resend `exactText`, `paragraphId` and `fingerprint`. A handle remembers the canonical
 * target descriptor and the package revision it was inspected at, and is refused once the document
 * has changed, so an edit can never silently land on a paragraph that moved or was rewritten.
 */

export function createTargetRegistry() {
  return {
    nextId: 1,
    revision: null,
    entries: new Map(),
    identities: new Map(),
    retired: new Set()
  };
}

function descriptorFor(paragraph) {
  const revisionView = paragraph.revisionView === "rejected" ? "rejected" : "accepted";
  return {
    exactText: paragraph.exactText,
    ...(paragraph.paragraphId ? { paragraphId: paragraph.paragraphId } : { index: paragraph.index }),
    ...(paragraph.fingerprint ? { fingerprint: paragraph.fingerprint } : {}),
    inTable: paragraph.inTable === true,
    revisionView
  };
}

/** Handles from an earlier revision stay known (so they report as stale rather than unknown). */
function retireOlderThan(registry, revision) {
  if (registry.revision === revision) return;
  for (const handle of registry.entries.keys()) registry.retired.add(handle);
  registry.entries.clear();
  registry.identities.clear();
  registry.revision = revision;
}

export function registerParagraph(registry, paragraph, revision) {
  retireOlderThan(registry, revision);
  const descriptor = descriptorFor(paragraph);
  const identity = [
    descriptor.revisionView,
    descriptor.paragraphId || `i${paragraph.index}`,
    descriptor.fingerprint || ""
  ].join(":");

  let handle = registry.identities.get(identity);
  if (!handle) {
    handle = `T${registry.nextId++}`;
    registry.identities.set(identity, handle);
    registry.entries.set(handle, {
      handle,
      revision,
      descriptor,
      humanReference: paragraph.humanReference || null,
      index: paragraph.index
    });
  }
  return handle;
}

export function resolveHandle(registry, handle, revision) {
  const name = String(handle);
  const entry = registry.entries.get(name);
  if (entry && entry.revision === revision) return { entry };
  if (entry || registry.retired.has(name)) {
    return {
      error: {
        code: "STALE_TARGET_HANDLE",
        message: `Target handle ${name} was issued for an earlier version of the document. Edits, accept/reject and write_part change the document, which retires every handle.`
      }
    };
  }
  return {
    error: {
      code: "TARGET_HANDLE_NOT_FOUND",
      message: `Target handle ${name} is not known for this document. Handles come from inspect_document (or refreshedTargets in an edit result).`
    }
  };
}
