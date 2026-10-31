import { createTargetRegistry } from "./target-registry.js";

function createHandle() {
  return `doc_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

export function createSessionStore() {
  const sessions = new Map();

  return {
    create(input) {
      const handle = createHandle();
      const session = {
        handle,
        sourcePath: input.sourcePath,
        zip: input.zip,
        dirty: false,
        metadata: input.metadata || null,
        savedPath: null,
        targets: createTargetRegistry()
      };
      sessions.set(handle, session);
      return session;
    },
    get(handle) {
      return sessions.get(String(handle)) || null;
    },
    update(handle, updater) {
      const key = String(handle);
      const current = sessions.get(key);
      if (!current) return null;
      const next = updater(current);
      sessions.set(key, next);
      return next;
    },
    delete(handle) {
      return sessions.delete(String(handle));
    },
    clear() {
      sessions.clear();
    }
  };
}
