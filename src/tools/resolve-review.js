import { requestError } from "../lib/errors.js";
import { commitOperations } from "../lib/session-apply.js";

const ACTIONS = Object.freeze(["accept", "reject", "delete_comments", "resolve_comment", "reopen_comment"]);

function scopeError(message, field) {
  return requestError({
    code: "REVIEW_SCOPE_REQUIRED",
    message,
    field,
    action: "ask_user",
    requiresUserAuthorization: true
  });
}

/**
 * Builds the maintenance operations for a review-resolution request. The scope must be explicit:
 * accepting, rejecting or deleting other reviewers' work is never a default.
 */
export function buildReviewOperations({ action, author, allAuthors, commentIds }) {
  if (!ACTIONS.includes(action)) {
    throw requestError({
      code: "INVALID_AGENT_REQUEST",
      message: `action must be one of ${ACTIONS.join(", ")}.`,
      field: "action"
    });
  }

  const ids = Array.isArray(commentIds) ? commentIds.map(String) : [];
  const hasAuthor = typeof author === "string" && author.trim().length > 0;
  const scopes = [hasAuthor, allAuthors === true, ids.length > 0].filter(Boolean).length;

  if (action === "resolve_comment" || action === "reopen_comment") {
    if (ids.length === 0) {
      throw scopeError(`${action} needs commentIds (from inspect_document comments).`, "commentIds");
    }
    return ids.map(commentId => ({
      type: "resolve_comment",
      commentId,
      resolved: action === "resolve_comment"
    }));
  }

  if (action === "delete_comments") {
    if (scopes !== 1) {
      throw scopeError(
        "delete_comments needs exactly one scope: author, commentIds, or allAuthors: true. Ask the user which comments to remove.",
        "author"
      );
    }
    return [{
      type: "delete_all_comments",
      ...(ids.length ? { ids } : hasAuthor ? { author } : {})
    }];
  }

  // accept / reject operate on tracked changes, which have no ids to select by.
  if (ids.length > 0 || scopes !== 1) {
    throw scopeError(
      `${action} needs exactly one scope: author, or allAuthors: true. Ask the user whose tracked changes to ${action}.`,
      "author"
    );
  }
  return [{
    type: action === "accept" ? "accept_all_tracked_changes" : "reject_all_tracked_changes",
    ...(hasAuthor ? { author } : {})
  }];
}

export function createResolveReviewHandler({ getSession, store }) {
  return async function resolveReview({ handle, action, author, allAuthors, commentIds, options }) {
    const operations = buildReviewOperations({ action, author, allAuthors, commentIds });
    const session = getSession(handle);
    const { session: updated, payload, parts } = await commitOperations({
      session,
      store,
      operations,
      options
    });

    return {
      handle: updated.handle,
      dirty: updated.dirty,
      action,
      ...payload,
      parts
    };
  };
}
