import * as z from "zod/v4";

/**
 * Input schemas shared by the tool definitions. Descriptions are written for the agent that calls
 * the tool: what the value is, where it comes from and what the default does.
 */

export const documentHandle = z.string().min(1).describe("Handle returned by open_docx");

export const authorName = z.string().min(1).optional().describe(
  "Reviewer name stamped on tracked changes and comments. Defaults to \"AI Agent\"; use the name the user asked for."
);

export const revisionToken = z.object({
  algorithm: z.literal("sha256"),
  version: z.literal(1),
  scope: z.enum(["package", "document-parts"]),
  value: z.string().min(1),
  coveredParts: z.array(z.string()).optional()
}).describe("A packageRevision returned by open_docx, inspect_document or an apply. Pass it back unchanged.");

export const executionOptions = z.object({
  atomic: z.boolean().optional().describe(
    "Default true: if any operation fails, nothing is applied. false keeps the operations that succeeded (status \"partial\")."
  ),
  strictTargets: z.boolean().optional().describe(
    "Default true: targets must match the inspected text exactly. false allows looser matching."
  ),
  existingRevisions: z.enum([
    "merge-same-author",
    "slice-cross-author",
    "auto",
    "accept-all-first",
    "reject-input"
  ]).optional().describe(
    "Editing paragraphs that already contain tracked changes. Default merge-same-author protects other reviewers' pending changes (EXISTING_REVISIONS). slice-cross-author edits inside them; auto chooses per paragraph. Change it only when the user wants that."
  ),
  generateRedlines: z.boolean().optional().describe(
    "Default true (tracked changes). false writes clean direct edits with no revision markup; use only when the user asks for a clean edit."
  ),
  structuredContent: z.boolean().optional().describe(
    "Default true: Markdown tables, headings and lists in new text become native Word structures. false keeps new text literal."
  ),
  verifyRevisionFidelity: z.enum(["enforce", "warn"]).optional().describe(
    "Default enforce: refuse an edit whose Reject All would not restore the original. Never relax this to force an edit through."
  ),
  strictListConversion: z.boolean().optional().describe(
    "Treat unsupported list conversions as errors instead of warnings."
  ),
  pairReplacements: z.boolean().optional(),
  removeFormatting: z.boolean().optional(),
  continueOnError: z.boolean().optional()
}).describe("Execution policy. The defaults are the safe ones; set a field only for a stated reason.");

export const replacement = z.object({
  find: z.string().min(1).describe("Exact, case-sensitive text inside the target paragraph"),
  replace: z.string().describe("Replacement text (empty string deletes the found text)"),
  occurrence: z.number().int().min(1).optional().describe("1-based match when the text occurs more than once")
});

export const edit = z.object({
  target: z.string().min(1).describe("Target handle from inspect_document (or refreshedTargets), such as \"T3\""),
  replacements: z.array(replacement).min(1).optional().describe(
    "Small literal changes inside the paragraph. Preferred for most edits."
  ),
  desiredText: z.string().optional().describe(
    "The complete new paragraph text, for a broad rewrite. Keep unchanged wording identical."
  ),
  delete: z.literal(true).optional().describe("Delete the paragraph as a tracked change"),
  comment: z.object({
    content: z.string().min(1),
    anchorText: z.string().min(1).optional().describe("Text inside the paragraph to attach the comment to (default: the paragraph)")
  }).optional().describe("Add a Word comment to the paragraph"),
  restoreText: z.string().optional().describe(
    "Counterpropose text for a paragraph another reviewer deleted. The target must come from inspect_document with revisionView \"rejected\"."
  ),
  existingRevisions: executionOptions.shape.existingRevisions,
  operationId: z.string().optional().describe("Optional label echoed back in results")
}).describe("One edit. Set exactly one of replacements, desiredText, delete, comment, restoreText.");

export const inspectFilters = {
  search: z.string().optional().describe("Case-insensitive text to find. Prefer a focused search over reading the whole document."),
  around: z.number().int().min(0).max(20).optional().describe("Paragraphs of context to include around each match"),
  indexes: z.array(z.number().int().positive()).optional(),
  range: z.object({
    start: z.number().int().positive(),
    end: z.number().int().positive()
  }).optional(),
  revisedOnly: z.boolean().optional().describe("Only paragraphs that contain tracked changes"),
  inTable: z.boolean().optional(),
  skipEmpty: z.boolean().optional().describe("Default true"),
  revisionView: z.enum(["accepted", "rejected", "current"]).optional().describe(
    "accepted (default) shows text as if all changes were accepted. rejected shows the original, including text another reviewer deleted: use it to find text to restore."
  ),
  limit: z.number().int().positive().optional().describe("Matches per page (default 100)"),
  after: z.number().int().min(0).optional().describe("Paging cursor: selection.nextAfter from the previous page"),
  excerptLength: z.number().int().positive().optional(),
  compact: z.boolean().optional().describe("Default true. false returns the library's full paragraph records, including paragraphId and fingerprint.")
};
