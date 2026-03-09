import test from "node:test";
import assert from "node:assert/strict";
import { access, mkdtemp } from "node:fs/promises";
import { join, resolve } from "node:path";
import { tmpdir } from "node:os";
import {
  generateDiagnosticVariants,
  getSaasAgreementDiagnosticVariants,
  getSaasAgreementFieldSafeVariants,
  getSaasAgreementSingleRedlineVariants
} from "../src/lib/simulation-diagnostics.js";

test("generates diagnostic variants from the clean SaaS agreement", async () => {
  const sourcePath = resolve("simulation/saas-agreement-2023-02-16.docx");
  await access(sourcePath);

  const outputDir = await mkdtemp(join(tmpdir(), "docx-redline-diagnostics-"));
  const results = await generateDiagnosticVariants({
    sourcePath,
    outputDir,
    author: "Roo"
  });

  assert.equal(results.length, 7);

  for (const result of results) {
    await access(result.outputPath);
    assert.equal(result.parserErrors, 0, `expected no parser errors for ${result.id}`);
    assert.equal(result.crossingCommentRanges, 0, `expected no crossing comment ranges for ${result.id}`);
  }

  const minimalMixed = results.find(result => result.id === "03-minimal-mixed");
  assert.equal(minimalMixed?.commentCount, 1);
  assert.equal(minimalMixed?.rooComments, 1);

  const commentsOnly = results.find(result => result.id === "05-comments-only");
  assert.equal(commentsOnly?.trackedChangesPresent, false);
  assert.equal(commentsOnly?.commentCount, 7);

  const fullSet = results.find(result => result.id === "07-full-nonoverlap-set");
  assert.equal(fullSet?.trackedChangesPresent, true);
  assert.equal(fullSet?.commentCount, 7);
  assert.ok((fullSet?.rooInsertions || 0) > 0);
  assert.ok((fullSet?.rooDeletions || 0) > 0);
});

test("batch diagnostic variants generate from the clean SaaS agreement", async () => {
  const sourcePath = resolve("simulation/saas-agreement-2023-02-16.docx");
  await access(sourcePath);

  const outputDir = await mkdtemp(join(tmpdir(), "docx-redline-diagnostics-batch-"));
  const batchVariantIds = new Set(["04-redlines-core", "06-mixed-subset", "07-full-nonoverlap-set"]);
  const results = await generateDiagnosticVariants({
    sourcePath,
    outputDir,
    author: "Roo",
    strategy: "batch",
    variants: getSaasAgreementDiagnosticVariants().filter(variant => batchVariantIds.has(variant.id))
  });

  assert.equal(results.length, 3);

  for (const result of results) {
    await access(result.outputPath);
    assert.equal(result.parserErrors, 0, `expected no parser errors for ${result.id}`);
    assert.equal(result.crossingCommentRanges, 0, `expected no crossing comment ranges for ${result.id}`);
  }

  const fullSet = results.find(result => result.id === "07-full-nonoverlap-set");
  assert.equal(fullSet?.commentCount, 7);
  assert.ok((fullSet?.rooInsertions || 0) > 0);
  assert.ok((fullSet?.rooDeletions || 0) > 0);
});

test("single-redline diagnostic variants generate from the clean SaaS agreement", async () => {
  const sourcePath = resolve("simulation/saas-agreement-2023-02-16.docx");
  await access(sourcePath);

  const outputDir = await mkdtemp(join(tmpdir(), "docx-redline-diagnostics-single-"));
  const results = await generateDiagnosticVariants({
    sourcePath,
    outputDir,
    author: "Roo",
    variants: getSaasAgreementSingleRedlineVariants()
  });

  assert.equal(results.length, 8);

  for (const result of results) {
    await access(result.outputPath);
    assert.equal(result.parserErrors, 0, `expected no parser errors for ${result.id}`);
    assert.equal(result.crossingCommentRanges, 0, `expected no crossing comment ranges for ${result.id}`);
    assert.equal(result.commentCount, 0, `expected no comments for ${result.id}`);
    assert.equal(result.trackedChangesPresent, true, `expected tracked changes for ${result.id}`);
  }
});

test("field-safe diagnostic variants generate from the clean SaaS agreement", async () => {
  const sourcePath = resolve("simulation/saas-agreement-2023-02-16.docx");
  await access(sourcePath);

  const outputDir = await mkdtemp(join(tmpdir(), "docx-redline-diagnostics-fieldsafe-"));
  const results = await generateDiagnosticVariants({
    sourcePath,
    outputDir,
    author: "Roo",
    variants: getSaasAgreementFieldSafeVariants()
  });

  assert.equal(results.length, 3);

  for (const result of results) {
    await access(result.outputPath);
    assert.equal(result.parserErrors, 0, `expected no parser errors for ${result.id}`);
    assert.equal(result.crossingCommentRanges, 0, `expected no crossing comment ranges for ${result.id}`);
    assert.equal(result.commentCount, 0, `expected no comments for ${result.id}`);
    assert.equal(result.trackedChangesPresent, true, `expected tracked changes for ${result.id}`);
  }
});
