import { resolve } from "node:path";
import {
  generateDiagnosticVariants,
  getSaasAgreementDiagnosticVariants,
  getSaasAgreementFieldSafeVariants,
  getSaasAgreementSingleRedlineVariants
} from "../src/lib/simulation-diagnostics.js";

const sourcePath = resolve("simulation/saas-agreement-2023-02-16.docx");
const outputDir = resolve("simulation/diagnostics");
const batchOutputDir = resolve("simulation/diagnostics-batch");
const singleRedlineOutputDir = resolve("simulation/diagnostics-single-redlines");
const fieldSafeOutputDir = resolve("simulation/diagnostics-fieldsafe");
const variants = getSaasAgreementDiagnosticVariants();
const failingSequentialIds = new Set(["04-redlines-core", "06-mixed-subset", "07-full-nonoverlap-set"]);

const sequentialResults = await generateDiagnosticVariants({
  sourcePath,
  outputDir,
  author: "Roo"
});

const batchResults = await generateDiagnosticVariants({
  sourcePath,
  outputDir: batchOutputDir,
  author: "Roo",
  strategy: "batch",
  variants: variants.filter(variant => failingSequentialIds.has(variant.id))
});

const singleRedlineResults = await generateDiagnosticVariants({
  sourcePath,
  outputDir: singleRedlineOutputDir,
  author: "Roo",
  variants: getSaasAgreementSingleRedlineVariants()
});

const fieldSafeResults = await generateDiagnosticVariants({
  sourcePath,
  outputDir: fieldSafeOutputDir,
  author: "Roo",
  variants: getSaasAgreementFieldSafeVariants()
});

console.log(JSON.stringify({
  sourcePath,
  outputDir,
  batchOutputDir,
  singleRedlineOutputDir,
  fieldSafeOutputDir,
  sequentialVariants: sequentialResults,
  batchVariants: batchResults,
  singleRedlineVariants: singleRedlineResults,
  fieldSafeVariants: fieldSafeResults
}, null, 2));
