import {
  acceptTrackedChangesInOoxml,
  configureXmlProvider,
  createDynamicNumberingIdState,
  deleteCommentsByAuthorInOoxml,
  ensureCommentsArtifactsInZip,
  ensureNumberingArtifactsInZip,
  rejectTrackedChangesInOoxml
} from "@ansonlai/docx-redline-js";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import {
  applyOperationToDocumentXml
} from "@ansonlai/docx-redline-js/services/standalone-operation-runner.js";
import { getEmptyCommentsXml, normalizeDocumentXml } from "./package-validation.js";
import { preprocessDocumentXmlForOperation } from "./redline-preprocess.js";

configureXmlProvider({ DOMParser, XMLSerializer });

async function createRuntimeContext(zip) {
  const existingNumberingXml = zip?.file("word/numbering.xml")
    ? await zip.file("word/numbering.xml").async("string")
    : "";

  return {
    numberingIdState: createDynamicNumberingIdState(existingNumberingXml, {
      minId: 1,
      maxPreferred: 32767
    }),
    listFallbackSharedNumIdByKey: new Map(),
    listFallbackSequenceState: { explicitByNumberingKey: new Map() },
    tableStructuralRedlineKeys: new Set()
  };
}

function getMaintenanceAction(operation) {
  const type = String(operation?.type || "").toLowerCase();
  if (type === "accept_all_tracked_changes") {
    return () => acceptTrackedChangesInOoxml(operation.documentXml, { allAuthors: true });
  }
  if (type === "reject_all_tracked_changes") {
    return () => rejectTrackedChangesInOoxml(operation.documentXml, { allAuthors: true });
  }
  if (type === "delete_all_comments") {
    return () => deleteCommentsByAuthorInOoxml(operation.documentXml, { allAuthors: true });
  }
  return null;
}

export async function applyStructuredOperation(documentXml, operation, author = "AI Agent", options = {}) {
  const preparedDocumentXml = preprocessDocumentXmlForOperation(documentXml, operation);
  const maintenanceAction = getMaintenanceAction({
    ...operation,
    documentXml: preparedDocumentXml
  });

  if (maintenanceAction) {
    const result = await maintenanceAction();
    return {
      documentXml: options.normalizeDocumentXml === false
        ? String(result?.oxml || preparedDocumentXml)
        : normalizeDocumentXml(result?.oxml || preparedDocumentXml),
      hasChanges: Boolean(result?.hasChanges),
      commentsXml: String(operation?.type || "").toLowerCase() === "delete_all_comments"
        ? getEmptyCommentsXml()
        : null,
      replaceCommentsPart: String(operation?.type || "").toLowerCase() === "delete_all_comments"
    };
  }

  const result = await applyOperationToDocumentXml(preparedDocumentXml, operation, author, options.runtimeContext, {
    generateRedlines: true
  });
  return {
    ...result,
    documentXml: options.normalizeDocumentXml === false
      ? String(result?.documentXml || preparedDocumentXml)
      : normalizeDocumentXml(result?.documentXml || preparedDocumentXml)
  };
}

export async function applyStructuredOperations(zip, documentXml, operations, author = "AI Agent") {
  const runtimeContext = await createRuntimeContext(zip);
  let nextDocumentXml = documentXml;
  let hasChanges = false;
  let replaceCommentsPart = false;
  let commentsXml = null;
  const numberingXml = [];

  for (const operation of operations) {
    const result = await applyStructuredOperation(nextDocumentXml, operation, author, {
      runtimeContext,
      normalizeDocumentXml: false
    });

    nextDocumentXml = result.documentXml;
    hasChanges = hasChanges || Boolean(result?.hasChanges);

    if (result?.numberingXml) {
      numberingXml.push(result.numberingXml);
    }

    if (result?.replaceCommentsPart === true && typeof result?.commentsXml === "string") {
      replaceCommentsPart = true;
      commentsXml = result.commentsXml;
    } else if (result?.commentsXml) {
      if (Array.isArray(commentsXml)) {
        commentsXml.push(result.commentsXml);
      } else if (typeof commentsXml === "string") {
        commentsXml = [commentsXml, result.commentsXml];
      } else {
        commentsXml = [result.commentsXml];
      }
    }
  }

  return {
    documentXml: normalizeDocumentXml(nextDocumentXml),
    hasChanges,
    numberingXml,
    commentsXml,
    replaceCommentsPart
  };
}

export async function mergeOperationArtifacts(zip, result) {
  if (Array.isArray(result?.numberingXml)) {
    for (const numberingXml of result.numberingXml) {
      await ensureNumberingArtifactsInZip(zip, numberingXml);
    }
  } else if (result?.numberingXml) {
    await ensureNumberingArtifactsInZip(zip, result.numberingXml);
  }

  if (result?.replaceCommentsPart === true && typeof result?.commentsXml === "string") {
    zip.file("word/comments.xml", result.commentsXml);
  } else if (Array.isArray(result?.commentsXml)) {
    for (const commentsXml of result.commentsXml) {
      await ensureCommentsArtifactsInZip(zip, commentsXml);
    }
  } else if (result?.commentsXml) {
    await ensureCommentsArtifactsInZip(zip, result.commentsXml);
  }
}
