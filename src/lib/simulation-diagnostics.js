import { mkdir, readFile } from "node:fs/promises";
import { basename, dirname, join } from "node:path";
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { createToolHandlers } from "../server.js";

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function createRedline(target, modified) {
  return {
    type: "redline",
    target,
    modified
  };
}

function createComment(target, textToComment, commentContent) {
  return {
    type: "comment",
    target,
    textToComment,
    commentContent
  };
}

const OPERATIONS = {
  nonExclusiveRedline: createRedline(
    "Non-exclusive: The Supplier’s provision of the Services to the Client is non-exclusive. Nothing in the Agreement prevents the Supplier from providing the Services to any other person.",
    "Non-exclusive: The Supplier’s provision of the Services to the Client is non-exclusive. Nothing in the Agreement prevents the Supplier from providing the Services to any other person, including competitors of the Client."
  ),
  noticesComment: createComment(
    "Notices: A notice given by a party under the Agreement must be delivered to the other party via email using the email address set out in the Key Details or otherwise notified by the other party for this purpose. If the notice is a notice of termination, a copy of that email must be immediately delivered (by hand or courier) to the Chief Executive or equivalent officer of the other party at the other party’s last known physical address.",
    "Notices:",
    "Diagnostic anchor comment. This sits well away from the substantive redlines so we can isolate comment-related package issues."
  ),
  availabilityRedline: createRedline(
    "[Subject to clause 2.3b, t][T]he Supplier will use reasonable efforts to ensure the SaaS Service is available [during normal business hours in New Zealand/on a 24/7 basis]. However, it is possible that on occasion the SaaS Service may be unavailable to permit maintenance or other development activity to take place, or in the event of Force Majeure. The Supplier will use reasonable efforts to publish on the Website [and/or notify the Client by email] advance details of any unavailability.",
    "The Supplier will use reasonable efforts to ensure the SaaS Service is available at least 99.0% of the time, measured monthly, excluding scheduled maintenance, Force Majeure, and outages caused by third party services or infrastructure. This availability level is appropriate for the retrieval of historical, non-critical, non-sensitive documents. The Supplier will use reasonable efforts to publish on the Website and/or notify the Client by email in advance of any scheduled maintenance or expected unavailability."
  ),
  availabilityComment: createComment(
    "Availability: [User note: The clause below places a relatively “low” service level standard on the Supplier. Clients may expect more service levels, e.g. availability (uptime/downtime), response times, fix times, etc., particularly if the Fees are high or the SaaS Service is critical to the Client’s operations.]",
    "Availability:",
    "Sets a concrete but modest uptime target suited to historical document retrieval, while preserving carve-outs for maintenance and third-party outages."
  ),
  backupsRedline: createRedline(
    "Backups of Data: While the Supplier will take standard industry measures to back up all Data stored using the Services, the Client agrees to keep a separate back-up copy of all Data uploaded by it onto the SaaS Service.",
    "Backups of Data: While the Supplier will maintain commercially reasonable backups of Data stored using the Services, the Client is responsible for keeping its own archival copies of all Data uploaded by it onto the SaaS Service, reflecting the low-criticality nature of the retrieval service."
  ),
  backupsComment: createComment(
    "DATA",
    "DATA",
    "Clarifies that supplier backups are operational safeguards, not the client’s archival solution, which is consistent with a low-criticality retrieval service."
  ),
  ipIndemnityComment: createComment(
    "Third party Intellectual Property Rights indemnity: [User note: This is the only indemnity provided by the Supplier. The indemnity protects the Client against claims by third parties that the SaaS Service infringes those third parties’ Intellectual Property Rights. While providing an indemnity is optional, most customers will insist on an indemnity of this nature. The scope of the indemnity will depend on the negotiating power of the parties.]",
    "Third party Intellectual Property Rights indemnity:",
    "Adds the usual carve-outs so the supplier is not indemnifying claims caused by client data, unauthorized changes, or unapproved combinations."
  ),
  ipUseRedline: createRedline(
    "use of the SaaS Service in a manner or for a purpose not reasonably contemplated by the Agreement or otherwise not authorised in writing by the Supplier; or",
    "use of the SaaS Service in a manner or for a purpose not reasonably contemplated by the Agreement or the Documentation, or otherwise not authorised in writing by the Supplier; or"
  ),
  ipDataRedline: createRedline(
    "any third party data or any Data.",
    "any third party data, any Data, or any modification, combination or configuration made by or on behalf of the Client without the Supplier’s written approval."
  ),
  liabilityRedline: createRedline(
    "Maximum liability: The maximum aggregate liability of the Supplier under or in connection with the Agreement or relating to the Services, whether in contract, tort (including negligence), breach of statutory duty or otherwise, must not in any Year exceed [insert cap e.g. an amount equal to the Fees paid by the Client under the Agreement in the previous Year (which in the first Year is deemed to be the total Fees paid by the Client from the Start Date to the date of the first event giving rise to liability)]. The cap in this clause 9.1 includes the cap set out in clause 8.2a.",
    "Maximum liability: The maximum aggregate liability of the Supplier under or in connection with the Agreement or relating to the Services, whether in contract, tort (including negligence), breach of statutory duty or otherwise, must not in any Year exceed [insert cap e.g. an amount equal to the Fees paid by the Client under the Agreement in the previous Year (which in the first Year is deemed to be the total Fees paid by the Client from the Start Date to the date of the first event giving rise to liability)], provided that the cap will not be less than NZD 5,000. The cap in this clause 9.1 includes the cap set out in clause 8.2a."
  ),
  liabilityComment: createComment(
    "LIABILITY",
    "LIABILITY",
    "Adds a NZD 5,000 floor so the liability cap remains commercially meaningful even if the fee base is low."
  ),
  renewalRedline: createRedline(
    "where no End Date is set out in the Key Details, continues for successive terms of [insert term e.g. 12 months] from the Start Date unless a party gives [insert period, e.g. 60 days’] notice that the Agreement will terminate on the expiry of the then-current term.",
    "where no End Date is set out in the Key Details, continues on a month-to-month basis after the Start Date unless a party gives at least 30 days’ notice that the Agreement will terminate."
  ),
  renewalComment: createComment(
    "TERM, TERMINATION AND SUSPENSION",
    "TERM, TERMINATION AND SUSPENSION",
    "Moves the default renewal to month-to-month with 30 days' notice, which reduces lock-in for lower-urgency users."
  ),
  dataAccessRedline: createRedline(
    "At any time prior to one month after the date of termination or expiry, the Client may request:",
    "At any time prior to 60 days after the date of termination or expiry, the Client may request:"
  ),
  dataAccessComment: createComment(
    "Consequences of termination or expiry:",
    "Consequences of termination or expiry:",
    "Extends the post-termination retrieval window to 60 days so the client has a more practical exit period."
  ),
  suspensionRedline: createRedline(
    "The Supplier must notify the Client where it restricts or suspends the Client’s access, or deletes, edits or removes Data, under clause 10.6.",
    "The Supplier must, where legally permitted and reasonably practicable, notify the Client before or promptly after it restricts or suspends the Client’s access, or deletes, edits or removes Data, under clause 10.6, and must reinstate access as soon as reasonably practicable after the relevant issue is resolved."
  ),
  suspensionComment: createComment(
    "Process:",
    "Process:",
    "Adds a practical notice obligation and a reinstatement expectation once the triggering issue has been addressed."
  )
};

export function getSaasAgreementDiagnosticVariants() {
  return [
    {
      id: "01-redline-only",
      description: "One simple redline only.",
      operations: [OPERATIONS.nonExclusiveRedline]
    },
    {
      id: "02-comment-only",
      description: "One simple comment only.",
      operations: [OPERATIONS.noticesComment]
    },
    {
      id: "03-minimal-mixed",
      description: "One redline and one comment in separate areas.",
      operations: [OPERATIONS.nonExclusiveRedline, OPERATIONS.noticesComment]
    },
    {
      id: "04-redlines-core",
      description: "Substantive commercial redlines without comments.",
      operations: [
        OPERATIONS.availabilityRedline,
        OPERATIONS.backupsRedline,
        OPERATIONS.ipUseRedline,
        OPERATIONS.ipDataRedline,
        OPERATIONS.liabilityRedline,
        OPERATIONS.renewalRedline,
        OPERATIONS.dataAccessRedline,
        OPERATIONS.suspensionRedline
      ]
    },
    {
      id: "05-comments-only",
      description: "Seven comments only, anchored away from revised text.",
      operations: [
        OPERATIONS.availabilityComment,
        OPERATIONS.backupsComment,
        OPERATIONS.ipIndemnityComment,
        OPERATIONS.liabilityComment,
        OPERATIONS.renewalComment,
        OPERATIONS.dataAccessComment,
        OPERATIONS.suspensionComment
      ]
    },
    {
      id: "06-mixed-subset",
      description: "Two redlines with nearby-but-separate comments.",
      operations: [
        OPERATIONS.availabilityRedline,
        OPERATIONS.liabilityRedline,
        OPERATIONS.availabilityComment,
        OPERATIONS.liabilityComment
      ]
    },
    {
      id: "07-full-nonoverlap-set",
      description: "Full requested change set with comments anchored to separate stable text.",
      operations: [
        OPERATIONS.availabilityRedline,
        OPERATIONS.backupsRedline,
        OPERATIONS.ipUseRedline,
        OPERATIONS.ipDataRedline,
        OPERATIONS.liabilityRedline,
        OPERATIONS.renewalRedline,
        OPERATIONS.dataAccessRedline,
        OPERATIONS.suspensionRedline,
        OPERATIONS.availabilityComment,
        OPERATIONS.backupsComment,
        OPERATIONS.ipIndemnityComment,
        OPERATIONS.liabilityComment,
        OPERATIONS.renewalComment,
        OPERATIONS.dataAccessComment,
        OPERATIONS.suspensionComment
      ]
    }
  ];
}

export function getSaasAgreementSingleRedlineVariants() {
  return [
    {
      id: "11-availability-redline",
      description: "Availability and uptime redline only.",
      operations: [OPERATIONS.availabilityRedline]
    },
    {
      id: "12-backups-redline",
      description: "Backup responsibilities redline only.",
      operations: [OPERATIONS.backupsRedline]
    },
    {
      id: "13-ip-use-redline",
      description: "IP indemnity unauthorized-use carve-out only.",
      operations: [OPERATIONS.ipUseRedline]
    },
    {
      id: "14-ip-data-redline",
      description: "IP indemnity client-data/modification carve-out only.",
      operations: [OPERATIONS.ipDataRedline]
    },
    {
      id: "15-liability-redline",
      description: "Liability floor redline only.",
      operations: [OPERATIONS.liabilityRedline]
    },
    {
      id: "16-renewal-redline",
      description: "Month-to-month renewal redline only.",
      operations: [OPERATIONS.renewalRedline]
    },
    {
      id: "17-data-access-redline",
      description: "Post-termination data access window redline only.",
      operations: [OPERATIONS.dataAccessRedline]
    },
    {
      id: "18-suspension-redline",
      description: "Suspension notice and reinstatement redline only.",
      operations: [OPERATIONS.suspensionRedline]
    }
  ];
}

function parseXml(xml) {
  return new DOMParser().parseFromString(xml, "text/xml");
}

function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc);
}

function normalizeText(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function paragraphText(paragraph) {
  return Array.from(paragraph.getElementsByTagNameNS(NS_W, "t"))
    .map(node => node.textContent || "")
    .join("");
}

function findParagraphByText(documentDoc, targetText) {
  const normalizedTarget = normalizeText(targetText);
  const paragraphs = Array.from(documentDoc.getElementsByTagNameNS(NS_W, "p"));
  return paragraphs.find(paragraph => normalizeText(paragraphText(paragraph)).includes(normalizedTarget)) || null;
}

function removeProofErrNodes(paragraph) {
  for (const node of Array.from(paragraph.getElementsByTagNameNS(NS_W, "proofErr"))) {
    node.parentNode?.removeChild(node);
  }
}

function directFieldCharType(node) {
  if (!node || node.nodeType !== 1 || node.namespaceURI !== NS_W || node.localName !== "r") {
    return "";
  }

  const fldChar = Array.from(node.childNodes || []).find(
    child => child.nodeType === 1 && child.namespaceURI === NS_W && child.localName === "fldChar"
  );

  return fldChar?.getAttribute("w:fldCharType") || fldChar?.getAttribute("fldCharType") || "";
}

function hasInstrText(node) {
  if (!node || node.nodeType !== 1) return false;
  return Array.from(node.childNodes || []).some(
    child => child.nodeType === 1 && child.namespaceURI === NS_W && child.localName === "instrText"
  );
}

function unlinkFieldsInParagraph(paragraph) {
  const children = Array.from(paragraph.childNodes || []);
  let inField = false;
  let seenSeparate = false;

  for (const child of children) {
    if (child.nodeType !== 1) continue;

    const fieldType = directFieldCharType(child);
    if (fieldType === "begin") {
      inField = true;
      seenSeparate = false;
      child.parentNode?.removeChild(child);
      continue;
    }

    if (!inField) {
      continue;
    }

    if (fieldType === "separate") {
      seenSeparate = true;
      child.parentNode?.removeChild(child);
      continue;
    }

    if (fieldType === "end") {
      inField = false;
      seenSeparate = false;
      child.parentNode?.removeChild(child);
      continue;
    }

    if (!seenSeparate || hasInstrText(child)) {
      child.parentNode?.removeChild(child);
    }
  }
}

function preprocessDocumentXmlForVariant(documentXml, variant) {
  if (variant?.preprocess !== "strip_fields_and_proofing") {
    return documentXml;
  }

  const documentDoc = parseXml(documentXml);
  const targetOperation = Array.isArray(variant.operations)
    ? variant.operations.find(operation => operation?.type === "redline")
    : null;

  if (!targetOperation?.target) {
    return documentXml;
  }

  const paragraph = findParagraphByText(documentDoc, targetOperation.target);
  if (!paragraph) {
    return documentXml;
  }

  removeProofErrNodes(paragraph);
  unlinkFieldsInParagraph(paragraph);
  return serializeXml(documentDoc);
}

export function getSaasAgreementFieldSafeVariants() {
  return [
    {
      id: "21-availability-redline-fieldsafe",
      description: "Availability redline after stripping proofing and field markup from the target paragraph.",
      preprocess: "strip_fields_and_proofing",
      operations: [OPERATIONS.availabilityRedline]
    },
    {
      id: "22-liability-redline-fieldsafe",
      description: "Liability redline after stripping proofing and field markup from the target paragraph.",
      preprocess: "strip_fields_and_proofing",
      operations: [OPERATIONS.liabilityRedline]
    },
    {
      id: "23-suspension-redline-fieldsafe",
      description: "Suspension redline after stripping proofing and field markup from the target paragraph.",
      preprocess: "strip_fields_and_proofing",
      operations: [OPERATIONS.suspensionRedline]
    }
  ];
}

function getWordId(node) {
  return node.getAttribute("w:id") || node.getAttribute("id") || "";
}

function nearestWordAncestor(node, localName) {
  let current = node;
  while (current) {
    if (current.nodeType === 1 && current.namespaceURI === NS_W && current.localName === localName) {
      return current;
    }
    current = current.parentNode;
  }
  return null;
}

function getParserErrorCount(doc) {
  return doc.getElementsByTagName("parsererror").length;
}

function countCrossingCommentRanges(documentDoc) {
  const starts = Array.from(documentDoc.getElementsByTagNameNS(NS_W, "commentRangeStart"));
  const ends = Array.from(documentDoc.getElementsByTagNameNS(NS_W, "commentRangeEnd"));
  const refs = Array.from(documentDoc.getElementsByTagNameNS(NS_W, "commentReference"));
  let count = 0;

  for (const start of starts) {
    const id = getWordId(start);
    const end = ends.find(node => getWordId(node) === id);
    const ref = refs.find(node => getWordId(node) === id);

    if (!end || !ref) {
      count += 1;
      continue;
    }

    const startIns = nearestWordAncestor(start, "ins");
    const endIns = nearestWordAncestor(end, "ins");
    const refIns = nearestWordAncestor(ref, "ins");
    const startDel = nearestWordAncestor(start, "del");
    const endDel = nearestWordAncestor(end, "del");
    const refDel = nearestWordAncestor(ref, "del");

    if (!(startIns === endIns && endIns === refIns)) {
      count += 1;
      continue;
    }

    if (!(startDel === endDel && endDel === refDel)) {
      count += 1;
    }
  }

  return count;
}

export async function analyzeDocxFile(filePath) {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const commentsFile = zip.file("word/comments.xml");
  const commentsXml = commentsFile ? await commentsFile.async("string") : "";
  const documentDoc = parseXml(documentXml);
  const commentsDoc = commentsXml ? parseXml(commentsXml) : null;

  return {
    fileName: basename(filePath),
    trackedChangesPresent: /<w:(ins|del)\b/.test(documentXml),
    commentCount: commentsDoc ? commentsDoc.getElementsByTagNameNS(NS_W, "comment").length : 0,
    crossingCommentRanges: countCrossingCommentRanges(documentDoc),
    parserErrors: getParserErrorCount(documentDoc) + (commentsDoc ? getParserErrorCount(commentsDoc) : 0),
    rooInsertions: Array.from(documentDoc.getElementsByTagNameNS(NS_W, "ins"))
      .filter(node => (node.getAttribute("w:author") || node.getAttribute("author")) === "Roo")
      .length,
    rooDeletions: Array.from(documentDoc.getElementsByTagNameNS(NS_W, "del"))
      .filter(node => (node.getAttribute("w:author") || node.getAttribute("author")) === "Roo")
      .length,
    rooComments: commentsDoc
      ? Array.from(commentsDoc.getElementsByTagNameNS(NS_W, "comment"))
        .filter(node => (node.getAttribute("w:author") || node.getAttribute("author")) === "Roo")
        .length
      : 0
  };
}

export async function generateDiagnosticVariant({
  sourcePath,
  outputPath,
  variant,
  author = "Roo",
  handlers = createToolHandlers(),
  strategy = "sequential"
}) {
  await mkdir(dirname(outputPath), { recursive: true });
  const opened = await handlers.openDocx({ path: sourcePath });
  const originalDocument = await handlers.readPart({
    handle: opened.handle,
    partPath: "word/document.xml"
  });

  if (originalDocument.kind !== "xml") {
    throw new Error("word/document.xml is not readable as XML");
  }

  const preprocessedDocumentXml = preprocessDocumentXmlForVariant(originalDocument.text, variant);
  if (preprocessedDocumentXml !== originalDocument.text) {
    await handlers.writePart({
      handle: opened.handle,
      partPath: "word/document.xml",
      xml: preprocessedDocumentXml
    });
  }

  if (strategy === "batch") {
    await handlers.applyOperations({
      handle: opened.handle,
      operations: variant.operations,
      author
    });
  } else {
    for (const operation of variant.operations) {
      await handlers.applyOperation({
        handle: opened.handle,
        operation,
        author
      });
    }
  }

  await handlers.saveDocx({
    handle: opened.handle,
    outputPath
  });

  return analyzeDocxFile(outputPath);
}

export async function generateDiagnosticVariants({
  sourcePath,
  outputDir,
  author = "Roo",
  handlers = createToolHandlers(),
  strategy = "sequential",
  variants = getSaasAgreementDiagnosticVariants()
}) {
  await mkdir(outputDir, { recursive: true });
  const results = [];

  for (const variant of variants) {
    const outputPath = join(outputDir, `${variant.id}.docx`);
    const analysis = await generateDiagnosticVariant({
      sourcePath,
      outputPath,
      variant,
      author,
      handlers,
      strategy
    });
    results.push({
      id: variant.id,
      description: variant.description,
      outputPath,
      ...analysis
    });
  }

  return results;
}
