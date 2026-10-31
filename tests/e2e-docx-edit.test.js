import test from "node:test";
import assert from "node:assert/strict";
import JSZip from "jszip";
import { mkdtemp, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { createToolHandlers } from "../src/server.js";
import { createDocxSessionFromPath, readPart } from "../src/lib/docx-package.js";

test("opens, edits, and saves a docx through the tool handlers", async () => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
      "</Types>"
    ].join("")
  );
  zip.file(
    "_rels/.rels",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
      "</Relationships>"
    ].join("")
  );
  zip.file(
    "word/_rels/document.xml.rels",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    ].join("")
  );
  zip.file(
    "word/document.xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body>",
      "<w:p><w:r><w:t>Original text</w:t></w:r></w:p>",
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
      "</w:body>",
      "</w:document>"
    ].join("")
  );

  const directoryPath = await mkdtemp(join(tmpdir(), "docx-redline-mcp-"));
  const filePath = join(directoryPath, "sample.docx");
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));

  const handlers = createToolHandlers();
  const opened = await handlers.openDocx({ path: filePath });
  assert.ok(opened.handle);

  const listed = await handlers.listParts({ handle: opened.handle });
  assert.ok(listed.parts.some(part => part.path === "word/document.xml"));

  const before = await handlers.readPart({
    handle: opened.handle,
    partPath: "word/document.xml"
  });
  assert.equal(before.kind, "xml");
  assert.match(before.text, /Original text/);

  const applied = await handlers.applyOperation({
    handle: opened.handle,
    operation: {
      type: "redline",
      target: "Original text",
      modified: "Updated text"
    }
  });
  assert.equal(applied.dirty, true);

  const saved = await handlers.saveDocx({ handle: opened.handle, inPlace: true });
  assert.equal(saved.savedPath, filePath);
  assert.equal(saved.dirty, false);

  const reopened = await createDocxSessionFromPath(filePath);
  const finalDocument = await readPart(reopened.zip, "word/document.xml");
  assert.match(finalDocument.text, /w:(ins|del)/);
});

test("rejects comment anchors that would cross tracked change boundaries and rolls back", async () => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
      "</Types>"
    ].join("")
  );
  zip.file(
    "_rels/.rels",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
      "</Relationships>"
    ].join("")
  );
  zip.file(
    "word/_rels/document.xml.rels",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    ].join("")
  );
  zip.file(
    "word/document.xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body>",
      "<w:p><w:r><w:t>Original text</w:t></w:r></w:p>",
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
      "</w:body>",
      "</w:document>"
    ].join("")
  );

  const directoryPath = await mkdtemp(join(tmpdir(), "docx-redline-mcp-"));
  const filePath = join(directoryPath, "sample-invalid-comment.docx");
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));

  const handlers = createToolHandlers();
  const opened = await handlers.openDocx({ path: filePath });

  await handlers.applyOperation({
    handle: opened.handle,
    operation: {
      type: "redline",
      target: "Original text",
      modified: "Updated text"
    }
  });

  await assert.rejects(
    handlers.applyOperation({
      handle: opened.handle,
      operation: {
        type: "comment",
        target: "Updated text",
        textToComment: "Updated text",
        commentContent: "Crosses change boundary"
      }
    }),
    /crosses a tracked insertion boundary/
  );

  const documentAfterFailure = await handlers.readPart({
    handle: opened.handle,
    partPath: "word/document.xml"
  });
  assert.doesNotMatch(documentAfterFailure.text, /commentRangeStart|commentReference/);
});

test("delete_all_comments clears the comments part", async () => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
      "</Types>"
    ].join("")
  );
  zip.file(
    "_rels/.rels",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
      "</Relationships>"
    ].join("")
  );
  zip.file(
    "word/_rels/document.xml.rels",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    ].join("")
  );
  zip.file(
    "word/document.xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body>",
      "<w:p><w:r><w:t>Stable paragraph</w:t></w:r></w:p>",
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
      "</w:body>",
      "</w:document>"
    ].join("")
  );

  const directoryPath = await mkdtemp(join(tmpdir(), "docx-redline-mcp-"));
  const filePath = join(directoryPath, "sample-comments.docx");
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));

  const handlers = createToolHandlers();
  const opened = await handlers.openDocx({ path: filePath });

  await handlers.applyOperation({
    handle: opened.handle,
    operation: {
      type: "comment",
      target: "Stable paragraph",
      textToComment: "Stable",
      commentContent: "Test comment"
    }
  });

  await handlers.applyOperation({
    handle: opened.handle,
    operation: {
      type: "delete_all_comments"
    }
  });

  const commentsPart = await handlers.readPart({
    handle: opened.handle,
    partPath: "word/comments.xml"
  });
  assert.equal(commentsPart.kind, "xml");
  assert.doesNotMatch(commentsPart.text, /<w:comment\b/);
});

test("redline preprocessing strips field scaffolding from the target paragraph before applying revisions", async () => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
      "</Types>"
    ].join("")
  );
  zip.file(
    "_rels/.rels",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
      "</Relationships>"
    ].join("")
  );
  zip.file(
    "word/_rels/document.xml.rels",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    ].join("")
  );
  zip.file(
    "word/document.xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      "<w:body>",
      "<w:p>",
      "<w:pPr/>",
      '<w:r><w:t>The Supplier must notify the Client where it restricts </w:t></w:r>',
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
      '<w:r><w:instrText xml:space="preserve"> REF _Ref12345 \\\\w \\\\h </w:instrText></w:r>',
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
      '<w:r><w:t>10.6</w:t></w:r>',
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
      '<w:proofErr w:type="spellStart"/>',
      '<w:r><w:t>.</w:t></w:r>',
      '<w:proofErr w:type="spellEnd"/>',
      "</w:p>",
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
      "</w:body>",
      "</w:document>"
    ].join("")
  );

  const directoryPath = await mkdtemp(join(tmpdir(), "docx-redline-mcp-"));
  const filePath = join(directoryPath, "sample-field-redline.docx");
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));

  const handlers = createToolHandlers();
  const opened = await handlers.openDocx({ path: filePath });

  await handlers.applyOperation({
    handle: opened.handle,
    author: "Roo",
    operation: {
      type: "redline",
      target: "The Supplier must notify the Client where it restricts 10.6.",
      modified: "The Supplier must, where legally permitted and reasonably practicable, notify the Client where it restricts 10.6."
    }
  });

  const documentPart = await handlers.readPart({
    handle: opened.handle,
    partPath: "word/document.xml"
  });

  assert.equal(documentPart.kind, "xml");
  assert.doesNotMatch(documentPart.text, /instrText/);
  assert.doesNotMatch(documentPart.text, /fldChar/);
  assert.doesNotMatch(documentPart.text, /proofErr/);
  assert.match(documentPart.text, /w:(ins|del)/);
  assert.match(documentPart.text, />10\.6</);
});
