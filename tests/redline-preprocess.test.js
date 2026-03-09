import test from "node:test";
import assert from "node:assert/strict";
import { preprocessDocumentXmlForOperation } from "../src/lib/redline-preprocess.js";

const DOCUMENT_XML = [
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
].join("");

test("preprocess removes field instructions and proofing markers from redline target paragraph", () => {
  const result = preprocessDocumentXmlForOperation(DOCUMENT_XML, {
    type: "redline",
    target: "The Supplier must notify the Client where it restricts 10.6."
  });

  assert.doesNotMatch(result, /instrText/);
  assert.doesNotMatch(result, /fldChar/);
  assert.doesNotMatch(result, /proofErr/);
  assert.match(result, />10\.6</);
});

test("preprocess leaves non-redline operations unchanged", () => {
  const result = preprocessDocumentXmlForOperation(DOCUMENT_XML, {
    type: "comment",
    target: "The Supplier must notify the Client where it restricts 10.6."
  });

  assert.equal(result, DOCUMENT_XML);
});
