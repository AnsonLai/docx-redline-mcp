import test from "node:test";
import assert from "node:assert/strict";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import {
  configureXmlProvider
} from "@ansonlai/docx-redline-js";
import { applyStructuredOperation } from "../src/lib/operation-service.js";

configureXmlProvider({ DOMParser, XMLSerializer });

test("applies a redline operation to word/document.xml", async () => {
  const documentXml = [
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    "<w:body>",
    "<w:p><w:r><w:t>Original text</w:t></w:r></w:p>",
    "</w:body>",
    "</w:document>"
  ].join("");

  const result = await applyStructuredOperation(documentXml, {
    type: "redline",
    target: "Original text",
    modified: "Updated text"
  });

  assert.match(result.documentXml, /w:(ins|del)/);
});
