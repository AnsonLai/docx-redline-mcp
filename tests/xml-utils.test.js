import test from "node:test";
import assert from "node:assert/strict";
import {
  assertXmlString,
  isXmlPartPath,
  isWritableXmlPartPath
} from "../src/lib/xml-utils.js";

test("accepts writable OOXML parts and rejects malformed xml", () => {
  assert.equal(isXmlPartPath("word/document.xml"), true);
  assert.equal(isWritableXmlPartPath("word/document.xml"), true);
  assert.equal(isWritableXmlPartPath("word/header1.xml"), true);
  assert.equal(isWritableXmlPartPath("word/media/image1.png"), false);

  assert.doesNotThrow(() =>
    assertXmlString('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:document>')
  );
  assert.throws(() => assertXmlString("<w:document>"), /Malformed XML/);
});
