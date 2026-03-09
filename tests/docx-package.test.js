import test from "node:test";
import assert from "node:assert/strict";
import JSZip from "jszip";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import {
  createDocxSessionFromPath,
  readPart,
  writeXmlPart,
  saveDocxSession
} from "../src/lib/docx-package.js";

test("opens a docx, reads document xml, writes xml, and saves atomically", async () => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>');
  zip.file(
    "word/document.xml",
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body /></w:document>'
  );
  zip.file("word/media/image1.png", Buffer.from([1, 2, 3, 4]));

  const dir = await mkdtemp(join(tmpdir(), "docx-redline-mcp-"));
  const filePath = join(dir, "sample.docx");
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));

  const session = await createDocxSessionFromPath(filePath);
  const documentPart = await readPart(session.zip, "word/document.xml");
  assert.equal(documentPart.kind, "xml");
  assert.match(documentPart.text, /w:document/);

  const imagePart = await readPart(session.zip, "word/media/image1.png");
  assert.equal(imagePart.kind, "binary");
  assert.equal(imagePart.byteLength, 4);

  writeXmlPart(
    session.zip,
    "word/document.xml",
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p /></w:body></w:document>'
  );
  await saveDocxSession(session, filePath);

  const saved = await readFile(filePath);
  assert.ok(saved.length > 0);

  const reopened = await createDocxSessionFromPath(filePath);
  const updatedPart = await readPart(reopened.zip, "word/document.xml");
  assert.match(updatedPart.text, /<w:p/);
});
