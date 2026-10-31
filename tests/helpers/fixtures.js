import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function escapeXml(text) {
  return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function paragraphXml(paragraph) {
  if (typeof paragraph === "object" && paragraph.xml) return paragraph.xml;
  return `<w:p><w:r><w:t xml:space="preserve">${escapeXml(paragraph)}</w:t></w:r></w:p>`;
}

/**
 * Builds a minimal but valid .docx package.
 *
 * `paragraphs` are strings (or `{ xml }` for hand-written paragraphs). `header` adds a default
 * header part referenced from the body section.
 */
export async function buildDocxBuffer({ paragraphs = ["Original text"], header = null, bodyXml = null } = {}) {
  const zip = new JSZip();
  const headerOverride = header
    ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    : "";
  zip.file(
    "[Content_Types].xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
      headerOverride,
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
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      header
        ? '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
        : "",
      "</Relationships>"
    ].join("")
  );
  const sectPr = header
    ? '<w:sectPr><w:headerReference w:type="default" r:id="rId10"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
    : '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
  zip.file(
    "word/document.xml",
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">`,
      "<w:body>",
      bodyXml ?? paragraphs.map(paragraphXml).join(""),
      sectPr,
      "</w:body>",
      "</w:document>"
    ].join("")
  );
  if (header) {
    zip.file(
      "word/header1.xml",
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        `<w:hdr xmlns:w="${NS_W}">`,
        paragraphXml(header),
        "</w:hdr>"
      ].join("")
    );
  }
  return zip.generateAsync({ type: "nodebuffer" });
}

export async function writeDocxFixture(name, options) {
  const directoryPath = await mkdtemp(join(tmpdir(), "docx-redline-mcp-"));
  const filePath = join(directoryPath, name);
  await writeFile(filePath, await buildDocxBuffer(options));
  return filePath;
}
