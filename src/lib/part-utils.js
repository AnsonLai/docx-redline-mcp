import { isXmlPartPath } from "./xml-utils.js";

function inferContentType(partPath) {
  const value = String(partPath || "").toLowerCase();

  if (value.endsWith(".xml") || value.endsWith(".rels")) {
    return "application/xml";
  }
  if (value.endsWith(".png")) {
    return "image/png";
  }
  if (value.endsWith(".jpg") || value.endsWith(".jpeg")) {
    return "image/jpeg";
  }
  if (value.endsWith(".gif")) {
    return "image/gif";
  }

  return "application/octet-stream";
}

export function listPartEntries(zip) {
  return Object.values(zip?.files || {})
    .filter(file => !file.dir)
    .map(file => ({
      path: file.name,
      isXml: isXmlPartPath(file.name),
      contentType: inferContentType(file.name)
    }))
    .sort((left, right) => left.path.localeCompare(right.path));
}
