import { XMLBuilder, XMLParser } from "fast-xml-parser";
import path from "node:path";
import { ensureArray } from "./object.js";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  removeNSPrefix: false
});

const builder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  suppressBooleanAttributes: false,
  suppressEmptyNode: true,
  format: true,
  indentBy: "  "
});

const xmlDeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function parseXml(xmlText) {
  return parser.parse(xmlText);
}

export function buildXml(xmlObject) {
  return `${xmlDeclaration}${builder.build(xmlObject)}`;
}

export function relsPartPath(partPath) {
  const dir = path.posix.dirname(partPath);
  const file = path.posix.basename(partPath);
  return path.posix.join(dir, "_rels", `${file}.rels`);
}

export function normalizePartPath(value) {
  return String(value || "").replace(/\\/g, "/").replace(/^\/+/, "");
}

export function resolveTargetPath(partPath, target) {
  const normalizedPart = normalizePartPath(partPath);
  const rawTarget = String(target || "").replace(/\\/g, "/");
  const normalizedTarget = normalizePartPath(rawTarget);

  if (rawTarget.startsWith("/")) {
    return normalizePartPath(normalizedTarget.slice(1));
  }

  const baseDir = path.posix.dirname(normalizedPart);
  return normalizePartPath(path.posix.normalize(path.posix.join(baseDir, normalizedTarget)));
}

export function relationshipMap(relsXml) {
  const root = relsXml?.Relationships;
  const rels = ensureArray(root?.Relationship);
  const map = new Map();
  for (const rel of rels) {
    const id = rel?.["@_Id"];
    if (!id) {
      continue;
    }
    map.set(id, {
      id,
      type: rel?.["@_Type"],
      target: rel?.["@_Target"],
      targetMode: rel?.["@_TargetMode"] || "Internal"
    });
  }
  return map;
}

export function ensureZipPath(p) {
  return normalizePartPath(p).replace(/^\/+/, "");
}

export function contentTypeMap(contentTypesXml) {
  const overrides = ensureArray(contentTypesXml?.Types?.Override);
  const defaults = ensureArray(contentTypesXml?.Types?.Default);
  const overrideMap = new Map();
  const defaultMap = new Map();

  for (const entry of overrides) {
    const partName = normalizePartPath(entry?.["@_PartName"] || "").replace(/^\//, "");
    const contentType = entry?.["@_ContentType"];
    if (partName && contentType) {
      overrideMap.set(partName, contentType);
    }
  }

  for (const entry of defaults) {
    const ext = entry?.["@_Extension"];
    const contentType = entry?.["@_ContentType"];
    if (ext && contentType) {
      defaultMap.set(ext.toLowerCase(), contentType);
    }
  }

  return {
    get(partPath) {
      const normalized = normalizePartPath(partPath);
      if (overrideMap.has(normalized)) {
        return overrideMap.get(normalized);
      }
      const ext = path.posix.extname(normalized).slice(1).toLowerCase();
      return defaultMap.get(ext) || null;
    },
    overrides: overrideMap,
    defaults: defaultMap
  };
}

export function detectImageMimeByExt(partPath) {
  const ext = path.posix.extname(partPath).toLowerCase();
  switch (ext) {
    case ".png":
      return "image/png";
    case ".jpg":
    case ".jpeg":
      return "image/jpeg";
    case ".gif":
      return "image/gif";
    case ".bmp":
      return "image/bmp";
    case ".svg":
      return "image/svg+xml";
    case ".webp":
      return "image/webp";
    case ".tif":
    case ".tiff":
      return "image/tiff";
    default:
      return "application/octet-stream";
  }
}

export function uint8ToBase64(bytes) {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }
  let binary = "";
  for (let i = 0; i < bytes.length; i += 1) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}
