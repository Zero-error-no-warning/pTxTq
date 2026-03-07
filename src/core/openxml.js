import JSZip from "jszip";
import { readFile } from "node:fs/promises";
import {
  buildXml,
  ensureZipPath,
  parseXml,
  relsPartPath,
  relationshipMap,
  resolveTargetPath
} from "../utils/xml.js";

function isArrayBufferLike(input) {
  return (
    input instanceof ArrayBuffer
    || ArrayBuffer.isView(input)
    || (typeof Buffer !== "undefined" && Buffer.isBuffer(input))
  );
}

function toUint8Array(input) {
  if (input instanceof Uint8Array) {
    return input;
  }
  if (input instanceof ArrayBuffer) {
    return new Uint8Array(input);
  }
  if (ArrayBuffer.isView(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  if (typeof Buffer !== "undefined" && Buffer.isBuffer(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  throw new TypeError("Expected ArrayBuffer-like input.");
}

export class OpenXmlPackage {
  constructor(zip) {
    this.zip = zip;
    this.xmlCache = new Map();
    this.relationshipCache = new Map();
  }

  static async load(source) {
    let bytes;
    if (typeof source === "string") {
      const fileBuffer = await readFile(source);
      bytes = new Uint8Array(fileBuffer.buffer, fileBuffer.byteOffset, fileBuffer.byteLength);
    } else if (isArrayBufferLike(source)) {
      bytes = toUint8Array(source);
    } else {
      throw new TypeError("source must be a file path, Buffer, Uint8Array, or ArrayBuffer");
    }

    const zip = await JSZip.loadAsync(bytes);
    return new OpenXmlPackage(zip);
  }

  hasPart(partPath) {
    const normalized = ensureZipPath(partPath);
    return !!this.zip.file(normalized);
  }

  listParts() {
    return Object.keys(this.zip.files).filter((name) => !this.zip.files[name].dir).sort();
  }

  async readText(partPath) {
    const normalized = ensureZipPath(partPath);
    const file = this.zip.file(normalized);
    if (!file) {
      throw new Error(`Part not found: ${normalized}`);
    }
    return file.async("string");
  }

  async readBinary(partPath) {
    const normalized = ensureZipPath(partPath);
    const file = this.zip.file(normalized);
    if (!file) {
      throw new Error(`Part not found: ${normalized}`);
    }
    return file.async("uint8array");
  }

  async readXml(partPath) {
    const normalized = ensureZipPath(partPath);
    if (this.xmlCache.has(normalized)) {
      return this.xmlCache.get(normalized);
    }
    const xml = await this.readText(normalized);
    const parsed = parseXml(xml);
    this.xmlCache.set(normalized, parsed);
    return parsed;
  }

  writeText(partPath, text) {
    const normalized = ensureZipPath(partPath);
    this.zip.file(normalized, text);
    this.xmlCache.delete(normalized);
    this.relationshipCache.delete(normalized);
  }

  writeBinary(partPath, bytes) {
    const normalized = ensureZipPath(partPath);
    this.zip.file(normalized, bytes);
    this.xmlCache.delete(normalized);
    this.relationshipCache.delete(normalized);
  }

  writeXml(partPath, xmlObject) {
    this.writeText(partPath, buildXml(xmlObject));
  }

  deletePart(partPath) {
    const normalized = ensureZipPath(partPath);
    this.zip.remove(normalized);
    this.xmlCache.delete(normalized);
    this.relationshipCache.delete(normalized);
  }

  async getRelationships(partPath) {
    const normalizedPart = ensureZipPath(partPath);
    if (this.relationshipCache.has(normalizedPart)) {
      return this.relationshipCache.get(normalizedPart);
    }

    const relsPath = relsPartPath(normalizedPart);
    if (!this.hasPart(relsPath)) {
      const empty = new Map();
      this.relationshipCache.set(normalizedPart, empty);
      return empty;
    }

    const relsXml = await this.readXml(relsPath);
    const map = relationshipMap(relsXml);
    this.relationshipCache.set(normalizedPart, map);
    return map;
  }

  async resolveRelationship(partPath, relationshipId) {
    const rels = await this.getRelationships(partPath);
    const rel = rels.get(relationshipId);
    if (!rel) {
      return null;
    }
    if (rel.targetMode === "External") {
      return {
        ...rel,
        targetPath: rel.target
      };
    }
    return {
      ...rel,
      targetPath: resolveTargetPath(partPath, rel.target)
    };
  }

  async clone() {
    const bytes = await this.zip.generateAsync({ type: "uint8array" });
    const zip = await JSZip.loadAsync(bytes);
    return new OpenXmlPackage(zip);
  }

  async toUint8Array() {
    return this.zip.generateAsync({ type: "uint8array" });
  }

  async toNodeBuffer() {
    return this.zip.generateAsync({ type: "nodebuffer" });
  }
}
