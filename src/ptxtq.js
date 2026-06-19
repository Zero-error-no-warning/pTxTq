import JSZip from "jszip";
import { parsePresentationModel } from "./core/model.js";
import { renderSlideToSvg } from "./render/renderSlideToSvg.js";
import { renderSlideToCanvas } from "./render/renderSlideToCanvas.js";
import { writeModelToPptx } from "./write/pptxWriter.js";
import { appendEmbeddedElements, prepareSlideEmbedding } from "./core/slideEmbedding.js";
import {
  buildXml,
  ensureZipPath,
  parseXml,
  relsPartPath,
  relationshipMap,
  resolveTargetPath
} from "./utils/xml.js";

export const PPTX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function isArrayBufferLike(input) {
  return (
    input instanceof ArrayBuffer
    || ArrayBuffer.isView(input)
    || (typeof Buffer !== "undefined" && Buffer.isBuffer(input))
  );
}

function isBlobLike(input) {
  return Boolean(input && typeof input.arrayBuffer === "function");
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

export class BrowserOpenXmlPackage {
  constructor(zip) {
    this.zip = zip;
    this.xmlCache = new Map();
    this.relationshipCache = new Map();
  }

  static async load(source) {
    let bytes;
    if (isBlobLike(source)) {
      bytes = new Uint8Array(await source.arrayBuffer());
    } else if (isArrayBufferLike(source)) {
      bytes = toUint8Array(source);
    } else {
      throw new TypeError("Browser load expects Blob, Uint8Array, or ArrayBuffer.");
    }

    const zip = await JSZip.loadAsync(bytes);
    return new BrowserOpenXmlPackage(zip);
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
    return new BrowserOpenXmlPackage(zip);
  }

  async toUint8Array() {
    return this.zip.generateAsync({ type: "uint8array" });
  }

  async toBlob(options = {}) {
    if (typeof Blob === "undefined") {
      throw new Error("Blob is not available in this runtime");
    }
    const mimeType = typeof options === "string"
      ? options
      : options.mimeType || PPTX_MIME_TYPE;
    const bytes = await this.toUint8Array();
    return new Blob([bytes], { type: mimeType });
  }
}

export class PptxThemeDocument {
  constructor(openXmlPackage, model) {
    this._package = openXmlPackage;
    this.model = model;
  }

  static async load(source) {
    const openXmlPackage = await BrowserOpenXmlPackage.load(source);
    const model = await parsePresentationModel(openXmlPackage);

    const size = model?.metadata?.slideSizeEmu || { cx: 9144000, cy: 6858000 };
    for (const slide of model.slides || []) {
      slide.cx = size.cx;
      slide.cy = size.cy;
    }

    return new PptxThemeDocument(openXmlPackage, model);
  }

  get metadata() {
    return this.model.metadata;
  }

  get slides() {
    return this.model.slides;
  }

  get slideCount() {
    return this.model.slides.length;
  }

  getSlide(index) {
    return this.model.slides[index] || null;
  }

  embedSlideIntoSlide(sourceIndex, targetIndex, options = {}) {
    const sourceSlide = this.getSlide(sourceIndex);
    const targetSlide = this.getSlide(targetIndex);
    if (!sourceSlide) {
      throw new Error(`Source slide index out of range: ${sourceIndex}`);
    }
    if (!targetSlide) {
      throw new Error(`Target slide index out of range: ${targetIndex}`);
    }

    const inserted = prepareSlideEmbedding(
      sourceSlide,
      targetSlide,
      sourceIndex,
      options,
      this.model?.metadata
    );
    appendEmbeddedElements(targetSlide, inserted);

    return {
      targetIndex,
      sourceIndex,
      insertedCount: inserted.length
    };
  }

  async renderSlide(index, options = {}) {
    const slide = this.getSlide(index);
    if (!slide) {
      throw new Error(`Slide index out of range: ${index}`);
    }

    const mode = options.mode || "svg";
    if (mode === "svg") {
      return renderSlideToSvg(slide, {
        ...options,
        slideSizeEmu: this.model.metadata.slideSizeEmu
      });
    }

    if (mode === "canvas") {
      return renderSlideToCanvas(slide, options.target, {
        ...options,
        slideSizeEmu: this.model.metadata.slideSizeEmu
      });
    }

    throw new Error(`Unsupported render mode: ${mode}`);
  }

  async renderSlides(options = {}) {
    const indices = Array.isArray(options.indices)
      ? options.indices
      : this.slides.map((_, index) => index);
    const results = [];
    for (const index of indices) {
      results.push(await this.renderSlide(index, options));
    }
    return results;
  }

  async toPptxPackage(options = {}) {
    return writeModelToPptx(this._package, this.model, options);
  }

  async toPptxBuffer(options = {}) {
    const packageOut = await this.toPptxPackage(options);
    if (options.type === "blob") {
      return packageOut.toBlob({ mimeType: options.mimeType || PPTX_MIME_TYPE });
    }
    return packageOut.toUint8Array();
  }

  async toPptxBlob(options = {}) {
    const packageOut = await this.toPptxPackage(options);
    return packageOut.toBlob({ mimeType: options.mimeType || PPTX_MIME_TYPE });
  }

  toJsonObject(options = {}) {
    const includePrivate = options.includePrivate === true;
    const includeDataUri = options.includeDataUri !== false;
    const includeRaw = options.includeRaw !== false;

    const replacer = (key, value) => {
      if (!includePrivate && key.startsWith("_")) {
        return undefined;
      }
      if (!includeDataUri && key === "dataUri") {
        return undefined;
      }
      if (!includeRaw && key === "raw") {
        return undefined;
      }
      return value;
    };

    return JSON.parse(JSON.stringify(this.model, replacer));
  }

  toJsonString(options = {}) {
    const space = Number.isInteger(options.space) ? options.space : 2;
    return JSON.stringify(this.toJsonObject(options), null, space);
  }
}

export async function loadPptx(source) {
  return PptxThemeDocument.load(source);
}

export default {
  BrowserOpenXmlPackage,
  PPTX_MIME_TYPE,
  PptxThemeDocument,
  loadPptx
};
