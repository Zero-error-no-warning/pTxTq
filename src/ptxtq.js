import JSZip from "jszip";
import { parsePresentationModel } from "./core/model.js";
import { renderSlideToSvg } from "./render/renderSlideToSvg.js";
import { renderSlideToCanvas } from "./render/renderSlideToCanvas.js";
import { writeModelToPptx } from "./write/pptxWriter.js";
import { appendEmbeddedElements, prepareSlideEmbedding } from "./core/slideEmbedding.js";
import { pxToEmu } from "./utils/units.js";
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


const HTML_SLIDE_W = 9144000;
const HTML_SLIDE_H = 5143500;
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

function htmlB64(value) {
  if (typeof Buffer !== "undefined") return Buffer.from(value, "utf8").toString("base64");
  return btoa(unescape(encodeURIComponent(value)));
}
function stripHtmlTags(value) {
  return String(value || "").replace(/<style[\s\S]*?<\/style>/gi, "").replace(/<script[\s\S]*?<\/script>/gi, "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}
function htmlAttr(tag, name) {
  return tag.match(new RegExp(`${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`, "i"))?.slice(1).find(Boolean) || null;
}
function cssPx(value, fallback) {
  const n = Number.parseFloat(String(value ?? "").replace("px", ""));
  return Number.isFinite(n) ? n : fallback;
}
function cssValue(style, name) {
  return String(style || "").match(new RegExp(`${name}\\s*:\\s*([^;]+)`, "i"))?.[1]?.trim() || null;
}
function cssColor(value, fallback = "#000000") {
  if (!value) return fallback;
  const v = String(value).trim();
  if (/^#[0-9a-f]{3}$/i.test(v)) return `#${v[1]}${v[1]}${v[2]}${v[2]}${v[3]}${v[3]}`;
  if (/^#[0-9a-f]{6}$/i.test(v)) return v;
  const rgb = v.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  return rgb ? `#${[rgb[1], rgb[2], rgb[3]].map((x) => Number(x).toString(16).padStart(2, "0")).join("")}` : fallback;
}
function htmlTextElement(id, text, box, style = {}) {
  return { id, type: "text", name: `HTML Text ${id}`, x: pxToEmu(box.x), y: pxToEmu(box.y), cx: pxToEmu(box.w), cy: pxToEmu(box.h), shapeType: "rect", fill: { type: "none" }, line: { color: "#FFFFFF", alpha: 0 }, text: { verticalAlign: "t", wrap: "square", paragraphs: [{ alignment: style.align || "l", runs: [{ text, style: { fontSize: style.fontSize || 18, color: style.color || "#000000", fontFamily: style.fontFamily || "Arial", bold: !!style.bold } }] }] } };
}
function htmlSvgElement(id, svg, box) {
  const width = cssPx(htmlAttr(svg, "width"), box.w);
  const height = cssPx(htmlAttr(svg, "height"), box.h);
  return { id, type: "image", name: `HTML SVG ${id}`, x: pxToEmu(box.x), y: pxToEmu(box.y), cx: pxToEmu(box.w || width), cy: pxToEmu(box.h || height), mimeType: "image/svg+xml", dataUri: `data:image/svg+xml;base64,${htmlB64(svg)}` };
}
function htmlImageElement(id, src, box) {
  const mimeType = src.match(/^data:([^;]+);base64,/)?.[1] || "image/png";
  return { id, type: "image", name: `HTML Image ${id}`, x: pxToEmu(box.x), y: pxToEmu(box.y), cx: pxToEmu(box.w), cy: pxToEmu(box.h), mimeType, dataUri: src };
}
function htmlToElements(html, options = {}) {
  const elements = [];
  let id = 2;
  let y = options.paddingPx ?? 36;
  const x = options.paddingPx ?? 36;
  const w = options.widthPx ?? 888;
  const source = typeof html === "string" ? html : html?.outerHTML || String(html || "");
  source.replace(/<svg\b[\s\S]*?<\/svg>/gi, (svg) => { elements.push(htmlSvgElement(id++, svg, { x, y, w, h: cssPx(htmlAttr(svg, "height"), 240) })); y += cssPx(htmlAttr(svg, "height"), 240) + 18; return svg; });
  source.replace(/<img\b[^>]*>/gi, (tag) => { const src = htmlAttr(tag, "src"); if (src?.startsWith("data:")) { elements.push(htmlImageElement(id++, src, { x, y, w: cssPx(htmlAttr(tag, "width"), 320), h: cssPx(htmlAttr(tag, "height"), 180) })); y += cssPx(htmlAttr(tag, "height"), 180) + 18; } return tag; });
  const withoutMedia = source.replace(/<svg\b[\s\S]*?<\/svg>/gi, " ").replace(/<img\b[^>]*>/gi, " ");
  const blockRe = /<(h[1-6]|p|div|li|span)\b([^>]*)>([\s\S]*?)<\/\1>/gi;
  let matched = false;
  withoutMedia.replace(blockRe, (all, tag, attrs, inner) => {
    const text = stripHtmlTags(inner);
    if (!text) return all;
    matched = true;
    const st = htmlAttr(attrs, "style") || "";
    const fs = tag === "h1" ? 32 : tag === "h2" ? 26 : tag?.startsWith("h") ? 22 : cssPx(cssValue(st, "font-size"), 18);
    elements.push(htmlTextElement(id++, tag === "li" ? `• ${text}` : text, { x, y, w, h: Math.max(36, fs * 1.8) }, { fontSize: fs, color: cssColor(cssValue(st, "color")), bold: /^h/i.test(tag), align: ({ center: "ctr", right: "r" })[cssValue(st, "text-align")] || "l", fontFamily: cssValue(st, "font-family") || "Arial" }));
    y += Math.max(36, fs * 1.8) + 8;
    return all;
  });
  if (!matched) {
    const text = stripHtmlTags(withoutMedia);
    if (text) elements.push(htmlTextElement(id++, text, { x, y, w, h: 120 }));
  }
  return elements;
}
function addXml(zip, path, xml) { zip.file(path, xml); }
async function blankHtmlPackage() {
  const zip = new JSZip();
  addXml(zip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="svg" ContentType="image/svg+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  addXml(zip, "_rels/.rels", `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${DOC_REL}" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
  addXml(zip, "docProps/core.xml", `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>HTML to PPTX</dc:title><dc:creator>pTxTq</dc:creator><cp:lastModifiedBy>pTxTq</cp:lastModifiedBy></cp:coreProperties>`);
  addXml(zip, "docProps/app.xml", `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>pTxTq</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Slides>1</Slides></Properties>`);
  addXml(zip, "ppt/presentation.xml", `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId2"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="${HTML_SLIDE_W}" cy="${HTML_SLIDE_H}" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle></p:presentation>`);
  addXml(zip, "ppt/_rels/presentation.xml.rels", `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${SLIDE_REL}" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>`);
  addXml(zip, "ppt/presProps.xml", `<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`);
  addXml(zip, "ppt/viewProps.xml", `<p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr></p:viewPr>`);
  addXml(zip, "ppt/tableStyles.xml", `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
  addXml(zip, "ppt/theme/theme1.xml", `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`);
  addXml(zip, "ppt/slideMasters/slideMaster1.xml", `<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`);
  addXml(zip, "ppt/slideMasters/_rels/slideMaster1.xml.rels", `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`);
  addXml(zip, "ppt/slideLayouts/slideLayout1.xml", `<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`);
  addXml(zip, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  addXml(zip, "ppt/slides/slide1.xml", `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Slide 1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  addXml(zip, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
  return new BrowserOpenXmlPackage(zip);
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

export async function htmlToPptxDocument(html, options = {}) {
  const pkg = await blankHtmlPackage();
  const doc = await PptxThemeDocument.load(await pkg.toUint8Array());
  const slide = doc.getSlide(0);
  slide.elements = htmlToElements(html, options);
  slide.renderElements = slide.elements;
  slide._snapshot = null;
  slide.cx = HTML_SLIDE_W;
  slide.cy = HTML_SLIDE_H;
  return doc;
}

export async function htmlToPptxBuffer(html, options = {}) {
  const doc = await htmlToPptxDocument(html, options);
  return doc.toPptxBuffer({ type: options.type });
}

export async function htmlToPptxBlob(html, options = {}) {
  const doc = await htmlToPptxDocument(html, options);
  return doc.toPptxBlob(options);
}

export default {
  BrowserOpenXmlPackage,
  PPTX_MIME_TYPE,
  PptxThemeDocument,
  loadPptx,
  htmlToPptxDocument,
  htmlToPptxBuffer,
  htmlToPptxBlob
};
