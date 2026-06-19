import JSZip from "jszip";
import { OpenXmlPackage } from "./core/openxml.js";
import { PptxThemeDocument } from "./core/PptxDocument.js";
import { pxToEmu } from "./utils/units.js";

const SLIDE_W = 9144000;
const SLIDE_H = 5143500;
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

function b64(s) { return typeof Buffer !== "undefined" ? Buffer.from(s, "utf8").toString("base64") : btoa(unescape(encodeURIComponent(s))); }
function stripTags(s) { return String(s || "").replace(/<style[\s\S]*?<\/style>/gi, "").replace(/<script[\s\S]*?<\/script>/gi, "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim(); }
function attr(tag, name) { return tag.match(new RegExp(`${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`, "i"))?.slice(1).find(Boolean) || null; }
function px(value, fallback) { const n = Number.parseFloat(String(value ?? "").replace("px", "")); return Number.isFinite(n) ? n : fallback; }
function styleValue(style, name) { return String(style || "").match(new RegExp(`${name}\\s*:\\s*([^;]+)`, "i"))?.[1]?.trim() || null; }
function color(value, fallback = "#000000") {
  if (!value) return fallback;
  const v = String(value).trim();
  if (/^#[0-9a-f]{3}$/i.test(v)) return `#${v[1]}${v[1]}${v[2]}${v[2]}${v[3]}${v[3]}`;
  if (/^#[0-9a-f]{6}$/i.test(v)) return v;
  const rgb = v.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  return rgb ? `#${[rgb[1], rgb[2], rgb[3]].map((x) => Number(x).toString(16).padStart(2, "0")).join("")}` : fallback;
}

function createTextElement(id, text, box, style = {}) {
  return { id, type: "text", name: `HTML Text ${id}`, x: pxToEmu(box.x), y: pxToEmu(box.y), cx: pxToEmu(box.w), cy: pxToEmu(box.h), shapeType: "rect", fill: { type: "none" }, line: { color: "#FFFFFF", alpha: 0 }, text: { verticalAlign: "t", wrap: "square", paragraphs: [{ alignment: style.align || "l", runs: [{ text, style: { fontSize: style.fontSize || 18, color: style.color || "#000000", fontFamily: style.fontFamily || "Arial", bold: !!style.bold } }] }] } };
}
function createSvgElement(id, svg, box) {
  const width = px(attr(svg, "width"), box.w);
  const height = px(attr(svg, "height"), box.h);
  return { id, type: "image", name: `HTML SVG ${id}`, x: pxToEmu(box.x), y: pxToEmu(box.y), cx: pxToEmu(box.w || width), cy: pxToEmu(box.h || height), mimeType: "image/svg+xml", dataUri: `data:image/svg+xml;base64,${b64(svg)}` };
}
function createImageElement(id, src, box) {
  const mimeType = src.match(/^data:([^;]+);base64,/)?.[1] || "image/png";
  return { id, type: "image", name: `HTML Image ${id}`, x: pxToEmu(box.x), y: pxToEmu(box.y), cx: pxToEmu(box.w), cy: pxToEmu(box.h), mimeType, dataUri: src };
}

function htmlToElements(html, options = {}) {
  const elements = [];
  let id = 2, y = options.paddingPx ?? 36;
  const x = options.paddingPx ?? 36, w = options.widthPx ?? 888;
  const source = typeof html === "string" ? html : html?.outerHTML || String(html || "");
  source.replace(/<svg\b[\s\S]*?<\/svg>/gi, (svg) => { elements.push(createSvgElement(id++, svg, { x, y, w, h: px(attr(svg, "height"), 240) })); y += px(attr(svg, "height"), 240) + 18; return svg; });
  source.replace(/<img\b[^>]*>/gi, (tag) => { const src = attr(tag, "src"); if (src?.startsWith("data:")) { elements.push(createImageElement(id++, src, { x, y, w: px(attr(tag, "width"), 320), h: px(attr(tag, "height"), 180) })); y += px(attr(tag, "height"), 180) + 18; } return tag; });
  const withoutMedia = source.replace(/<svg\b[\s\S]*?<\/svg>/gi, " ").replace(/<img\b[^>]*>/gi, " ");
  const blockRe = /<(h[1-6]|p|div|li|span)\b([^>]*)>([\s\S]*?)<\/\1>/gi;
  let matched = false;
  withoutMedia.replace(blockRe, (all, tag, attrs, inner) => {
    const text = stripTags(inner); if (!text) return all; matched = true;
    const st = attr(attrs, "style") || "";
    const fs = tag === "h1" ? 32 : tag === "h2" ? 26 : tag?.startsWith("h") ? 22 : px(styleValue(st, "font-size"), 18);
    elements.push(createTextElement(id++, tag === "li" ? `• ${text}` : text, { x, y, w, h: Math.max(36, fs * 1.8) }, { fontSize: fs, color: color(styleValue(st, "color")), bold: /^h/i.test(tag), align: ({ center: "ctr", right: "r" })[styleValue(st, "text-align")] || "l", fontFamily: styleValue(st, "font-family") || "Arial" }));
    y += Math.max(36, fs * 1.8) + 8; return all;
  });
  if (!matched) { const text = stripTags(withoutMedia); if (text) elements.push(createTextElement(id++, text, { x, y, w, h: 120 })); }
  return elements;
}

function addXml(zip, path, xml) { zip.file(path, xml); }
async function blankPackage() {
  const zip = new JSZip();
  addXml(zip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="svg" ContentType="image/svg+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  addXml(zip, "_rels/.rels", `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${DOC_REL}" Target="ppt/presentation.xml"/></Relationships>`);
  addXml(zip, "ppt/presentation.xml", `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst/><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="${SLIDE_W}" cy="${SLIDE_H}" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  addXml(zip, "ppt/_rels/presentation.xml.rels", `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${SLIDE_REL}" Target="slides/slide1.xml"/></Relationships>`);
  addXml(zip, "ppt/slides/slide1.xml", `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Slide 1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  addXml(zip, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="${REL_NS}"/>`);
  return new OpenXmlPackage(zip);
}

export async function htmlToPptxDocument(html, options = {}) {
  const pkg = await blankPackage();
  const doc = await PptxThemeDocument.load(await pkg.toUint8Array());
  const slide = doc.getSlide(0);
  slide.elements = htmlToElements(html, options);
  slide.renderElements = slide.elements;
  slide._snapshot = null;
  slide.cx = SLIDE_W; slide.cy = SLIDE_H;
  return doc;
}
export async function htmlToPptxBuffer(html, options = {}) {
  const doc = await htmlToPptxDocument(html, options);
  return doc.toPptxBuffer({ type: options.type });
}
