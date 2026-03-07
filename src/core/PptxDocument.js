import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { OpenXmlPackage } from "./openxml.js";
import { parsePresentationModel } from "./model.js";
import { renderSlideToSvg } from "../render/renderSlideToSvg.js";
import { renderSlideToCanvas } from "../render/renderSlideToCanvas.js";
import { writeModelToPptx } from "../write/pptxWriter.js";

const EMBEDDABLE_TYPES = new Set(["shape", "text", "line", "image", "table"]);

function deepCloneValue(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function toNumber(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function nextElementId(slide) {
  let maxId = 0;
  for (const element of slide?.elements || []) {
    const parsed = Number.parseInt(String(element?.id || "0"), 10);
    if (Number.isFinite(parsed)) {
      maxId = Math.max(maxId, parsed);
    }
  }
  return maxId;
}

function scaleTextBody(textBody, scaleX, scaleY, fontScale) {
  if (!textBody) {
    return textBody;
  }
  const cloned = deepCloneValue(textBody);
  cloned.leftInset = toNumber(cloned.leftInset, 0) * scaleX;
  cloned.rightInset = toNumber(cloned.rightInset, 0) * scaleX;
  cloned.topInset = toNumber(cloned.topInset, 0) * scaleY;
  cloned.bottomInset = toNumber(cloned.bottomInset, 0) * scaleY;

  for (const paragraph of cloned.paragraphs || []) {
    paragraph.marginLeft = toNumber(paragraph.marginLeft, 0) * scaleX;
    paragraph.marginRight = toNumber(paragraph.marginRight, 0) * scaleX;
    paragraph.indent = toNumber(paragraph.indent, 0) * scaleX;

    for (const run of paragraph.runs || []) {
      if (run?.style?.fontSizePt) {
        run.style.fontSizePt = toNumber(run.style.fontSizePt, 11) * fontScale;
      }
    }
  }

  return cloned;
}

function scaleTableElement(table, scaleX, scaleY, fontScale) {
  table.gridCols = (table.gridCols || []).map((w) => toNumber(w, 0) * scaleX);
  for (const row of table.rows || []) {
    row.height = toNumber(row.height, 0) * scaleY;
    for (const cell of row.cells || []) {
      cell.marginLeft = toNumber(cell.marginLeft, 0) * scaleX;
      cell.marginRight = toNumber(cell.marginRight, 0) * scaleX;
      cell.marginTop = toNumber(cell.marginTop, 0) * scaleY;
      cell.marginBottom = toNumber(cell.marginBottom, 0) * scaleY;
      cell.text = scaleTextBody(cell.text, scaleX, scaleY, fontScale);
      for (const side of ["left", "right", "top", "bottom"]) {
        if (cell.borders?.[side]?.width) {
          cell.borders[side].width = toNumber(cell.borders[side].width, 0) * ((scaleX + scaleY) / 2);
        }
      }
    }
  }
}

function transformEmbeddedElement(element, transform) {
  const out = deepCloneValue(element);
  const { sourceX, sourceY, targetX, targetY, scaleX, scaleY, fontScale } = transform;

  out.x = targetX + (toNumber(out.x, 0) - sourceX) * scaleX;
  out.y = targetY + (toNumber(out.y, 0) - sourceY) * scaleY;
  out.cx = toNumber(out.cx, 0) * scaleX;
  out.cy = toNumber(out.cy, 0) * scaleY;

  if (out.line?.width) {
    out.line.width = toNumber(out.line.width, 0) * ((scaleX + scaleY) / 2);
  }
  if (out.text) {
    out.text = scaleTextBody(out.text, scaleX, scaleY, fontScale);
  }
  if (out.type === "table") {
    scaleTableElement(out, scaleX, scaleY, fontScale);
  }

  return out;
}

export class PptxThemeDocument {
  constructor(openXmlPackage, model) {
    this._package = openXmlPackage;
    this.model = model;
  }

  static async load(source) {
    const openXmlPackage = await OpenXmlPackage.load(source);
    const model = await parsePresentationModel(openXmlPackage);

    const size = model?.metadata?.slideSizeEmu || { cx: 9144000, cy: 6858000 };
    for (const slide of model.slides || []) {
      slide.cx = size.cx;
      slide.cy = size.cy;
    }

    return new PptxThemeDocument(openXmlPackage, model);
  }

  static async loadFile(filePath) {
    return this.load(filePath);
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

  /**
   * Embed one slide as editable vector elements into another slide.
   * Browser-side operation: no server step required.
   */
  embedSlideIntoSlide(sourceIndex, targetIndex, options = {}) {
    const sourceSlide = this.getSlide(sourceIndex);
    const targetSlide = this.getSlide(targetIndex);
    if (!sourceSlide) {
      throw new Error(`Source slide index out of range: ${sourceIndex}`);
    }
    if (!targetSlide) {
      throw new Error(`Target slide index out of range: ${targetIndex}`);
    }

    const useRenderElements = options.useRenderElements !== false;
    const includeBackground = options.includeBackground !== false;
    const includeUnsupported = options.includeUnsupported === true;

    const sourceWidth = toNumber(sourceSlide.cx, toNumber(this.model?.metadata?.slideSizeEmu?.cx, 9144000));
    const sourceHeight = toNumber(sourceSlide.cy, toNumber(this.model?.metadata?.slideSizeEmu?.cy, 6858000));
    const targetX = toNumber(options.x, 0);
    const targetY = toNumber(options.y, 0);
    const targetCx = toNumber(options.cx, sourceWidth);
    const targetCy = toNumber(options.cy, sourceHeight);

    const scaleX = sourceWidth ? targetCx / sourceWidth : 1;
    const scaleY = sourceHeight ? targetCy / sourceHeight : 1;
    const fontScale = (scaleX + scaleY) / 2;
    const sourceElements = useRenderElements
      ? (sourceSlide.renderElements || sourceSlide.elements || [])
      : (sourceSlide.elements || []);

    const transform = {
      sourceX: 0,
      sourceY: 0,
      targetX,
      targetY,
      scaleX,
      scaleY,
      fontScale
    };

    const inserted = [];
    let idCounter = nextElementId(targetSlide) + 1;

    if (includeBackground && sourceSlide.background?.type === "solid" && sourceSlide.background?.color) {
      inserted.push({
        id: String(idCounter++),
        name: `Embedded Background ${sourceIndex + 1}`,
        description: "Embedded source slide background",
        type: "shape",
        shapeType: "rect",
        x: targetX,
        y: targetY,
        cx: targetCx,
        cy: targetCy,
        rotation: 0,
        flipH: false,
        flipV: false,
        fill: {
          type: "solid",
          color: sourceSlide.background.color,
          alpha: sourceSlide.background.alpha ?? 1,
          source: "embedded-slide-background"
        },
        line: {
          width: 0,
          color: null,
          alpha: 0,
          dash: "none",
          source: "embedded-slide-background"
        },
        text: null
      });
    }

    if (includeBackground && sourceSlide.background?.type === "image" && sourceSlide.background?.dataUri) {
      inserted.push({
        id: String(idCounter++),
        name: `Embedded Background Image ${sourceIndex + 1}`,
        description: "Embedded source slide background image",
        type: "image",
        x: targetX,
        y: targetY,
        cx: targetCx,
        cy: targetCy,
        rotation: 0,
        flipH: false,
        flipV: false,
        imagePath: sourceSlide.background.imagePath || null,
        mimeType: sourceSlide.background.mimeType || null,
        dataUri: sourceSlide.background.dataUri,
        fill: null,
        line: null,
        text: null
      });
    }

    for (const element of sourceElements) {
      if (!includeUnsupported && !EMBEDDABLE_TYPES.has(element?.type)) {
        continue;
      }
      const embedded = transformEmbeddedElement(element, transform);
      embedded.id = String(idCounter++);
      embedded.name = `${embedded.name || embedded.type} (embedded from slide ${sourceIndex + 1})`;
      embedded.sourceLayer = "embedded-slide";
      inserted.push(embedded);
    }

    if (!Array.isArray(targetSlide.elements)) {
      targetSlide.elements = [];
    }
    targetSlide.elements.push(...inserted);

    if (Array.isArray(targetSlide.renderElements)) {
      targetSlide.renderElements.push(...inserted);
    } else {
      targetSlide.renderElements = [...targetSlide.elements];
    }

    targetSlide._snapshot = null;
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

  async toPptxPackage(options = {}) {
    return writeModelToPptx(this._package, this.model, options);
  }

  async toPptxBuffer(options = {}) {
    const packageOut = await this.toPptxPackage(options);
    if (options.type === "uint8array") {
      return packageOut.toUint8Array();
    }
    return packageOut.toNodeBuffer();
  }

  async saveAs(filePath, options = {}) {
    const buffer = await this.toPptxBuffer(options);
    await writeFile(filePath, buffer);
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

  async saveModelJson(filePath, options = {}) {
    const outDir = path.dirname(path.resolve(filePath));
    await mkdir(outDir, { recursive: true });
    await writeFile(filePath, this.toJsonString(options), "utf8");
  }
}
