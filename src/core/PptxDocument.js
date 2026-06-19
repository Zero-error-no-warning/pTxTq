import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { OpenXmlPackage } from "./openxml.js";
import { parsePresentationModel } from "./model.js";
import { renderSlideToSvg } from "../render/renderSlideToSvg.js";
import { renderSlideToCanvas } from "../render/renderSlideToCanvas.js";
import { writeModelToPptx } from "../write/pptxWriter.js";

import { appendEmbeddedElements, prepareSlideEmbedding } from "./slideEmbedding.js";

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
