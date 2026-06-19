import { PptxThemeDocument } from "./core/PptxDocument.js";
import { htmlToPptxDocument, htmlToPptxBuffer } from "./htmlToPptx.js";

export { PptxThemeDocument, htmlToPptxDocument, htmlToPptxBuffer };

export async function loadPptx(source) {
  return PptxThemeDocument.load(source);
}

export default {
  PptxThemeDocument,
  loadPptx,
  htmlToPptxDocument,
  htmlToPptxBuffer
};
