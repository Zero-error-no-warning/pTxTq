import { PptxThemeDocument } from "./core/PptxDocument.js";

export { PptxThemeDocument };

export async function loadPptx(source) {
  return PptxThemeDocument.load(source);
}

export default {
  PptxThemeDocument,
  loadPptx
};
