import path from "node:path";
import {
  createSlideSnapshot,
  buildShapeSpPr,
  buildTextBodyNode,
  modelColorToOpenXmlFill,
  modelLineToOpenXmlLn
} from "../core/model.js";
import { ensureArray, toInt, deepClone } from "../utils/object.js";

const MANAGED_SP_TREE_KEYS = new Set([
  "p:sp",
  "p:cxnSp",
  "p:pic",
  "p:graphicFrame"
]);

function relativeTarget(fromPartPath, toPartPath) {
  const fromDir = path.posix.dirname(fromPartPath);
  const rel = path.posix.relative(fromDir, toPartPath).replace(/\\/g, "/");
  return rel || path.posix.basename(toPartPath);
}

function decodeDataUri(dataUri) {
  if (!dataUri || typeof dataUri !== "string") {
    return null;
  }
  const match = dataUri.match(/^data:([^;]+);base64,(.*)$/);
  if (!match) {
    return null;
  }
  const base64 = match[2];
  if (typeof Buffer !== "undefined") {
    return Buffer.from(base64, "base64");
  }
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function mapMimeToExt(mimeType = "") {
  switch (mimeType.toLowerCase()) {
    case "image/png":
      return "png";
    case "image/jpeg":
      return "jpg";
    case "image/gif":
      return "gif";
    case "image/bmp":
      return "bmp";
    case "image/webp":
      return "webp";
    case "image/tiff":
      return "tiff";
    case "image/svg+xml":
      return "svg";
    default:
      return "png";
  }
}

function buildXfrmNode(element) {
  const xfrm = {
    "a:off": {
      "@_x": toInt(element.x, 0),
      "@_y": toInt(element.y, 0)
    },
    "a:ext": {
      "@_cx": toInt(element.cx, 0),
      "@_cy": toInt(element.cy, 0)
    }
  };

  if (element.rotation) {
    xfrm["@_rot"] = toInt((element.rotation || 0) * 60000, 0);
  }
  if (element.flipH) {
    xfrm["@_flipH"] = "1";
  }
  if (element.flipV) {
    xfrm["@_flipV"] = "1";
  }

  return xfrm;
}

function applyCommonNonVisualProps(cNvPrNode, element, fallbackName) {
  const cNvPr = cNvPrNode || {};
  cNvPr["@_id"] = toInt(element.id, 0);
  cNvPr["@_name"] = element.name || `${fallbackName} ${element.id}`;
  cNvPr["@_descr"] = element.description || "";
  if (element.hidden) {
    cNvPr["@_hidden"] = "1";
  } else {
    delete cNvPr["@_hidden"];
  }
  return cNvPr;
}

function createNvSpPr(element) {
  return {
    "p:cNvPr": {
      "@_id": toInt(element.id, 0),
      "@_name": element.name || `Shape ${element.id}`,
      "@_descr": element.description || ""
    },
    "p:cNvSpPr": {
      "@_txBox": element.type === "text" ? "1" : "0"
    },
    "p:nvPr": element.placeholder
      ? {
          "p:ph": {
            "@_type": element.placeholder.type || "body",
            "@_idx": element.placeholder.idx || "0"
          }
        }
      : {}
  };
}

function createShapeNode(element) {
  const node = {
    "p:nvSpPr": createNvSpPr(element),
    "p:spPr": buildShapeSpPr(element)
  };

  if (element.text) {
    node["p:txBody"] = buildTextBodyNode(element.text);
  }

  return node;
}

function createConnectorNode(element) {
  return {
    "p:nvCxnSpPr": {
      "p:cNvPr": {
        "@_id": toInt(element.id, 0),
        "@_name": element.name || `Connector ${element.id}`,
        "@_descr": element.description || ""
      },
      "p:cNvCxnSpPr": {},
      "p:nvPr": {}
    },
    "p:spPr": buildShapeSpPr({ ...element, shapeType: element.shapeType || "line" })
  };
}

function createPictureNode(element) {
  return {
    "p:nvPicPr": {
      "p:cNvPr": {
        "@_id": toInt(element.id, 0),
        "@_name": element.name || `Picture ${element.id}`,
        "@_descr": element.description || ""
      },
      "p:cNvPicPr": {
        "a:picLocks": {
          "@_noChangeAspect": "1"
        }
      },
      "p:nvPr": {}
    },
    "p:blipFill": {
      "a:blip": element.relId ? { "@_r:embed": element.relId } : {},
      "a:stretch": {
        "a:fillRect": {}
      }
    },
    "p:spPr": {
      "a:xfrm": buildXfrmNode(element),
      "a:prstGeom": {
        "@_prst": "rect",
        "a:avLst": {}
      }
    }
  };
}

function tableCellToNode(cell) {
  const tcPr = {
    "@_anchor": cell.verticalAlign || "t",
    "@_marL": toInt(cell.marginLeft, 0),
    "@_marR": toInt(cell.marginRight, 0),
    "@_marT": toInt(cell.marginTop, 0),
    "@_marB": toInt(cell.marginBottom, 0)
  };

  if (cell.gridSpan && cell.gridSpan > 1) tcPr["@_gridSpan"] = toInt(cell.gridSpan, 1);
  if (cell.rowSpan && cell.rowSpan > 1) tcPr["@_rowSpan"] = toInt(cell.rowSpan, 1);
  if (cell.hMerge) tcPr["@_hMerge"] = "1";
  if (cell.vMerge) tcPr["@_vMerge"] = "1";

  Object.assign(tcPr, modelColorToOpenXmlFill(cell.fill));

  const left = modelLineToOpenXmlLn(cell.borders?.left || null)["a:ln"];
  const right = modelLineToOpenXmlLn(cell.borders?.right || null)["a:ln"];
  const top = modelLineToOpenXmlLn(cell.borders?.top || null)["a:ln"];
  const bottom = modelLineToOpenXmlLn(cell.borders?.bottom || null)["a:ln"];

  if (left) tcPr["a:lnL"] = left;
  if (right) tcPr["a:lnR"] = right;
  if (top) tcPr["a:lnT"] = top;
  if (bottom) tcPr["a:lnB"] = bottom;

  return {
    "a:txBody": buildTextBodyNode(cell.text),
    "a:tcPr": tcPr
  };
}

function createTableNode(element) {
  const tbl = {
    "a:tblPr": {
      "@_firstRow": element.firstRow ? "1" : "0",
      "@_firstCol": element.firstCol ? "1" : "0",
      "@_lastRow": element.lastRow ? "1" : "0",
      "@_lastCol": element.lastCol ? "1" : "0",
      "@_bandRow": element.bandRow ? "1" : "0",
      "@_bandCol": element.bandCol ? "1" : "0"
    },
    "a:tblGrid": {
      "a:gridCol": ensureArray(element.gridCols).map((w) => ({ "@_w": toInt(w, 0) }))
    },
    "a:tr": ensureArray(element.rows).map((row) => ({
      "@_h": toInt(row.height, 0),
      "a:tc": ensureArray(row.cells).map((cell) => tableCellToNode(cell))
    }))
  };

  if (element.styleId) {
    tbl["a:tblPr"]["@_tableStyleId"] = element.styleId;
  }

  return {
    "p:nvGraphicFramePr": {
      "p:cNvPr": {
        "@_id": toInt(element.id, 0),
        "@_name": element.name || `Table ${element.id}`,
        "@_descr": element.description || ""
      },
      "p:cNvGraphicFramePr": {
        "a:graphicFrameLocks": {
          "@_noGrp": "1"
        }
      },
      "p:nvPr": {}
    },
    "p:xfrm": {
      "a:off": {
        "@_x": toInt(element.x, 0),
        "@_y": toInt(element.y, 0)
      },
      "a:ext": {
        "@_cx": toInt(element.cx, 0),
        "@_cy": toInt(element.cy, 0)
      }
    },
    "a:graphic": {
      "a:graphicData": {
        "@_uri": "http://schemas.openxmlformats.org/drawingml/2006/table",
        "a:tbl": tbl
      }
    }
  };
}

function createBackgroundNode(background) {
  if (!background || background.type !== "solid" || !background.color) {
    return null;
  }

  return {
    "p:bgPr": {
      ...modelColorToOpenXmlFill(background),
      "a:effectLst": {}
    }
  };
}

function defaultSpTreeContainer() {
  return {
    "p:nvGrpSpPr": {
      "p:cNvPr": {
        "@_id": "1",
        "@_name": ""
      },
      "p:cNvGrpSpPr": {},
      "p:nvPr": {}
    },
    "p:grpSpPr": {
      "a:xfrm": {
        "a:off": { "@_x": 0, "@_y": 0 },
        "a:ext": { "@_cx": 0, "@_cy": 0 },
        "a:chOff": { "@_x": 0, "@_y": 0 },
        "a:chExt": { "@_cx": 0, "@_cy": 0 }
      }
    }
  };
}

function appendSpTreeNode(spTree, key, node) {
  if (!key || !node) {
    return;
  }
  if (!spTree[key]) {
    spTree[key] = [];
  }
  if (!Array.isArray(spTree[key])) {
    spTree[key] = [spTree[key]];
  }
  const serialized = JSON.stringify(node);
  const exists = spTree[key].some((existing) => JSON.stringify(existing) === serialized);
  if (!exists) {
    spTree[key].push(node);
  }
}

function nodeElementId(key, node) {
  if (!node || typeof node !== "object") {
    return null;
  }
  if (key === "p:sp") {
    return node?.["p:nvSpPr"]?.["p:cNvPr"]?.["@_id"] ?? null;
  }
  if (key === "p:cxnSp") {
    return node?.["p:nvCxnSpPr"]?.["p:cNvPr"]?.["@_id"] ?? null;
  }
  if (key === "p:pic") {
    return node?.["p:nvPicPr"]?.["p:cNvPr"]?.["@_id"] ?? null;
  }
  if (key === "p:graphicFrame") {
    return node?.["p:nvGraphicFramePr"]?.["p:cNvPr"]?.["@_id"] ?? null;
  }
  return null;
}

function buildManagedNodeMap(spTree, key) {
  const map = new Map();
  for (const node of ensureArray(spTree?.[key])) {
    const id = nodeElementId(key, node);
    if (id === null || id === undefined) {
      continue;
    }
    map.set(String(id), node);
  }
  return map;
}

function patchShapeNode(existingNode, element) {
  const node = deepClone(existingNode || createShapeNode(element));
  node["p:nvSpPr"] = node["p:nvSpPr"] || {};
  node["p:nvSpPr"]["p:cNvPr"] = applyCommonNonVisualProps(node["p:nvSpPr"]["p:cNvPr"], element, "Shape");
  node["p:nvSpPr"]["p:cNvSpPr"] = node["p:nvSpPr"]["p:cNvSpPr"] || {};
  node["p:nvSpPr"]["p:cNvSpPr"]["@_txBox"] = element.type === "text" ? "1" : "0";

  const nvPr = node["p:nvSpPr"]["p:nvPr"] || {};
  if (element.placeholder) {
    nvPr["p:ph"] = {
      "@_type": element.placeholder.type || "body",
      "@_idx": element.placeholder.idx || "0"
    };
  } else {
    delete nvPr["p:ph"];
  }
  node["p:nvSpPr"]["p:nvPr"] = nvPr;

  node["p:spPr"] = buildShapeSpPr(element);
  if (element.text) {
    node["p:txBody"] = buildTextBodyNode(element.text);
  } else {
    delete node["p:txBody"];
  }
  return node;
}

function patchConnectorNode(existingNode, element) {
  const node = deepClone(existingNode || createConnectorNode(element));
  node["p:nvCxnSpPr"] = node["p:nvCxnSpPr"] || {};
  node["p:nvCxnSpPr"]["p:cNvPr"] = applyCommonNonVisualProps(node["p:nvCxnSpPr"]["p:cNvPr"], element, "Connector");
  node["p:nvCxnSpPr"]["p:cNvCxnSpPr"] = node["p:nvCxnSpPr"]["p:cNvCxnSpPr"] || {};
  node["p:nvCxnSpPr"]["p:nvPr"] = node["p:nvCxnSpPr"]["p:nvPr"] || {};
  node["p:spPr"] = buildShapeSpPr({ ...element, shapeType: element.shapeType || "line" });
  return node;
}

function patchPictureNode(existingNode, element) {
  const node = deepClone(existingNode || createPictureNode(element));
  node["p:nvPicPr"] = node["p:nvPicPr"] || {};
  node["p:nvPicPr"]["p:cNvPr"] = applyCommonNonVisualProps(node["p:nvPicPr"]["p:cNvPr"], element, "Picture");
  node["p:nvPicPr"]["p:cNvPicPr"] = node["p:nvPicPr"]["p:cNvPicPr"] || {
    "a:picLocks": {
      "@_noChangeAspect": "1"
    }
  };
  node["p:nvPicPr"]["p:nvPr"] = node["p:nvPicPr"]["p:nvPr"] || {};

  node["p:blipFill"] = node["p:blipFill"] || {};
  node["p:blipFill"]["a:blip"] = node["p:blipFill"]["a:blip"] || {};
  if (element.relId) {
    node["p:blipFill"]["a:blip"]["@_r:embed"] = element.relId;
  }
  node["p:blipFill"]["a:stretch"] = node["p:blipFill"]["a:stretch"] || {
    "a:fillRect": {}
  };

  node["p:spPr"] = node["p:spPr"] || {};
  node["p:spPr"]["a:xfrm"] = buildXfrmNode(element);
  node["p:spPr"]["a:prstGeom"] = node["p:spPr"]["a:prstGeom"] || {
    "@_prst": "rect",
    "a:avLst": {}
  };

  return node;
}

function patchTableNode(existingNode, element) {
  const node = deepClone(existingNode || createTableNode(element));
  const fresh = createTableNode(element);

  node["p:nvGraphicFramePr"] = node["p:nvGraphicFramePr"] || {};
  node["p:nvGraphicFramePr"]["p:cNvPr"] = applyCommonNonVisualProps(
    node["p:nvGraphicFramePr"]["p:cNvPr"],
    element,
    "Table"
  );
  node["p:nvGraphicFramePr"]["p:cNvGraphicFramePr"] = node["p:nvGraphicFramePr"]["p:cNvGraphicFramePr"] || {
    "a:graphicFrameLocks": {
      "@_noGrp": "1"
    }
  };
  node["p:nvGraphicFramePr"]["p:nvPr"] = node["p:nvGraphicFramePr"]["p:nvPr"] || {};

  node["p:xfrm"] = buildXfrmNode(element);
  node["a:graphic"] = fresh["a:graphic"];

  return node;
}

function patchGenericGraphicFrameNode(existingNode, element) {
  const node = deepClone(existingNode || element.raw || null);
  if (!node) {
    return null;
  }

  node["p:nvGraphicFramePr"] = node["p:nvGraphicFramePr"] || {};
  node["p:nvGraphicFramePr"]["p:cNvPr"] = applyCommonNonVisualProps(
    node["p:nvGraphicFramePr"]["p:cNvPr"],
    element,
    "GraphicFrame"
  );
  node["p:nvGraphicFramePr"]["p:cNvGraphicFramePr"] = node["p:nvGraphicFramePr"]["p:cNvGraphicFramePr"] || {};
  node["p:nvGraphicFramePr"]["p:nvPr"] = node["p:nvGraphicFramePr"]["p:nvPr"] || {};
  node["p:xfrm"] = buildXfrmNode(element);
  return node;
}

function buildSpTreeFromModel(slide, sourceSpTree = {}) {
  const spTree = defaultSpTreeContainer();
  if (sourceSpTree?.["p:nvGrpSpPr"]) {
    spTree["p:nvGrpSpPr"] = deepClone(sourceSpTree["p:nvGrpSpPr"]);
  }
  if (sourceSpTree?.["p:grpSpPr"]) {
    spTree["p:grpSpPr"] = deepClone(sourceSpTree["p:grpSpPr"]);
  }

  const sourceShapeMap = buildManagedNodeMap(sourceSpTree, "p:sp");
  const sourceConnectorMap = buildManagedNodeMap(sourceSpTree, "p:cxnSp");
  const sourcePictureMap = buildManagedNodeMap(sourceSpTree, "p:pic");
  const sourceGraphicMap = buildManagedNodeMap(sourceSpTree, "p:graphicFrame");

  const shapes = [];
  const pictures = [];
  const graphics = [];
  const connectors = [];

  for (const element of slide.elements || []) {
    const id = String(element?.id ?? "");
    if (element.type === "shape" || element.type === "text") {
      shapes.push(patchShapeNode(sourceShapeMap.get(id), element));
      continue;
    }
    if (element.type === "line") {
      connectors.push(patchConnectorNode(sourceConnectorMap.get(id), element));
      continue;
    }
    if (element.type === "image") {
      pictures.push(patchPictureNode(sourcePictureMap.get(id), element));
      continue;
    }
    if (element.type === "table") {
      graphics.push(patchTableNode(sourceGraphicMap.get(id), element));
      continue;
    }
    if (element.type === "chart" || element.type === "diagram") {
      const preserved = patchGenericGraphicFrameNode(sourceGraphicMap.get(id), element);
      if (preserved) {
        graphics.push(preserved);
      }
    }
  }

  if (shapes.length) {
    spTree["p:sp"] = shapes;
  }
  if (pictures.length) {
    spTree["p:pic"] = pictures;
  }
  if (graphics.length) {
    spTree["p:graphicFrame"] = graphics;
  }
  if (connectors.length) {
    spTree["p:cxnSp"] = connectors;
  }

  for (const [key, value] of Object.entries(sourceSpTree || {})) {
    if (key === "p:nvGrpSpPr" || key === "p:grpSpPr" || MANAGED_SP_TREE_KEYS.has(key)) {
      continue;
    }
    spTree[key] = deepClone(value);
  }

  for (const raw of slide.unhandledNodes || []) {
    if (!raw?.type || !raw?.node) {
      continue;
    }
    appendSpTreeNode(spTree, raw.type, deepClone(raw.node));
  }

  return spTree;
}

function buildSlideXml(slide) {
  const spTree = buildSpTreeFromModel(slide, {});

  const cSld = {
    "@_name": slide.name || `Slide ${slide.index + 1}`,
    "p:spTree": spTree
  };

  const bg = createBackgroundNode(slide.background);
  if (bg) {
    cSld["p:bg"] = bg;
  }

  return {
    "p:sld": {
      "@_xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "@_xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "@_xmlns:p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "@_xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
      "@_xmlns:a14": "http://schemas.microsoft.com/office/drawing/2010/main",
      "@_mc:Ignorable": "a14",
      "p:cSld": cSld,
      "p:clrMapOvr": {
        "a:masterClrMapping": {}
      }
    }
  };
}

function buildSlideXmlPreservingSource(slide) {
  if (!slide?._sourceXml?.["p:sld"]) {
    return buildSlideXml(slide);
  }

  const rootXml = deepClone(slide._sourceXml);
  const slideRoot = rootXml["p:sld"] || {};
  const cSld = slideRoot["p:cSld"] || {};

  cSld["@_name"] = slide.name || cSld["@_name"] || `Slide ${slide.index + 1}`;
  cSld["p:spTree"] = buildSpTreeFromModel(slide, cSld["p:spTree"] || {});

  const bg = createBackgroundNode(slide.background);
  if (bg) {
    cSld["p:bg"] = bg;
  }

  slideRoot["p:cSld"] = cSld;
  if (!slideRoot["p:clrMapOvr"]) {
    slideRoot["p:clrMapOvr"] = {
      "a:masterClrMapping": {}
    };
  }
  rootXml["p:sld"] = slideRoot;

  return rootXml;
}

function buildRelationshipsXml(relationships) {
  return {
    Relationships: {
      "@_xmlns": "http://schemas.openxmlformats.org/package/2006/relationships",
      Relationship: relationships.map((rel) => {
        const relationshipNode = {
          "@_Id": rel.id,
          "@_Type": rel.type,
          "@_Target": rel.target
        };
        if (rel.targetMode === "External") {
          relationshipNode["@_TargetMode"] = "External";
        }
        return relationshipNode;
      })
    }
  };
}

function upsertImageRelationship(slide, slidePath, relationships, usedIds) {
  let counter = 1000;

  for (const element of slide.elements || []) {
    if (element.type !== "image") {
      continue;
    }

    let targetPart = element.imagePath;
    if (!targetPart && element.dataUri) {
      const ext = mapMimeToExt(element.mimeType);
      targetPart = `ppt/media/generated-${slide.index + 1}-${element.id}.${ext}`;
      element.imagePath = targetPart;
    }

    if (!targetPart) {
      continue;
    }

    let relId = element.relId;
    if (!relId || usedIds.has(relId)) {
      do {
        relId = `rId${counter}`;
        counter += 1;
      } while (usedIds.has(relId));
      element.relId = relId;
    }

    usedIds.add(relId);

    const existing = relationships.find((r) => r.id === relId);
    const target = relativeTarget(slidePath, targetPart);

    if (existing) {
      existing.type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
      existing.target = target;
      existing.targetMode = "Internal";
    } else {
      relationships.push({
        id: relId,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        target,
        targetMode: "Internal"
      });
    }
  }
}

export async function writeModelToPptx(openXmlPackage, model, options = {}) {
  const outputPackage = options.inPlace ? openXmlPackage : await openXmlPackage.clone();

  for (const slide of model.slides || []) {
    const snapshot = slide?._snapshot || null;
    const unchanged = !options.forceRewrite && snapshot && snapshot === createSlideSnapshot(slide);
    if (unchanged) {
      continue;
    }

    const slidePath = slide.sourcePath || `ppt/slides/slide${slide.index + 1}.xml`;
    const preserveSource = options.preserveSourceXml !== false;
    const slideXml = preserveSource
      ? buildSlideXmlPreservingSource(slide)
      : buildSlideXml(slide);
    outputPackage.writeXml(slidePath, slideXml);

    const relationships = ensureArray(slide.sourceRelationships).map((rel) => ({
      id: rel.id,
      type: rel.type,
      target: rel.target,
      targetMode: rel.targetMode || "Internal"
    }));

    const usedIds = new Set(relationships.map((r) => r.id));
    upsertImageRelationship(slide, slidePath, relationships, usedIds);

    for (const element of slide.elements || []) {
      if (element.type !== "image" || !element.imagePath || !element.dataUri) {
        continue;
      }
      const binary = decodeDataUri(element.dataUri);
      if (binary) {
        outputPackage.writeBinary(element.imagePath, binary);
      }
    }

    if (relationships.length) {
      outputPackage.writeXml(slide.sourceRelsPath || `ppt/slides/_rels/slide${slide.index + 1}.xml.rels`, buildRelationshipsXml(relationships));
    }
  }

  return outputPackage;
}
