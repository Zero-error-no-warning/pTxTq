import { ensureArray } from "../utils/object.js";
import { resolveTargetPath } from "../utils/xml.js";

const REL_TYPE_SLIDE_LAYOUT = "/relationships/slideLayout";
const REL_TYPE_SLIDE_MASTER = "/relationships/slideMaster";
const REL_TYPE_THEME = "/relationships/theme";

function relationshipByType(relsMap, typeSuffix) {
  for (const rel of relsMap.values()) {
    if ((rel.type || "").endsWith(typeSuffix)) {
      return rel;
    }
  }
  return null;
}

async function getPartWithRels(openXmlPackage, partPath, cache) {
  if (!partPath) {
    return { xml: null, rels: new Map() };
  }
  if (!cache.has(partPath)) {
    cache.set(partPath, {
      xml: await openXmlPackage.readXml(partPath),
      rels: await openXmlPackage.getRelationships(partPath)
    });
  }
  return cache.get(partPath);
}

/**
 * Build the PresentationML part graph according to the Office Open XML structure:
 * presentation.xml -> slideX.xml -> slideLayoutX.xml -> slideMasterX.xml -> themeX.xml
 * Reference (user-requested): http://officeopenxml.com/prPresentation.php
 */
export async function buildPresentationGraph(openXmlPackage) {
  const presentationPath = "ppt/presentation.xml";
  const presentationXml = await openXmlPackage.readXml(presentationPath);
  const presentationRoot = presentationXml?.["p:presentation"] || {};
  const presentationRels = await openXmlPackage.getRelationships(presentationPath);
  const slideIdNodes = ensureArray(presentationRoot?.["p:sldIdLst"]?.["p:sldId"]);

  const slideCache = new Map();
  const layoutCache = new Map();
  const masterCache = new Map();
  const themeCache = new Map();

  const slides = [];

  for (let i = 0; i < slideIdNodes.length; i += 1) {
    const slideIdNode = slideIdNodes[i];
    const relId = slideIdNode?.["@_r:id"];
    if (!relId || !presentationRels.has(relId)) {
      continue;
    }

    const slideRel = presentationRels.get(relId);
    const slidePath = resolveTargetPath(presentationPath, slideRel.target);
    const slidePart = await getPartWithRels(openXmlPackage, slidePath, slideCache);
    const slideXml = slidePart.xml;
    const slideRels = slidePart.rels;

    const layoutRel = relationshipByType(slideRels, REL_TYPE_SLIDE_LAYOUT);
    const layoutPath = layoutRel ? resolveTargetPath(slidePath, layoutRel.target) : null;
    const layoutPart = await getPartWithRels(openXmlPackage, layoutPath, layoutCache);
    const layoutXml = layoutPart.xml;
    const layoutRels = layoutPart.rels;

    const masterRel = relationshipByType(layoutRels, REL_TYPE_SLIDE_MASTER);
    const masterPath = masterRel && layoutPath ? resolveTargetPath(layoutPath, masterRel.target) : null;
    const masterPart = await getPartWithRels(openXmlPackage, masterPath, masterCache);
    const masterXml = masterPart.xml;
    const masterRels = masterPart.rels;

    const themeRel = relationshipByType(masterRels, REL_TYPE_THEME);
    const themePath = themeRel && masterPath ? resolveTargetPath(masterPath, themeRel.target) : null;
    const themePart = await getPartWithRels(openXmlPackage, themePath, themeCache);

    slides.push({
      index: i,
      slideIdNode,
      relId,
      slidePath,
      slideXml,
      slideRels,
      layoutPath,
      layoutXml,
      layoutRels,
      masterPath,
      masterXml,
      masterRels,
      themePath,
      themeXml: themePart.xml,
      themeRels: themePart.rels
    });
  }

  return {
    presentationPath,
    presentationXml,
    presentationRoot,
    presentationRels,
    slides
  };
}

