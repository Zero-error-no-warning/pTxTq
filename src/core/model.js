
import {
  alphaToPct,
  resolveDrawingColor,
  rgbHexNoPrefix
} from "../utils/color.js";
import { deg60000ToDeg, centipointsToPt } from "../utils/units.js";
import {
  clamp,
  ensureArray,
  first,
  toInt,
  deepClone
} from "../utils/object.js";
import {
  contentTypeMap,
  detectImageMimeByExt,
  parseXmlPreserveOrder,
  relsPartPath,
  resolveTargetPath,
  uint8ToBase64
} from "../utils/xml.js";
import {
  createThemeContext,
  mergeColorMap,
  parseColorMapFromMaster,
  parseColorMapOverride,
  parseTheme
} from "./theme.js";
import { buildPresentationGraph } from "./presentationGraph.js";

function normalizeBool(value) {
  if (value === undefined || value === null) {
    return false;
  }
  const normalized = String(value).toLowerCase();
  return value === true || value === "1" || normalized === "true" || normalized === "on" || normalized === "yes";
}

function hasNode(node, key) {
  return Boolean(node && typeof node === "object" && Object.prototype.hasOwnProperty.call(node, key));
}

function pickTextColor(node, themeContext, fallback = "#000000") {
  const solid = node?.["a:solidFill"];
  if (!solid) {
    return fallback;
  }
  const resolved = resolveDrawingColor(first(solid), themeContext);
  if (!resolved?.color) {
    return fallback;
  }
  return resolved.color;
}

function parseTransform(xfrmNode) {
  const off = xfrmNode?.["a:off"] || xfrmNode?.["p:off"];
  const ext = xfrmNode?.["a:ext"] || xfrmNode?.["p:ext"];
  return {
    x: toInt(off?.["@_x"], 0),
    y: toInt(off?.["@_y"], 0),
    cx: toInt(ext?.["@_cx"], 0),
    cy: toInt(ext?.["@_cy"], 0),
    rotation: deg60000ToDeg(toInt(xfrmNode?.["@_rot"], 0)),
    flipH: normalizeBool(xfrmNode?.["@_flipH"]),
    flipV: normalizeBool(xfrmNode?.["@_flipV"])
  };
}

function placeholderInfo(shapeNode) {
  const ph = shapeNode?.["p:nvSpPr"]?.["p:nvPr"]?.["p:ph"];
  if (!ph) {
    return null;
  }
  return {
    type: ph?.["@_type"] || "body",
    idx: ph?.["@_idx"] || "0",
    orient: ph?.["@_orient"] || null,
    sz: ph?.["@_sz"] || null
  };
}

function normalizePlaceholderType(type) {
  return String(type || "body").toLowerCase();
}

function buildPlaceholderMap(spTree) {
  const entries = [];
  const shapes = ensureArray(spTree?.["p:sp"]);
  for (const shape of shapes) {
    const ph = placeholderInfo(shape);
    if (!ph) {
      continue;
    }
    entries.push({ ph, node: shape });
  }
  return entries;
}

function placeholderTypeCompatible(targetType, candidateType) {
  const target = normalizePlaceholderType(targetType);
  const candidate = normalizePlaceholderType(candidateType);
  if (target === candidate) {
    return true;
  }

  const titleTypes = new Set(["title", "ctrtitle"]);
  if (titleTypes.has(target) && titleTypes.has(candidate)) {
    return true;
  }

  const bodyTypes = new Set(["body", "obj"]);
  if (bodyTypes.has(target) && bodyTypes.has(candidate)) {
    return true;
  }

  return false;
}

function nodeHasTransform(node) {
  return Boolean(node?.["p:spPr"]?.["a:xfrm"]);
}

function placeholderMatchScore(targetPh, candidatePh, candidateNode) {
  const targetIdx = String(targetPh?.idx ?? "0");
  const candidateIdx = String(candidatePh?.idx ?? "0");
  const typeCompatible = placeholderTypeCompatible(targetPh?.type, candidatePh?.type);
  const idxMatch = targetIdx === candidateIdx;

  if (!typeCompatible && !idxMatch) {
    return -1;
  }

  let score = 0;
  if (typeCompatible) {
    score += 200;
  }
  if (idxMatch) {
    score += 120;
  }
  if (typeCompatible && idxMatch) {
    score += 100;
  }
  if (nodeHasTransform(candidateNode)) {
    score += 5;
  }
  if (candidateNode?.["p:txBody"]) {
    score += 1;
  }

  return score;
}

function findBestPlaceholder(placeholderEntries, targetPh) {
  let bestNode = null;
  let bestScore = -1;
  for (const entry of ensureArray(placeholderEntries)) {
    const score = placeholderMatchScore(targetPh, entry?.ph, entry?.node);
    if (score > bestScore) {
      bestScore = score;
      bestNode = entry.node;
    }
  }
  return bestNode;
}

function resolveInheritedShape(shapeNode, layoutPlaceholderMap, masterPlaceholderMap) {
  const ph = placeholderInfo(shapeNode);
  if (!ph) {
    return null;
  }

  const layoutShape = findBestPlaceholder(layoutPlaceholderMap, ph);
  const masterFromSlide = findBestPlaceholder(masterPlaceholderMap, ph);
  const masterFromLayout = layoutShape
    ? findBestPlaceholder(masterPlaceholderMap, placeholderInfo(layoutShape) || ph)
    : null;
  const masterShape = masterFromLayout || masterFromSlide;

  if (!layoutShape && !masterShape) {
    return null;
  }
  return deepMergeNodes(masterShape, layoutShape);
}

function deepMergeNodes(baseNode, overrideNode) {
  if (overrideNode === "") {
    return deepClone(baseNode);
  }
  if (baseNode === "") {
    return deepClone(overrideNode);
  }
  if (!baseNode) {
    return deepClone(overrideNode);
  }
  if (!overrideNode) {
    return deepClone(baseNode);
  }
  if (Array.isArray(baseNode) || Array.isArray(overrideNode)) {
    return deepClone(overrideNode);
  }
  if (typeof baseNode !== "object" || typeof overrideNode !== "object") {
    return deepClone(overrideNode);
  }

  const merged = deepClone(baseNode);
  for (const [key, value] of Object.entries(overrideNode)) {
    if (value === undefined) {
      continue;
    }
    if (merged[key] === undefined) {
      merged[key] = deepClone(value);
      continue;
    }
    merged[key] = deepMergeNodes(merged[key], value);
  }
  return merged;
}

function resolveFillFromRef(styleNode, themeContext) {
  const fillRef = styleNode?.["a:fillRef"];
  if (!fillRef) {
    return null;
  }

  const idx = toInt(fillRef?.["@_idx"], 0);
  const styleEntry = themeContext?.getFillStyle?.(idx);

  const referenceColor = resolveDrawingColor(fillRef, themeContext);
  if (styleEntry?.type === "a:solidFill") {
    const resolved = resolveDrawingColor(styleEntry.node, themeContext, referenceColor);
    if (resolved) {
      return {
        type: "solid",
        color: resolved.color,
        alpha: resolved.alpha,
        source: "theme-fill-style"
      };
    }
  }

  if (styleEntry?.type === "a:gradFill") {
    const gsLst = styleEntry.node?.["a:gsLst"];
    const stops = ensureArray(gsLst?.["a:gs"]);
    if (stops.length) {
      const resolved = resolveDrawingColor(first(stops), themeContext, referenceColor);
      if (resolved) {
        return {
          type: "solid",
          color: resolved.color,
          alpha: resolved.alpha,
          source: "theme-grad-fallback"
        };
      }
    }
  }

  if (referenceColor) {
    return {
      type: "solid",
      color: referenceColor.color,
      alpha: referenceColor.alpha,
      source: "theme-fill-ref"
    };
  }

  return null;
}

function resolveFill(spPrNode, styleNode, themeContext) {
  if (hasNode(spPrNode, "a:noFill")) {
    return { type: "none", color: null, alpha: 0, source: "shape" };
  }

  if (hasNode(spPrNode, "a:solidFill")) {
    const resolved = resolveDrawingColor(first(spPrNode["a:solidFill"]), themeContext);
    if (resolved) {
      return {
        type: "solid",
        color: resolved.color,
        alpha: resolved.alpha,
        source: "shape"
      };
    }
  }

  if (spPrNode?.["a:gradFill"]) {
    const grad = first(spPrNode["a:gradFill"]);
    const stops = ensureArray(grad?.["a:gsLst"]?.["a:gs"]);
    if (stops.length) {
      const resolved = resolveDrawingColor(first(stops), themeContext);
      if (resolved) {
        return {
          type: "solid",
          color: resolved.color,
          alpha: resolved.alpha,
          source: "shape-grad-fallback"
        };
      }
    }
  }

  if (spPrNode?.["a:pattFill"]) {
    const patt = first(spPrNode["a:pattFill"]);
    const resolved = resolveDrawingColor({ "a:schemeClr": patt?.["a:fgClr"]?.["a:schemeClr"] }, themeContext)
      || resolveDrawingColor({ "a:srgbClr": patt?.["a:fgClr"]?.["a:srgbClr"] }, themeContext);
    if (resolved) {
      return {
        type: "solid",
        color: resolved.color,
        alpha: resolved.alpha,
        source: "shape-pattern-fallback"
      };
    }
  }

  return resolveFillFromRef(styleNode, themeContext);
}

function parseLineEndNode(endNode) {
  const node = first(endNode);
  if (!node) {
    return undefined;
  }
  return {
    type: String(node?.["@_type"] || "none").toLowerCase(),
    width: String(node?.["@_w"] || "med").toLowerCase(),
    length: String(node?.["@_len"] || "med").toLowerCase()
  };
}

function normalizeLineEnd(end, fallback = null) {
  if (end === undefined) {
    return fallback;
  }
  if (!end || end.type === "none") {
    return null;
  }
  return {
    type: String(end.type || "none").toLowerCase(),
    width: String(end.width || "med").toLowerCase(),
    length: String(end.length || "med").toLowerCase()
  };
}

function parseCustomDash(custDashNode) {
  const node = first(custDashNode);
  if (!node) {
    return null;
  }
  const dashStops = ensureArray(node?.["a:ds"]).map((stop) => ({
    d: toInt(stop?.["@_d"], 0),
    sp: toInt(stop?.["@_sp"], 0)
  })).filter((stop) => stop.d > 0 || stop.sp > 0);
  return dashStops.length ? dashStops : null;
}

function resolveDashNode(lnNode, fallbackDash = "solid", fallbackCustomDash = null) {
  if (!lnNode) {
    return {
      dash: fallbackDash || "solid",
      customDash: fallbackCustomDash || null
    };
  }

  if (lnNode?.["a:custDash"]) {
    return {
      dash: "cust",
      customDash: parseCustomDash(lnNode?.["a:custDash"])
    };
  }

  const preset = lnNode?.["a:prstDash"]?.["@_val"];
  if (preset) {
    return {
      dash: String(preset).toLowerCase(),
      customDash: null
    };
  }

  return {
    dash: fallbackDash || "solid",
    customDash: fallbackCustomDash || null
  };
}

function resolveLineFromRef(styleNode, themeContext) {
  const lnRef = styleNode?.["a:lnRef"];
  if (!lnRef) {
    return null;
  }
  const idx = toInt(lnRef?.["@_idx"], 0);
  const lineStyle = themeContext?.getLineStyle?.(idx);
  const refColor = resolveDrawingColor(lnRef, themeContext);

  if (lineStyle) {
    const styleLn = first(lineStyle);
    const width = toInt(styleLn?.["@_w"], 0);
    const direct = resolveDrawingColor(styleLn?.["a:solidFill"], themeContext, refColor) || refColor;
    const dashInfo = resolveDashNode(styleLn, "solid");
    if (direct?.color) {
      return {
        width,
        color: direct.color,
        alpha: direct.alpha,
        dash: dashInfo.dash,
        customDash: dashInfo.customDash,
        cap: styleLn?.["@_cap"] || "flat",
        headEnd: normalizeLineEnd(parseLineEndNode(styleLn?.["a:headEnd"])),
        tailEnd: normalizeLineEnd(parseLineEndNode(styleLn?.["a:tailEnd"])),
        source: "theme-line-style"
      };
    }
  }

  if (refColor) {
    return {
      width: 0,
      color: refColor.color,
      alpha: refColor.alpha,
      dash: "solid",
      customDash: null,
      cap: "flat",
      headEnd: null,
      tailEnd: null,
      source: "theme-line-ref"
    };
  }

  return null;
}

function resolveLine(spPrNode, styleNode, themeContext) {
  const lineNode = spPrNode?.["a:ln"];
  const fallback = resolveLineFromRef(styleNode, themeContext);
  if (lineNode) {
    const ln = first(lineNode);
    if (hasNode(ln, "a:noFill")) {
      return {
        width: toInt(ln?.["@_w"], fallback?.width ?? 0),
        color: null,
        alpha: 0,
        dash: "none",
        customDash: null,
        cap: ln?.["@_cap"] || fallback?.cap || "flat",
        headEnd: normalizeLineEnd(parseLineEndNode(ln?.["a:headEnd"]), fallback?.headEnd || null),
        tailEnd: normalizeLineEnd(parseLineEndNode(ln?.["a:tailEnd"]), fallback?.tailEnd || null),
        source: "shape"
      };
    }
    const color = resolveDrawingColor(ln?.["a:solidFill"], themeContext);
    const dashInfo = resolveDashNode(ln, fallback?.dash || "solid", fallback?.customDash || null);
    return {
      width: toInt(ln?.["@_w"], fallback?.width ?? 0),
      color: color?.color || fallback?.color || null,
      alpha: color?.alpha ?? fallback?.alpha ?? 1,
      dash: dashInfo.dash,
      customDash: dashInfo.customDash,
      cap: ln?.["@_cap"] || fallback?.cap || "flat",
      headEnd: normalizeLineEnd(parseLineEndNode(ln?.["a:headEnd"]), fallback?.headEnd || null),
      tailEnd: normalizeLineEnd(parseLineEndNode(ln?.["a:tailEnd"]), fallback?.tailEnd || null),
      source: "shape"
    };
  }

  return fallback;
}
function parseRunStyle(rPrNode, themeContext, fallbackColor, fallbackFont, defaultRPrNode = null) {
  const merged = deepMergeNodes(defaultRPrNode || {}, rPrNode || {}) || {};
  const color = pickTextColor(merged, themeContext, fallbackColor);
  const hasSolidFill = hasNode(merged, "a:solidFill");
  const hyperlinkColor = themeContext?.resolveSchemeColor?.("hlink") || null;
  const lang = merged?.["@_lang"] || null;
  const latinTypeface = merged?.["a:latin"]?.["@_typeface"];
  const eaTypeface = merged?.["a:ea"]?.["@_typeface"];
  const csTypeface = merged?.["a:cs"]?.["@_typeface"];

  const resolvedLatin = themeContext?.resolveThemeFont?.(latinTypeface, lang) || latinTypeface;
  const resolvedEa = themeContext?.resolveThemeFont?.(eaTypeface, lang) || eaTypeface;
  const resolvedCs = themeContext?.resolveThemeFont?.(csTypeface, lang) || csTypeface;
  const eastAsianLang = /^(ja|zh|ko)/i.test(String(lang || ""));
  const themeEastAsiaFallback = eastAsianLang
    ? (themeContext?.resolveThemeFont?.("+mn-ea", lang)
      || themeContext?.resolveThemeFont?.("+mj-ea", lang)
      || null)
    : null;
  const preferredFont = eastAsianLang
    ? (resolvedEa || themeEastAsiaFallback || resolvedLatin || fallbackFont || themeContext?.theme?.fontScheme?.minor?.latin || "Calibri")
    : (resolvedLatin || resolvedEa || fallbackFont || themeContext?.theme?.fontScheme?.minor?.latin || "Calibri");

  const style = {
    fontSizePt: centipointsToPt(toInt(merged?.["@_sz"], 1800)),
    bold: normalizeBool(merged?.["@_b"]),
    italic: normalizeBool(merged?.["@_i"]),
    underline: (merged?.["@_u"] || "none") !== "none",
    strike: (merged?.["@_strike"] || "noStrike") !== "noStrike",
    kerning: toInt(merged?.["@_kern"], 0),
    baseline: toInt(merged?.["@_baseline"], 0),
    caps: merged?.["@_cap"] || "none",
    spacing: toInt(merged?.["@_spc"], 0),
    lang,
    color,
    explicitColor: hasSolidFill,
    alpha: resolveDrawingColor(merged?.["a:solidFill"], themeContext)?.alpha ?? 1,
    fontFamily: preferredFont,
    eastAsiaFont: resolvedEa || null,
    complexScriptFont: resolvedCs || null
  };

  if (hasNode(merged, "a:hlinkClick")) {
    style.underline = true;
    if (!hasSolidFill && hyperlinkColor) {
      style.color = hyperlinkColor;
    }
  }

  return style;
}

function parseParagraphBullet(pPrNode, themeContext) {
  if (!pPrNode || hasNode(pPrNode, "a:buNone")) {
    return null;
  }

  if (hasNode(pPrNode, "a:buAutoNum")) {
    const auto = first(pPrNode?.["a:buAutoNum"]) || {};
    return {
      type: "autoNum",
      format: auto?.["@_type"] || "arabicPeriod",
      startAt: toInt(auto?.["@_startAt"], 1)
    };
  }

  const buCharNode = first(pPrNode?.["a:buChar"]);
  if (buCharNode) {
    const buClrNode = first(pPrNode?.["a:buClr"]);
    const buSzPctNode = first(pPrNode?.["a:buSzPct"]);
    const buSzPtsNode = first(pPrNode?.["a:buSzPts"]);
    const buFontNode = first(pPrNode?.["a:buFont"]);
    const color = resolveDrawingColor(buClrNode, themeContext);
    return {
      type: "char",
      char: buCharNode?.["@_char"] || "\u2022",
      fontFamily: buFontNode?.["@_typeface"] || null,
      color: color?.color || null,
      alpha: color?.alpha ?? 1,
      sizePct: buSzPctNode ? toInt(buSzPctNode?.["@_val"], 100000) / 100000 : null,
      sizePt: buSzPtsNode ? centipointsToPt(toInt(buSzPtsNode?.["@_val"], 0)) : null
    };
  }

  return null;
}

function parseBreakRuns(
  paragraphNode,
  themeContext,
  defaultColor,
  defaultFont,
  paragraphDefaultRPr,
  endRPr
) {
  return ensureArray(paragraphNode?.["a:br"]).map((brNode) => ({
    text: "\n",
    style: parseRunStyle(brNode?.["a:rPr"] || endRPr, themeContext, defaultColor, defaultFont, paragraphDefaultRPr)
  }));
}

function insertBreakRuns(runs, breakRuns) {
  if (!breakRuns.length) {
    return runs;
  }
  if (!runs.length) {
    return [...breakRuns];
  }
  if (runs.length === 1) {
    return [...runs, ...breakRuns];
  }
  const insertIndex = Math.max(1, runs.length - 1);
  return [
    ...runs.slice(0, insertIndex),
    ...breakRuns,
    ...runs.slice(insertIndex)
  ];
}

function parseTextParagraph(
  paragraphNode,
  themeContext,
  inheritedStyle = {},
  paragraphDefaultRPr = null,
  paragraphDefaultPPr = null
) {
  const pPr = deepMergeNodes(paragraphDefaultPPr || {}, paragraphNode?.["a:pPr"] || {}) || {};
  const endRPr = paragraphNode?.["a:endParaRPr"] || {};
  const defaultColor = inheritedStyle.color || "#000000";
  const defaultFont = inheritedStyle.fontFamily || themeContext?.theme?.fontScheme?.minor?.latin || "Calibri";
  const paraDefaultRPr = deepMergeNodes(paragraphDefaultRPr || {}, pPr?.["a:defRPr"] || {}) || {};

  const runs = [];

  for (const runNode of ensureArray(paragraphNode?.["a:r"])) {
    const runStyle = parseRunStyle(runNode?.["a:rPr"] || {}, themeContext, defaultColor, defaultFont, paraDefaultRPr);
    runs.push({
      text: runNode?.["a:t"] || "",
      style: runStyle
    });
  }

  for (const fieldNode of ensureArray(paragraphNode?.["a:fld"])) {
    const runStyle = parseRunStyle(fieldNode?.["a:rPr"] || {}, themeContext, defaultColor, defaultFont, paraDefaultRPr);
    runs.push({
      text: fieldNode?.["a:t"] || "",
      style: runStyle,
      fieldType: fieldNode?.["@_type"] || null,
      fieldId: fieldNode?.["@_id"] || null
    });
  }

  if (!runs.length && paragraphNode?.["a:t"]) {
    runs.push({
      text: paragraphNode["a:t"],
      style: parseRunStyle(endRPr, themeContext, defaultColor, defaultFont, paraDefaultRPr)
    });
  }

  if (!runs.length) {
    runs.push({
      text: "",
      style: parseRunStyle(endRPr, themeContext, defaultColor, defaultFont, paraDefaultRPr)
    });
  }

  const breakRuns = parseBreakRuns(
    paragraphNode,
    themeContext,
    defaultColor,
    defaultFont,
    paraDefaultRPr,
    endRPr
  );
  const runsWithBreaks = insertBreakRuns(runs, breakRuns);

  return {
    alignment: pPr?.["@_algn"] || "l",
    level: toInt(pPr?.["@_lvl"], 0),
    marginLeft: toInt(pPr?.["@_marL"], 0),
    marginRight: toInt(pPr?.["@_marR"], 0),
    indent: toInt(pPr?.["@_indent"], 0),
    defaultTabSize: toInt(pPr?.["@_defTabSz"], 0),
    rtl: normalizeBool(pPr?.["@_rtl"]),
    eaLineBreak: normalizeBool(pPr?.["@_eaLnBrk"]),
    latinLineBreak: normalizeBool(pPr?.["@_latinLnBrk"]),
    hangingPunctuation: normalizeBool(pPr?.["@_hangingPunct"]),
    spaceBefore: toInt(pPr?.["a:spcBef"]?.["a:spcPts"]?.["@_val"], 0),
    spaceAfter: toInt(pPr?.["a:spcAft"]?.["a:spcPts"]?.["@_val"], 0),
    lineSpacing: toInt(pPr?.["a:lnSpc"]?.["a:spcPct"]?.["@_val"], 0),
    lineSpacingPt: centipointsToPt(toInt(pPr?.["a:lnSpc"]?.["a:spcPts"]?.["@_val"], 0)),
    bullet: parseParagraphBullet(pPr, themeContext),
    runs: runsWithBreaks
  };
}

function parseTextBody(txBodyNode, themeContext, inheritedStyle = {}, fallbackLstStyle = null) {
  if (!txBodyNode) {
    return null;
  }

  const bodyPr = txBodyNode?.["a:bodyPr"] || {};
  const lstStyle = deepMergeNodes(fallbackLstStyle || {}, txBodyNode?.["a:lstStyle"] || {}) || {};
  const defPPr = lstStyle?.["a:defPPr"] || {};
  const defRPr = defPPr?.["a:defRPr"] || {};

  const defaultPropsByLevel = (level) => {
    const normalized = Math.max(0, Math.min(8, toInt(level, 0)));
    const levelKey = `a:lvl${normalized + 1}pPr`;
    const lvlPPr = lstStyle?.[levelKey] || {};
    const mergedPPr = deepMergeNodes(defPPr || {}, lvlPPr || {}) || {};
    const mergedRPr = deepMergeNodes(defRPr || {}, lvlPPr?.["a:defRPr"] || {}) || {};
    return {
      pPr: mergedPPr,
      rPr: mergedRPr
    };
  };

  const paragraphs = ensureArray(txBodyNode?.["a:p"]).map((paragraph) => {
    const level = toInt(paragraph?.["a:pPr"]?.["@_lvl"], 0);
    const defaults = defaultPropsByLevel(level);
    return parseTextParagraph(paragraph, themeContext, inheritedStyle, defaults.rPr, defaults.pPr);
  });

  return {
    direction: bodyPr?.["@_vert"] || "horz",
    rotation: toInt(bodyPr?.["@_rot"], 0),
    verticalAlign: bodyPr?.["@_anchor"] || "t",
    wrap: bodyPr?.["@_wrap"] || "square",
    rtlCol: normalizeBool(bodyPr?.["@_rtlCol"]),
    fromWordArt: normalizeBool(bodyPr?.["@_fromWordArt"]),
    anchorCtr: normalizeBool(bodyPr?.["@_anchorCtr"]),
    forceAA: normalizeBool(bodyPr?.["@_forceAA"]),
    upright: normalizeBool(bodyPr?.["@_upright"]),
    numCol: toInt(bodyPr?.["@_numCol"], 1),
    leftInset: toInt(bodyPr?.["@_lIns"], 45720),
    topInset: toInt(bodyPr?.["@_tIns"], 22860),
    rightInset: toInt(bodyPr?.["@_rIns"], 45720),
    bottomInset: toInt(bodyPr?.["@_bIns"], 22860),
    paragraphs
  };
}

function resolveTextStyleForPlaceholder(placeholder, masterTextStyles) {
  if (!masterTextStyles || !placeholder) {
    return null;
  }
  const phType = normalizePlaceholderType(placeholder?.type || "body");
  if (phType === "title" || phType === "ctrtitle") {
    return masterTextStyles.title || masterTextStyles.other || masterTextStyles.body;
  }
  if (phType === "body" || phType === "obj" || phType === "subbody") {
    return masterTextStyles.body || masterTextStyles.other || masterTextStyles.title;
  }
  return masterTextStyles.other || masterTextStyles.body || masterTextStyles.title;
}

function parseGuideList(gdListNode) {
  return ensureArray(gdListNode?.["a:gd"]).map((gd) => ({
    name: gd?.["@_name"] || null,
    fmla: gd?.["@_fmla"] || null
  })).filter((gd) => gd.name);
}

function parsePointExpr(ptNode) {
  const pt = first(ptNode) || {};
  return {
    x: String(pt?.["@_x"] ?? "0"),
    y: String(pt?.["@_y"] ?? "0")
  };
}

function parsePathCommandList(pathNode) {
  const commands = [];
  const source = pathNode || {};

  for (const [key, value] of Object.entries(source)) {
    if (key.startsWith("@_")) {
      continue;
    }
    const nodes = ensureArray(value);
    switch (key) {
      case "a:moveTo":
        for (const node of nodes) {
          const pt = parsePointExpr(node?.["a:pt"]);
          commands.push({ type: "moveTo", x: pt.x, y: pt.y });
        }
        break;
      case "a:lnTo":
        for (const node of nodes) {
          const pt = parsePointExpr(node?.["a:pt"]);
          commands.push({ type: "lnTo", x: pt.x, y: pt.y });
        }
        break;
      case "a:arcTo":
        for (const node of nodes) {
          const arc = first(node) || node || {};
          commands.push({
            type: "arcTo",
            wR: String(arc?.["@_wR"] ?? "0"),
            hR: String(arc?.["@_hR"] ?? "0"),
            stAng: String(arc?.["@_stAng"] ?? "0"),
            swAng: String(arc?.["@_swAng"] ?? "0")
          });
        }
        break;
      case "a:quadBezTo":
        for (const node of nodes) {
          const points = ensureArray(node?.["a:pt"]).map((ptNode) => parsePointExpr(ptNode));
          if (points.length >= 2) {
            commands.push({
              type: "quadBezTo",
              points: points.slice(0, 2)
            });
          }
        }
        break;
      case "a:cubicBezTo":
        for (const node of nodes) {
          const points = ensureArray(node?.["a:pt"]).map((ptNode) => parsePointExpr(ptNode));
          if (points.length >= 3) {
            commands.push({
              type: "cubicBezTo",
              points: points.slice(0, 3)
            });
          }
        }
        break;
      case "a:close":
        for (const _ of nodes) {
          commands.push({ type: "close" });
        }
        break;
      default:
        break;
    }
  }

  return commands;
}

function parseCustomGeometry(spPrNode) {
  const custGeom = first(spPrNode?.["a:custGeom"]);
  if (!custGeom) {
    return null;
  }

  const pathLst = first(custGeom?.["a:pathLst"]) || {};
  const pathDefaults = {
    w: toInt(pathLst?.["@_w"], 21600),
    h: toInt(pathLst?.["@_h"], 21600)
  };

  const paths = ensureArray(pathLst?.["a:path"]).map((pathNode) => {
    const path = first(pathNode) || pathNode || {};
    return {
      w: toInt(path?.["@_w"], pathDefaults.w || 21600),
      h: toInt(path?.["@_h"], pathDefaults.h || 21600),
      fill: path?.["@_fill"] || "norm",
      stroke: path?.["@_stroke"] ?? "1",
      extrusionOk: path?.["@_extrusionOk"] ?? "1",
      commands: parsePathCommandList(path)
    };
  });

  return {
    kind: "cust",
    preset: "custom",
    adjustValues: parseGuideList(custGeom?.["a:avLst"]),
    guideValues: parseGuideList(custGeom?.["a:gdLst"]),
    pathDefaults,
    paths,
    raw: deepClone(custGeom)
  };
}

function parseGeometry(spPrNode) {
  const custom = parseCustomGeometry(spPrNode);
  if (custom) {
    return custom;
  }

  const preset = spPrNode?.["a:prstGeom"]?.["@_prst"] || "rect";
  return {
    kind: "prst",
    preset,
    adjustValues: ensureArray(spPrNode?.["a:prstGeom"]?.["a:avLst"]?.["a:gd"]).map((gd) => ({
      name: gd?.["@_name"],
      fmla: gd?.["@_fmla"]
    })),
    raw: null
  };
}

function parseShapeCommon(shapeNode, effectiveSpPrNode, effectiveStyleNode, themeContext) {
  const cNvPr = shapeNode?.["p:nvSpPr"]?.["p:cNvPr"] || shapeNode?.["p:nvPicPr"]?.["p:cNvPr"];
  const transformNode = effectiveSpPrNode?.["a:xfrm"] || shapeNode?.["p:xfrm"] || shapeNode?.["a:xfrm"];
  const transform = parseTransform(transformNode || {});

  const fill = resolveFill(effectiveSpPrNode, effectiveStyleNode, themeContext);
  const line = resolveLine(effectiveSpPrNode, effectiveStyleNode, themeContext);

  return {
    id: String(cNvPr?.["@_id"] || "0"),
    name: cNvPr?.["@_name"] || "",
    description: cNvPr?.["@_descr"] || "",
    hidden: normalizeBool(cNvPr?.["@_hidden"]),
    x: transform.x,
    y: transform.y,
    cx: transform.cx,
    cy: transform.cy,
    rotation: transform.rotation,
    flipH: transform.flipH,
    flipV: transform.flipV,
    fill,
    line
  };
}

function placeholderSuppresssInheritedParagraphs(placeholder) {
  return Boolean(placeholder);
}

function inheritTextBodyFormattingOnly(txBodyNode) {
  if (!txBodyNode) {
    return null;
  }
  const inherited = {};
  if (txBodyNode["a:bodyPr"]) {
    inherited["a:bodyPr"] = deepClone(txBodyNode["a:bodyPr"]);
  }
  if (txBodyNode["a:lstStyle"]) {
    inherited["a:lstStyle"] = deepClone(txBodyNode["a:lstStyle"]);
  }
  return Object.keys(inherited).length ? inherited : null;
}

function resolveEffectiveTextBody(shapeNode, inheritedShapeNode, placeholder) {
  const directTxBody = shapeNode?.["p:txBody"] || null;
  const inheritedTxBody = inheritedShapeNode?.["p:txBody"] || null;
  if (!directTxBody && !inheritedTxBody) {
    return null;
  }

  if (!placeholderSuppresssInheritedParagraphs(placeholder)) {
    return deepMergeNodes(inheritedTxBody, directTxBody);
  }

  const inheritedFormatting = inheritTextBodyFormattingOnly(inheritedTxBody);
  if (!directTxBody) {
    return inheritedFormatting;
  }

  const merged = deepMergeNodes(inheritedFormatting, directTxBody) || {};
  if (Object.prototype.hasOwnProperty.call(directTxBody, "a:p")) {
    merged["a:p"] = deepClone(directTxBody["a:p"]);
  } else {
    delete merged["a:p"];
  }
  return merged;
}

function parseShapeElement(shapeNode, inheritedShapeNode, themeContext, masterTextStyles = null) {
  const effectiveSpPr = deepMergeNodes(inheritedShapeNode?.["p:spPr"], shapeNode?.["p:spPr"]);
  const effectiveStyle = deepMergeNodes(inheritedShapeNode?.["p:style"], shapeNode?.["p:style"]);
  const placeholder = placeholderInfo(shapeNode) || placeholderInfo(inheritedShapeNode);
  const effectiveTxBody = resolveEffectiveTextBody(shapeNode, inheritedShapeNode, placeholder);
  const textStyleFallback = resolveTextStyleForPlaceholder(placeholder, masterTextStyles);

  const common = parseShapeCommon(shapeNode, effectiveSpPr || {}, effectiveStyle || {}, themeContext);
  const fontRefColor = resolveDrawingColor(effectiveStyle?.["a:fontRef"], themeContext)?.color || null;
  const fontRefIdx = (effectiveStyle?.["a:fontRef"]?.["@_idx"] || "").toLowerCase();
  const fontFromRef = fontRefIdx === "major"
    ? themeContext?.theme?.fontScheme?.major?.latin
    : themeContext?.theme?.fontScheme?.minor?.latin;
  const text = parseTextBody(effectiveTxBody, themeContext, {
    color: fontRefColor || "#000000",
    fontFamily: fontFromRef || themeContext?.theme?.fontScheme?.minor?.latin || "Calibri"
  }, textStyleFallback);

  const geometry = parseGeometry(effectiveSpPr || {});

  const hasText = text && text.paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text.length > 0));

  return {
    ...common,
    type: hasText ? "text" : "shape",
    shapeType: geometry.preset || "rect",
    geometry,
    placeholder,
    text,
    raw: {
      style: deepClone(effectiveStyle),
      spPr: deepClone(effectiveSpPr)
    }
  };
}

function parseConnectorElement(connectorNode, themeContext) {
  const cNvPr = connectorNode?.["p:nvCxnSpPr"]?.["p:cNvPr"] || {};
  const spPr = connectorNode?.["p:spPr"] || {};
  const geometry = parseGeometry(spPr);
  const common = parseShapeCommon(
    { "p:nvSpPr": { "p:cNvPr": cNvPr } },
    spPr,
    connectorNode?.["p:style"] || {},
    themeContext
  );
  return {
    ...common,
    type: "line",
    shapeType: geometry?.preset || "line",
    geometry,
    text: null,
    placeholder: null,
    raw: {
      style: deepClone(connectorNode?.["p:style"] || {}),
      spPr: deepClone(spPr)
    }
  };
}

function parsePictureElement(picNode, relsMap, slidePath, packageModel, contentTypes) {
  const cNvPr = picNode?.["p:nvPicPr"]?.["p:cNvPr"] || {};
  const spPr = picNode?.["p:spPr"] || {};
  const transform = parseTransform(spPr?.["a:xfrm"] || {});

  const blip = picNode?.["p:blipFill"]?.["a:blip"];
  const relId = blip?.["@_r:embed"] || null;
  const linkId = blip?.["@_r:link"] || null;

  let imagePath = null;
  let mimeType = null;
  let dataUri = null;

  if (relId && relsMap.has(relId)) {
    const rel = relsMap.get(relId);
    imagePath = resolveTargetPath(slidePath, rel.target);
    mimeType = contentTypes.get(imagePath) || detectImageMimeByExt(imagePath);

    if (packageModel.hasPart(imagePath)) {
      dataUri = packageModel.readBinary(imagePath).then((bytes) => `data:${mimeType};base64,${uint8ToBase64(bytes)}`);
    }
  }

  return {
    id: String(cNvPr?.["@_id"] || "0"),
    name: cNvPr?.["@_name"] || "",
    description: cNvPr?.["@_descr"] || "",
    hidden: normalizeBool(cNvPr?.["@_hidden"]),
    type: "image",
    x: transform.x,
    y: transform.y,
    cx: transform.cx,
    cy: transform.cy,
    rotation: transform.rotation,
    flipH: transform.flipH,
    flipV: transform.flipV,
    relId,
    linkId,
    imagePath,
    mimeType,
    dataUri,
    fill: null,
    line: null,
    text: null,
    raw: {
      spPr: deepClone(spPr)
    }
  };
}

function parseTableCell(cellNode, themeContext, fallbackFill) {
  const tcPr = cellNode?.["a:tcPr"] || {};
  const cellAttrs = cellNode || {};
  const explicitFill = resolveFill(tcPr, null, themeContext);
  const fill = explicitFill || fallbackFill || {
    type: "solid",
    color: "#FFFFFF",
    alpha: 1
  };

  const explicitBorderNodes = {
    left: Boolean(tcPr?.["a:lnL"]),
    right: Boolean(tcPr?.["a:lnR"]),
    top: Boolean(tcPr?.["a:lnT"]),
    bottom: Boolean(tcPr?.["a:lnB"])
  };
  const borders = {
    left: resolveLine(tcPr?.["a:lnL"] ? { "a:ln": tcPr["a:lnL"] } : null, null, themeContext),
    right: resolveLine(tcPr?.["a:lnR"] ? { "a:ln": tcPr["a:lnR"] } : null, null, themeContext),
    top: resolveLine(tcPr?.["a:lnT"] ? { "a:ln": tcPr["a:lnT"] } : null, null, themeContext),
    bottom: resolveLine(tcPr?.["a:lnB"] ? { "a:ln": tcPr["a:lnB"] } : null, null, themeContext)
  };

  return {
    text: parseTextBody(cellNode?.["a:txBody"], themeContext, {
      color: "#000000",
      fontFamily: themeContext?.theme?.fontScheme?.minor?.latin || "Calibri"
    }),
    fill,
    borders,
    verticalAlign: tcPr?.["@_anchor"] || "t",
    marginLeft: toInt(tcPr?.["@_marL"], 0),
    marginRight: toInt(tcPr?.["@_marR"], 0),
    marginTop: toInt(tcPr?.["@_marT"], 0),
    marginBottom: toInt(tcPr?.["@_marB"], 0),
    rowSpan: toInt(cellAttrs?.["@_rowSpan"] ?? tcPr?.["@_rowSpan"], 1),
    gridSpan: toInt(cellAttrs?.["@_gridSpan"] ?? tcPr?.["@_gridSpan"], 1),
    hMerge: normalizeBool(cellAttrs?.["@_hMerge"] ?? tcPr?.["@_hMerge"]),
    vMerge: normalizeBool(cellAttrs?.["@_vMerge"] ?? tcPr?.["@_vMerge"]),
    _styleMeta: {
      explicitFill: Boolean(explicitFill),
      explicitBorders: explicitBorderNodes
    },
    raw: deepClone(tcPr)
  };
}

function tableStyleIdFromTblPr(tblPr, tableStylesXml) {
  const raw = tblPr?.["a:tableStyleId"];
  if (typeof raw === "string" && raw) {
    return raw;
  }
  if (raw && typeof raw === "object" && typeof raw["#text"] === "string") {
    return raw["#text"];
  }
  if (typeof tblPr?.["@_tableStyleId"] === "string" && tblPr["@_tableStyleId"]) {
    return tblPr["@_tableStyleId"];
  }
  return tableStylesXml?.["a:tblStyleLst"]?.["@_def"] || null;
}

function tableStyleNodeById(tableStylesXml, styleId) {
  const styles = ensureArray(tableStylesXml?.["a:tblStyleLst"]?.["a:tblStyle"]);
  if (!styles.length) {
    return null;
  }
  if (!styleId) {
    return first(styles) || null;
  }
  return styles.find((style) => style?.["@_styleId"] === styleId) || null;
}

function resolveTableStyleFill(fillNode, themeContext) {
  if (!fillNode) {
    return null;
  }
  const fillContainer = {};
  for (const key of ["a:solidFill", "a:gradFill", "a:blipFill", "a:noFill", "a:pattFill"]) {
    if (fillNode[key]) {
      fillContainer[key] = fillNode[key];
    }
  }
  return resolveFill(fillContainer, null, themeContext);
}

function resolveTableStyleLine(edgeNode, themeContext) {
  if (!edgeNode?.["a:ln"]) {
    return null;
  }
  return resolveLine({ "a:ln": edgeNode["a:ln"] }, null, themeContext);
}

function resolveTableTextStyle(tcTxStyle, themeContext) {
  if (!tcTxStyle) {
    return null;
  }

  const colorSource = {};
  if (tcTxStyle["a:schemeClr"]) {
    colorSource["a:schemeClr"] = tcTxStyle["a:schemeClr"];
  }
  if (tcTxStyle["a:srgbClr"]) {
    colorSource["a:srgbClr"] = tcTxStyle["a:srgbClr"];
  }
  if (tcTxStyle["a:prstClr"]) {
    colorSource["a:prstClr"] = tcTxStyle["a:prstClr"];
  }
  const color = resolveDrawingColor(colorSource, themeContext);
  const fontRefIdx = tcTxStyle?.["a:fontRef"]?.["@_idx"] || null;

  let fontFamily = null;
  if (fontRefIdx === "minor") {
    fontFamily = themeContext?.theme?.fontScheme?.minor?.ea
      || themeContext?.theme?.fontScheme?.minor?.latin
      || null;
  } else if (fontRefIdx === "major") {
    fontFamily = themeContext?.theme?.fontScheme?.major?.ea
      || themeContext?.theme?.fontScheme?.major?.latin
      || null;
  }

  return {
    bold: normalizeBool(tcTxStyle?.["@_b"]),
    italic: normalizeBool(tcTxStyle?.["@_i"]),
    color: color?.color || null,
    alpha: color?.alpha ?? 1,
    fontFamily
  };
}

function resolveTableStyleRegion(styleNode, regionName, themeContext) {
  const region = styleNode?.[regionName];
  if (!region) {
    return null;
  }

  const tcStyle = region?.["a:tcStyle"] || {};
  const tcBdr = tcStyle?.["a:tcBdr"] || {};

  return {
    fill: resolveTableStyleFill(tcStyle?.["a:fill"], themeContext),
    borders: {
      left: resolveTableStyleLine(tcBdr?.["a:left"], themeContext),
      right: resolveTableStyleLine(tcBdr?.["a:right"], themeContext),
      top: resolveTableStyleLine(tcBdr?.["a:top"], themeContext),
      bottom: resolveTableStyleLine(tcBdr?.["a:bottom"], themeContext),
      insideH: resolveTableStyleLine(tcBdr?.["a:insideH"], themeContext),
      insideV: resolveTableStyleLine(tcBdr?.["a:insideV"], themeContext)
    },
    textStyle: resolveTableTextStyle(region?.["a:tcTxStyle"], themeContext)
  };
}

function cloneStyleValue(value) {
  return value ? deepClone(value) : value;
}

function applyTextStyleOverrides(textBody, textStyle) {
  if (!textBody || !textStyle) {
    return;
  }
  for (const paragraph of ensureArray(textBody.paragraphs)) {
    for (const run of ensureArray(paragraph?.runs)) {
      run.style = {
        ...run.style,
        bold: textStyle.bold ? true : run.style?.bold || false,
        italic: textStyle.italic ? true : run.style?.italic || false,
        color: run.style?.explicitColor
          ? (run.style?.color || "#000000")
          : (textStyle.color || run.style?.color || "#000000"),
        alpha: run.style?.explicitColor
          ? (run.style?.alpha ?? 1)
          : (textStyle.alpha ?? run.style?.alpha ?? 1),
        fontFamily: textStyle.fontFamily || run.style?.fontFamily || null,
        eastAsiaFont: textStyle.fontFamily || run.style?.eastAsiaFont || null
      };
    }
  }
}

function pickStyledBorder(regionStyle, side, rowIndex, rowCount, colIndex, colCount) {
  if (!regionStyle?.borders) {
    return null;
  }

  switch (side) {
    case "left":
      return colIndex === 0
        ? (regionStyle.borders.left || regionStyle.borders.insideV || null)
        : (regionStyle.borders.insideV || regionStyle.borders.left || null);
    case "right":
      return colIndex === colCount - 1
        ? (regionStyle.borders.right || regionStyle.borders.insideV || null)
        : (regionStyle.borders.insideV || regionStyle.borders.right || null);
    case "top":
      return rowIndex === 0
        ? (regionStyle.borders.top || regionStyle.borders.insideH || null)
        : (regionStyle.borders.insideH || regionStyle.borders.top || null);
    case "bottom":
      return rowIndex === rowCount - 1
        ? (regionStyle.borders.bottom || regionStyle.borders.insideH || null)
        : (regionStyle.borders.insideH || regionStyle.borders.bottom || null);
    default:
      return null;
  }
}

function applyTableRegionStyle(cell, regionStyle, rowIndex, rowCount, colIndex, colCount) {
  if (!regionStyle) {
    return;
  }

  if (!cell?._styleMeta?.explicitFill && regionStyle.fill) {
    cell.fill = cloneStyleValue(regionStyle.fill);
  }

  for (const side of ["left", "right", "top", "bottom"]) {
    if (cell?._styleMeta?.explicitBorders?.[side]) {
      continue;
    }
    const border = pickStyledBorder(regionStyle, side, rowIndex, rowCount, colIndex, colCount);
    if (border) {
      cell.borders[side] = cloneStyleValue(border);
    }
  }

  applyTextStyleOverrides(cell.text, regionStyle.textStyle);
}

function applyResolvedTableStyle(tableModel, tableStyleNode, themeContext) {
  if (!tableModel || !tableStyleNode) {
    return;
  }

  const rowCount = ensureArray(tableModel.rows).length;
  const colCount = Math.max(...ensureArray(tableModel.rows).map((row) => ensureArray(row?.cells).length), 0);
  const regions = {
    wholeTbl: resolveTableStyleRegion(tableStyleNode, "a:wholeTbl", themeContext),
    firstRow: resolveTableStyleRegion(tableStyleNode, "a:firstRow", themeContext),
    lastRow: resolveTableStyleRegion(tableStyleNode, "a:lastRow", themeContext),
    firstCol: resolveTableStyleRegion(tableStyleNode, "a:firstCol", themeContext),
    lastCol: resolveTableStyleRegion(tableStyleNode, "a:lastCol", themeContext),
    band1H: resolveTableStyleRegion(tableStyleNode, "a:band1H", themeContext),
    band2H: resolveTableStyleRegion(tableStyleNode, "a:band2H", themeContext),
    band1V: resolveTableStyleRegion(tableStyleNode, "a:band1V", themeContext),
    band2V: resolveTableStyleRegion(tableStyleNode, "a:band2V", themeContext)
  };

  for (let ri = 0; ri < rowCount; ri += 1) {
    const row = tableModel.rows[ri];
    const bodyRowIndex = ri - (tableModel.firstRow ? 1 : 0);
    for (let ci = 0; ci < ensureArray(row?.cells).length; ci += 1) {
      const cell = row.cells[ci];
      applyTableRegionStyle(cell, regions.wholeTbl, ri, rowCount, ci, colCount);

      if (tableModel.bandRow && bodyRowIndex >= 0) {
        const bandRegion = bodyRowIndex % 2 === 0 ? regions.band1H : regions.band2H;
        applyTableRegionStyle(cell, bandRegion, ri, rowCount, ci, colCount);
      }
      if (tableModel.bandCol) {
        const bodyColIndex = ci - (tableModel.firstCol ? 1 : 0);
        if (bodyColIndex >= 0) {
          const bandRegion = bodyColIndex % 2 === 0 ? regions.band1V : regions.band2V;
          applyTableRegionStyle(cell, bandRegion, ri, rowCount, ci, colCount);
        }
      }
      if (tableModel.firstRow && ri === 0) {
        applyTableRegionStyle(cell, regions.firstRow, ri, rowCount, ci, colCount);
      }
      if (tableModel.lastRow && ri === rowCount - 1) {
        applyTableRegionStyle(cell, regions.lastRow, ri, rowCount, ci, colCount);
      }
      if (tableModel.firstCol && ci === 0) {
        applyTableRegionStyle(cell, regions.firstCol, ri, rowCount, ci, colCount);
      }
      if (tableModel.lastCol && ci === colCount - 1) {
        applyTableRegionStyle(cell, regions.lastCol, ri, rowCount, ci, colCount);
      }
    }
  }
}

function chartTextFromRich(richNode) {
  const parts = [];
  for (const paragraph of ensureArray(richNode?.["a:p"])) {
    for (const run of ensureArray(paragraph?.["a:r"])) {
      parts.push(run?.["a:t"] || "");
    }
    if (!ensureArray(paragraph?.["a:r"]).length && paragraph?.["a:t"]) {
      parts.push(paragraph["a:t"]);
    }
  }
  return parts.join("");
}

function chartTitleText(chartNode) {
  const titleNode = chartNode?.["c:title"];
  if (!titleNode) {
    return null;
  }
  const tx = titleNode?.["c:tx"] || {};
  const rich = tx?.["c:rich"];
  if (rich) {
    const text = chartTextFromRich(rich);
    if (text) {
      return text;
    }
  }
  const strRef = tx?.["c:strRef"]?.["c:strCache"];
  const pt = first(strRef?.["c:pt"]);
  return pt?.["c:v"] || null;
}

function readChartPointValues(cacheNode) {
  const points = ensureArray(cacheNode?.["c:pt"]).map((pt) => ({
    idx: toInt(pt?.["@_idx"], 0),
    value: pt?.["c:v"] || ""
  }));
  points.sort((a, b) => a.idx - b.idx);
  return points.map((p) => p.value);
}

function readChartCategoryValues(catNode) {
  const strCache = catNode?.["c:strRef"]?.["c:strCache"] || catNode?.["c:strLit"];
  if (strCache) {
    return readChartPointValues(strCache).map((v) => String(v));
  }
  const numCache = catNode?.["c:numRef"]?.["c:numCache"] || catNode?.["c:numLit"];
  if (numCache) {
    return readChartPointValues(numCache).map((v) => String(v));
  }
  return [];
}

function readChartNumericValues(valNode) {
  const numCache = valNode?.["c:numRef"]?.["c:numCache"] || valNode?.["c:numLit"];
  if (!numCache) {
    return [];
  }
  return readChartPointValues(numCache).map((v) => Number.parseFloat(v));
}

function parseChartSeriesColorFromSpPr(spPrNode, themeContext) {
  const fill = resolveFill(spPrNode || {}, null, themeContext);
  return fill?.color || null;
}

function parsePieChartModel(chartNode, themeContext) {
  const pieChart = first(chartNode?.["c:plotArea"]?.["c:pieChart"]);
  if (!pieChart) {
    return null;
  }

  const series = ensureArray(pieChart?.["c:ser"]).map((serNode) => {
    const categories = readChartCategoryValues(serNode?.["c:cat"]);
    const values = readChartNumericValues(serNode?.["c:val"]);
    const pointColorMap = new Map();
    for (const dPtNode of ensureArray(serNode?.["c:dPt"])) {
      const idx = toInt(dPtNode?.["c:idx"]?.["@_val"], -1);
      if (idx < 0) {
        continue;
      }
      const color = parseChartSeriesColorFromSpPr(dPtNode?.["c:spPr"], themeContext);
      if (color) {
        pointColorMap.set(idx, color);
      }
    }
    const fallbackColor = parseChartSeriesColorFromSpPr(serNode?.["c:spPr"], themeContext) || "#92278F";
    const colors = values.map((_, idx) => pointColorMap.get(idx) || fallbackColor);

    return {
      categories,
      values: values.map((v) => (Number.isFinite(v) ? v : 0)),
      colors
    };
  });

  return {
    chartType: "pie",
    title: chartTitleText(chartNode),
    series
  };
}

function parseChartElement(graphicFrameNode, themeContext, relsMap, partPath, packageModel) {
  const chartNode = first(graphicFrameNode?.["a:graphic"]?.["a:graphicData"]?.["c:chart"]);
  if (!chartNode) {
    return null;
  }

  const cNvPr = graphicFrameNode?.["p:nvGraphicFramePr"]?.["p:cNvPr"] || {};
  const transform = parseTransform(graphicFrameNode?.["p:xfrm"] || {});
  const chartRelId = chartNode?.["@_r:id"] || null;
  const relationship = chartRelId && relsMap?.has(chartRelId) ? relsMap.get(chartRelId) : null;
  const chartPath = relationship ? resolveTargetPath(partPath, relationship.target) : null;

  let chart = null;
  if (chartPath && packageModel.hasPart(chartPath)) {
    chart = packageModel.readXml(chartPath).then((chartXml) => {
      const chartRoot = chartXml?.["c:chartSpace"]?.["c:chart"];
      if (!chartRoot) {
        return null;
      }
      return parsePieChartModel(chartRoot, themeContext);
    }).catch(() => null);
  }

  return {
    id: String(cNvPr?.["@_id"] || "0"),
    name: cNvPr?.["@_name"] || "",
    description: cNvPr?.["@_descr"] || "",
    hidden: normalizeBool(cNvPr?.["@_hidden"]),
    type: "chart",
    x: transform.x,
    y: transform.y,
    cx: transform.cx,
    cy: transform.cy,
    rotation: transform.rotation,
    flipH: transform.flipH,
    flipV: transform.flipV,
    chartRelId,
    chartPath,
    chart,
    fill: null,
    line: null,
    text: null,
    raw: deepClone(graphicFrameNode)
  };
}

function parseDiagramShapeNode(shapeNode, themeContext) {
  const cNvPr = shapeNode?.["dsp:nvSpPr"]?.["dsp:cNvPr"] || {};
  const spPr = shapeNode?.["dsp:spPr"] || {};
  const style = shapeNode?.["dsp:style"] || {};
  const transform = parseTransform(spPr?.["a:xfrm"] || {});
  const fill = resolveFill(spPr, style, themeContext);
  const line = resolveLine(spPr, style, themeContext);
  const text = parseTextBody(shapeNode?.["dsp:txBody"], themeContext, {
    color: "#000000",
    fontFamily: themeContext?.theme?.fontScheme?.minor?.latin || "Calibri"
  });
  const geometry = parseGeometry(spPr);
  const shapeType = geometry?.preset || "rect";
  const hasText = text && text.paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text.length > 0));
  const isLine = shapeType === "line" && !hasText;

  return {
    id: String(cNvPr?.["@_id"] || "0"),
    name: cNvPr?.["@_name"] || "",
    description: cNvPr?.["@_descr"] || "",
    hidden: normalizeBool(cNvPr?.["@_hidden"]),
    type: isLine ? "line" : (hasText ? "text" : "shape"),
    x: transform.x,
    y: transform.y,
    cx: transform.cx,
    cy: transform.cy,
    rotation: transform.rotation,
    flipH: transform.flipH,
    flipV: transform.flipV,
    fill,
    line,
    text,
    shapeType,
    geometry,
    raw: deepClone(shapeNode)
  };
}

function parseDiagramElement(graphicFrameNode, themeContext, relsMap, partPath, packageModel) {
  const graphicData = graphicFrameNode?.["a:graphic"]?.["a:graphicData"];
  if (!graphicData || graphicData?.["@_uri"] !== "http://schemas.openxmlformats.org/drawingml/2006/diagram") {
    return null;
  }

  let drawingRel = null;
  for (const rel of relsMap.values()) {
    const type = String(rel?.type || "");
    if (type.endsWith("/diagramDrawing") || type.includes("/diagramDrawing")) {
      drawingRel = rel;
      break;
    }
  }
  if (!drawingRel) {
    return null;
  }

  const cNvPr = graphicFrameNode?.["p:nvGraphicFramePr"]?.["p:cNvPr"] || {};
  const transform = parseTransform(graphicFrameNode?.["p:xfrm"] || {});
  const drawingPath = resolveTargetPath(partPath, drawingRel.target);
  let diagramElements = [];

  if (drawingPath && packageModel.hasPart(drawingPath)) {
    diagramElements = packageModel.readXml(drawingPath).then((drawingXml) => {
      const spTree = drawingXml?.["dsp:drawing"]?.["dsp:spTree"] || {};
      const localElements = ensureArray(spTree?.["dsp:sp"]).map((shapeNode) => parseDiagramShapeNode(shapeNode, themeContext));
      if (!localElements.length) {
        return [];
      }

      let maxRight = 1;
      let maxBottom = 1;
      for (const local of localElements) {
        maxRight = Math.max(maxRight, (local.x || 0) + Math.max(0, local.cx || 0));
        maxBottom = Math.max(maxBottom, (local.y || 0) + Math.max(0, local.cy || 0));
      }
      const scaleX = transform.cx ? transform.cx / maxRight : 1;
      const scaleY = transform.cy ? transform.cy / maxBottom : 1;

      return localElements.map((local) => ({
        ...local,
        x: transform.x + (local.x || 0) * scaleX,
        y: transform.y + (local.y || 0) * scaleY,
        cx: (local.cx || 0) * scaleX,
        cy: (local.cy || 0) * scaleY,
        rotation: (local.rotation || 0) + (transform.rotation || 0)
      }));
    }).catch(() => []);
  }

  return {
    id: String(cNvPr?.["@_id"] || "0"),
    name: cNvPr?.["@_name"] || "",
    description: cNvPr?.["@_descr"] || "",
    hidden: normalizeBool(cNvPr?.["@_hidden"]),
    type: "diagram",
    x: transform.x,
    y: transform.y,
    cx: transform.cx,
    cy: transform.cy,
    rotation: transform.rotation,
    flipH: transform.flipH,
    flipV: transform.flipV,
    drawingPath,
    diagramElements,
    fill: null,
    line: null,
    text: null,
    raw: deepClone(graphicFrameNode)
  };
}

function parseTableElement(graphicFrameNode, themeContext, relsMap, partPath, packageModel, tableStylesXml = null) {
  const cNvPr = graphicFrameNode?.["p:nvGraphicFramePr"]?.["p:cNvPr"] || {};
  const xfrm = graphicFrameNode?.["p:xfrm"] || {};
  const transform = parseTransform(xfrm);

  const table = graphicFrameNode?.["a:graphic"]?.["a:graphicData"]?.["a:tbl"];
  if (!table) {
    return parseChartElement(graphicFrameNode, themeContext, relsMap, partPath, packageModel)
      || parseDiagramElement(graphicFrameNode, themeContext, relsMap, partPath, packageModel);
  }

  const gridCols = ensureArray(table?.["a:tblGrid"]?.["a:gridCol"]).map((col) => toInt(col?.["@_w"], 0));
  const rows = ensureArray(table?.["a:tr"]).map((row) => ({
    height: toInt(row?.["@_h"], 0),
    cells: ensureArray(row?.["a:tc"]).map((cell) => parseTableCell(cell, themeContext, null))
  }));

  const tblPr = table?.["a:tblPr"] || {};
  const styleId = tableStyleIdFromTblPr(tblPr, tableStylesXml);

  const tableModel = {
    id: String(cNvPr?.["@_id"] || "0"),
    name: cNvPr?.["@_name"] || "",
    description: cNvPr?.["@_descr"] || "",
    hidden: normalizeBool(cNvPr?.["@_hidden"]),
    type: "table",
    x: transform.x,
    y: transform.y,
    cx: transform.cx,
    cy: transform.cy,
    rotation: transform.rotation,
    flipH: transform.flipH,
    flipV: transform.flipV,
    styleId,
    firstRow: normalizeBool(tblPr?.["@_firstRow"]),
    firstCol: normalizeBool(tblPr?.["@_firstCol"]),
    lastRow: normalizeBool(tblPr?.["@_lastRow"]),
    lastCol: normalizeBool(tblPr?.["@_lastCol"]),
    bandRow: normalizeBool(tblPr?.["@_bandRow"]),
    bandCol: normalizeBool(tblPr?.["@_bandCol"]),
    gridCols,
    rows,
    fill: null,
    line: null,
    text: null,
    raw: deepClone(graphicFrameNode)
  };

  const tableStyleNode = tableStyleNodeById(tableStylesXml, styleId);
  applyResolvedTableStyle(tableModel, tableStyleNode, themeContext);

  return tableModel;
}
function resolveGradientFirstColor(gradFillNode, themeContext) {
  const stops = ensureArray(gradFillNode?.["a:gsLst"]?.["a:gs"]);
  if (!stops.length) {
    return null;
  }
  const resolved = resolveDrawingColor(first(stops), themeContext);
  if (!resolved?.color) {
    return null;
  }
  return {
    type: "solid",
    color: resolved.color,
    alpha: resolved.alpha,
    source: "gradient-fallback"
  };
}

function parseGradientFillModel(gradFillNode, themeContext, placeholderColor = null, source = "gradient") {
  const stops = ensureArray(gradFillNode?.["a:gsLst"]?.["a:gs"])
    .map((stop) => {
      const resolved = resolveDrawingColor(stop, themeContext, placeholderColor);
      if (!resolved?.color) {
        return null;
      }
      return {
        pos: clamp(toInt(stop?.["@_pos"], 0), 0, 100000),
        color: resolved.color,
        alpha: resolved.alpha
      };
    })
    .filter(Boolean);

  if (!stops.length) {
    return null;
  }

  const linearNode = first(gradFillNode?.["a:lin"]) || gradFillNode?.["a:lin"] || null;
  const pathNode = first(gradFillNode?.["a:path"]) || gradFillNode?.["a:path"] || null;
  const fillToRectNode = first(pathNode?.["a:fillToRect"]) || pathNode?.["a:fillToRect"] || null;

  return {
    type: "gradient",
    gradientType: pathNode ? "path" : "linear",
    angle: deg60000ToDeg(toInt(linearNode?.["@_ang"], 0)),
    scaled: normalizeBool(linearNode?.["@_scaled"]),
    path: pathNode?.["@_path"] || null,
    fillToRect: fillToRectNode ? {
      l: toInt(fillToRectNode?.["@_l"], 0),
      t: toInt(fillToRectNode?.["@_t"], 0),
      r: toInt(fillToRectNode?.["@_r"], 0),
      b: toInt(fillToRectNode?.["@_b"], 0)
    } : null,
    stops,
    source
  };
}

async function resolveBackgroundImageFromBlip(
  blipFillNode,
  relsMap,
  partPath,
  packageModel,
  contentTypes,
  imageDataCache,
  source
) {
  const blipNode = first(blipFillNode?.["a:blip"]) || blipFillNode?.["a:blip"];
  const relId = blipNode?.["@_r:embed"] || null;
  if (!relId || !relsMap?.has(relId) || !partPath) {
    return null;
  }

  const rel = relsMap.get(relId);
  if (!rel || rel.targetMode === "External") {
    return null;
  }

  const imagePath = resolveTargetPath(partPath, rel.target);
  if (!imagePath || !packageModel.hasPart(imagePath)) {
    return null;
  }

  const mimeType = contentTypes.get(imagePath) || detectImageMimeByExt(imagePath);
  let dataUri = imageDataCache.get(imagePath);
  if (!dataUri) {
    try {
      const bytes = await packageModel.readBinary(imagePath);
      dataUri = `data:${mimeType};base64,${uint8ToBase64(bytes)}`;
      imageDataCache.set(imagePath, dataUri);
    } catch {
      return null;
    }
  }

  return {
    type: "image",
    imagePath,
    mimeType,
    dataUri,
    relId,
    source
  };
}

async function resolveBackgroundFromBgPr(
  bgPrNode,
  themeContext,
  relsMap,
  partPath,
  packageModel,
  contentTypes,
  imageDataCache,
  source
) {
  if (!bgPrNode) {
    return null;
  }

  if (hasNode(bgPrNode, "a:noFill")) {
    return {
      type: "none",
      color: null,
      alpha: 0,
      source
    };
  }

  const blipFillNode = first(bgPrNode?.["a:blipFill"]) || bgPrNode?.["a:blipFill"];
  if (blipFillNode) {
    const image = await resolveBackgroundImageFromBlip(
      blipFillNode,
      relsMap,
      partPath,
      packageModel,
      contentTypes,
      imageDataCache,
      source
    );
    if (image) {
      return image;
    }
  }

  if (hasNode(bgPrNode, "a:solidFill")) {
    const color = resolveDrawingColor(first(bgPrNode["a:solidFill"]), themeContext);
    if (color?.color) {
      return {
        type: "solid",
        color: color.color,
        alpha: color.alpha,
        source
      };
    }
  }

  if (hasNode(bgPrNode, "a:gradFill")) {
    const gradient = parseGradientFillModel(first(bgPrNode["a:gradFill"]), themeContext, null, source)
      || resolveGradientFirstColor(first(bgPrNode["a:gradFill"]), themeContext);
    if (gradient) {
      return {
        ...gradient,
        source
      };
    }
  }

  return null;
}

function resolveBackgroundFromThemeStyle(bgStyleEntry, themeContext, placeholderColor = null) {
  if (!bgStyleEntry) {
    return null;
  }

  if (bgStyleEntry.type === "a:solidFill") {
    const resolved = resolveDrawingColor(bgStyleEntry.node, themeContext, placeholderColor);
    if (resolved?.color) {
      return {
        type: "solid",
        color: resolved.color,
        alpha: resolved.alpha,
        source: "theme-bg-style"
      };
    }
  }

  if (bgStyleEntry.type === "a:gradFill") {
    const gradient = parseGradientFillModel(bgStyleEntry.node, themeContext, placeholderColor, "theme-bg-style")
      || resolveGradientFirstColor(bgStyleEntry.node, themeContext);
    if (gradient) {
      return {
        ...gradient,
        source: "theme-bg-style"
      };
    }
  }

  return null;
}

async function resolveBackgroundFromBgRef(
  bgRefNode,
  themeContext,
  themeRels,
  themePath,
  packageModel,
  contentTypes,
  imageDataCache,
  source
) {
  if (!bgRefNode) {
    return null;
  }

  const idx = toInt(bgRefNode?.["@_idx"], 0);
  const referenceColor = resolveDrawingColor(bgRefNode, themeContext);
  const bgStyle = themeContext?.getBgFillStyle?.(idx);
  if (bgStyle?.type === "a:blipFill") {
    const image = await resolveBackgroundImageFromBlip(
      bgStyle.node,
      themeRels,
      themePath,
      packageModel,
      contentTypes,
      imageDataCache,
      source
    );
    if (image) {
      return image;
    }
  }

  const styleColor = resolveBackgroundFromThemeStyle(bgStyle, themeContext, referenceColor);
  if (styleColor) {
    return {
      ...styleColor,
      source
    };
  }

  const color = referenceColor;
  if (color?.color) {
    return {
      type: "solid",
      color: color.color,
      alpha: color.alpha,
      source
    };
  }

  return null;
}

async function resolveBackgroundColor(slideXml, layoutXml, masterXml, themeContext, options = {}) {
  const sourceNodes = [
    {
      name: "slide",
      node: slideXml?.["p:sld"]?.["p:cSld"]?.["p:bg"],
      rels: options.slideRels,
      partPath: options.slidePath
    },
    {
      name: "layout",
      node: layoutXml?.["p:sldLayout"]?.["p:cSld"]?.["p:bg"],
      rels: options.layoutRels,
      partPath: options.layoutPath
    },
    {
      name: "master",
      node: masterXml?.["p:sldMaster"]?.["p:cSld"]?.["p:bg"],
      rels: options.masterRels,
      partPath: options.masterPath
    }
  ];

  for (const sourceNode of sourceNodes) {
    const bgNode = sourceNode.node;
    if (!bgNode) {
      continue;
    }

    const bgPr = await resolveBackgroundFromBgPr(
      bgNode?.["p:bgPr"],
      themeContext,
      sourceNode.rels,
      sourceNode.partPath,
      options.packageModel,
      options.contentTypes,
      options.imageDataCache || new Map(),
      `bgPr-${sourceNode.name}`
    );
    if (bgPr) {
      return bgPr;
    }

    const bgRef = await resolveBackgroundFromBgRef(
      bgNode?.["p:bgRef"],
      themeContext,
      options.themeRels,
      options.themePath,
      options.packageModel,
      options.contentTypes,
      options.imageDataCache || new Map(),
      `bgRef-${sourceNode.name}`
    );
    if (bgRef) {
      return bgRef;
    }
  }

  return {
    type: "solid",
    color: "#FFFFFF",
    alpha: 1,
    source: "default"
  };
}

function parseGroupXfrm(grpSpNode) {
  const xfrm = grpSpNode?.["p:grpSpPr"]?.["a:xfrm"] || {};
  const offX = toInt(xfrm?.["a:off"]?.["@_x"], 0);
  const offY = toInt(xfrm?.["a:off"]?.["@_y"], 0);
  const extX = toInt(xfrm?.["a:ext"]?.["@_cx"], 0);
  const extY = toInt(xfrm?.["a:ext"]?.["@_cy"], 0);
  const chOffX = toInt(xfrm?.["a:chOff"]?.["@_x"], 0);
  const chOffY = toInt(xfrm?.["a:chOff"]?.["@_y"], 0);
  const chExtX = toInt(xfrm?.["a:chExt"]?.["@_cx"], extX || 1);
  const chExtY = toInt(xfrm?.["a:chExt"]?.["@_cy"], extY || 1);

  const scaleX = chExtX ? extX / chExtX : 1;
  const scaleY = chExtY ? extY / chExtY : 1;

  return {
    offX: offX - chOffX * scaleX,
    offY: offY - chOffY * scaleY,
    scaleX,
    scaleY,
    rotation: deg60000ToDeg(toInt(xfrm?.["@_rot"], 0))
  };
}

function composeGroupTransform(parentTransform, groupTransform) {
  return {
    offX: parentTransform.offX + parentTransform.scaleX * groupTransform.offX,
    offY: parentTransform.offY + parentTransform.scaleY * groupTransform.offY,
    scaleX: parentTransform.scaleX * groupTransform.scaleX,
    scaleY: parentTransform.scaleY * groupTransform.scaleY,
    rotation: (parentTransform.rotation || 0) + (groupTransform.rotation || 0)
  };
}

function applyGroupTransformToElement(element, groupTransform) {
  const transformed = element;
  transformed.x = groupTransform.offX + transformed.x * groupTransform.scaleX;
  transformed.y = groupTransform.offY + transformed.y * groupTransform.scaleY;
  transformed.cx *= groupTransform.scaleX;
  transformed.cy *= groupTransform.scaleY;
  transformed.rotation = (transformed.rotation || 0) + (groupTransform.rotation || 0);
  return transformed;
}

function parseOrderedTreeChildren(
  treeNode,
  themeContext,
  relsMap,
  partPath,
  packageModel,
  contentTypes,
  tableStylesXml,
  sourceLayer,
  pushElement
) {
  for (const key of Object.keys(treeNode || {})) {
    switch (key) {
      case "p:sp": {
        for (const shapeNode of ensureArray(treeNode[key])) {
          if (placeholderInfo(shapeNode)) {
            continue;
          }
          const shape = parseShapeElement(shapeNode, null, themeContext);
          shape.sourceLayer = sourceLayer;
          pushElement(shape);
        }
        break;
      }
      case "p:cxnSp": {
        for (const connectorNode of ensureArray(treeNode[key])) {
          const connector = parseConnectorElement(connectorNode, themeContext);
          connector.sourceLayer = sourceLayer;
          pushElement(connector);
        }
        break;
      }
      case "p:pic": {
        for (const picNode of ensureArray(treeNode[key])) {
          const picture = parsePictureElement(picNode, relsMap, partPath, packageModel, contentTypes);
          picture.sourceLayer = sourceLayer;
          pushElement(picture);
        }
        break;
      }
      case "p:graphicFrame": {
        for (const graphicFrameNode of ensureArray(treeNode[key])) {
          const table = parseTableElement(
            graphicFrameNode,
            themeContext,
            relsMap,
            partPath,
            packageModel,
            tableStylesXml
          );
          if (table) {
            table.sourceLayer = sourceLayer;
            pushElement(table);
          }
        }
        break;
      }
      default:
        break;
    }
  }
}

function orderedEntryName(entry) {
  if (!entry || typeof entry !== "object" || Array.isArray(entry)) {
    return null;
  }
  return Object.keys(entry || {}).find((key) => key !== ":@") || null;
}

function orderedEntryChildren(entry) {
  const name = orderedEntryName(entry);
  return name && Array.isArray(entry?.[name]) ? entry[name] : [];
}

function findOrderedChildrenByPath(entries, path) {
  let current = ensureArray(entries);
  for (const segment of ensureArray(path)) {
    const next = current.find((entry) => orderedEntryName(entry) === segment);
    if (!next) {
      return [];
    }
    current = orderedEntryChildren(next);
  }
  return current;
}

function findFirstOrderedEntryByName(entries, targetName) {
  for (const entry of ensureArray(entries)) {
    const name = orderedEntryName(entry);
    if (name === targetName) {
      return entry;
    }
    const nested = findFirstOrderedEntryByName(orderedEntryChildren(entry), targetName);
    if (nested) {
      return nested;
    }
  }
  return null;
}

function orderedElementId(entry) {
  const cNvPr = findFirstOrderedEntryByName([entry], "p:cNvPr");
  const id = cNvPr?.[":@"]?.["@_id"];
  return id === undefined || id === null ? null : String(id);
}

const DRAW_ORDER_TAGS = new Set(["p:sp", "p:cxnSp", "p:pic", "p:graphicFrame"]);
const DRAW_ORDER_CONTAINER_TAGS = new Set(["p:grpSp", "mc:AlternateContent", "mc:Choice", "mc:Fallback"]);

function collectOrderedElementKeys(entries, prefix = [], orderMap = new Map()) {
  let position = 0;
  for (const entry of ensureArray(entries)) {
    const name = orderedEntryName(entry);
    if (!name) {
      continue;
    }
    const segment = String(position).padStart(4, "0");
    const nextPrefix = [...prefix, segment];
    if (DRAW_ORDER_TAGS.has(name)) {
      const id = orderedElementId(entry);
      if (id && !orderMap.has(id)) {
        orderMap.set(id, nextPrefix.join("."));
      }
    }
    if (DRAW_ORDER_CONTAINER_TAGS.has(name)) {
      collectOrderedElementKeys(orderedEntryChildren(entry), nextPrefix, orderMap);
    }
    position += 1;
  }
  return orderMap;
}

function buildElementOrderMap(xmlText, treePath) {
  if (!xmlText) {
    return new Map();
  }
  const orderedRoot = parseXmlPreserveOrder(xmlText);
  const spTreeChildren = findOrderedChildrenByPath(orderedRoot, treePath);
  return collectOrderedElementKeys(spTreeChildren);
}

function sortElementsByOrder(elements, orderMap) {
  return ensureArray(elements)
    .map((element, index) => ({
      element,
      index,
      key: orderMap.get(String(element?.id || "")) || `~${String(index).padStart(6, "0")}`
    }))
    .sort((a, b) => a.key.localeCompare(b.key) || a.index - b.index)
    .map((entry) => {
      entry.element.drawOrderKey = entry.key;
      return entry.element;
    });
}

function parseGroupTreeElements(
  grpSpNode,
  themeContext,
  relsMap,
  partPath,
  packageModel,
  contentTypes,
  tableStylesXml,
  sourceLayer,
  parentTransform
) {
  const elements = [];
  const local = parseGroupXfrm(grpSpNode);
  const aggregate = composeGroupTransform(parentTransform, local);

  parseOrderedTreeChildren(
    grpSpNode,
    themeContext,
    relsMap,
    partPath,
    packageModel,
    contentTypes,
    tableStylesXml,
    sourceLayer,
    (element) => {
      elements.push(applyGroupTransformToElement(element, aggregate));
    }
  );

  for (const nestedGroup of ensureArray(grpSpNode?.["p:grpSp"])) {
    elements.push(
      ...parseGroupTreeElements(
        nestedGroup,
        themeContext,
        relsMap,
        partPath,
        packageModel,
        contentTypes,
        tableStylesXml,
        sourceLayer,
        aggregate
      )
    );
  }

  return elements;
}

function flattenRelMap(relsMap) {
  return Array.from(relsMap.values()).map((rel) => ({
    id: rel.id,
    type: rel.type,
    target: rel.target,
    targetMode: rel.targetMode || "Internal"
  }));
}

function snapshotElement(element) {
  if (!element || typeof element !== "object") {
    return element;
  }
  const base = {
    id: element.id,
    type: element.type,
    name: element.name,
    description: element.description,
    hidden: element.hidden,
    x: element.x,
    y: element.y,
    cx: element.cx,
    cy: element.cy,
    rotation: element.rotation,
    flipH: element.flipH,
    flipV: element.flipV,
    fill: element.fill,
    line: element.line,
    text: element.text,
    shapeType: element.shapeType,
    geometry: element.geometry,
    relId: element.relId,
    imagePath: element.imagePath,
    mimeType: element.mimeType,
    chartRelId: element.chartRelId,
    chartPath: element.chartPath,
    chart: element.chart,
    drawingPath: element.drawingPath,
    diagramElements: element.diagramElements,
    styleId: element.styleId,
    firstRow: element.firstRow,
    firstCol: element.firstCol,
    lastRow: element.lastRow,
    lastCol: element.lastCol,
    bandRow: element.bandRow,
    bandCol: element.bandCol,
    gridCols: element.gridCols,
    rows: element.rows
  };

  if (element.type === "image") {
    base.dataUri = element.dataUri || null;
  }

  return base;
}

export function createSlideSnapshot(slide) {
  return JSON.stringify({
    background: slide?.background || null,
    elements: ensureArray(slide?.elements).map((element) => snapshotElement(element))
  });
}

async function resolveImageDataUris(elements) {
  const pending = [];
  for (const element of elements) {
    if (element.type !== "image" || typeof element.dataUri !== "object" || typeof element.dataUri.then !== "function") {
      continue;
    }
    pending.push(
      element.dataUri.then((dataUri) => {
        element.dataUri = dataUri;
      }).catch(() => {
        element.dataUri = null;
      })
    );
  }
  await Promise.all(pending);
}

async function resolveChartData(elements) {
  const pending = [];
  for (const element of elements) {
    if (element.type !== "chart" || typeof element.chart !== "object" || typeof element.chart.then !== "function") {
      continue;
    }
    pending.push(
      element.chart.then((chartModel) => {
        element.chart = chartModel;
      }).catch(() => {
        element.chart = null;
      })
    );
  }
  await Promise.all(pending);
}

async function resolveDiagramData(elements) {
  const pending = [];
  for (const element of elements) {
    if (element.type !== "diagram" || typeof element.diagramElements !== "object" || typeof element.diagramElements.then !== "function") {
      continue;
    }
    pending.push(
      element.diagramElements.then((items) => {
        element.diagramElements = items;
      }).catch(() => {
        element.diagramElements = [];
      })
    );
  }
  await Promise.all(pending);
}

function parseDecorativeTreeElements(
  spTree,
  themeContext,
  relsMap,
  partPath,
  packageModel,
  contentTypes,
  tableStylesXml,
  sourceLayer
) {
  const elements = [];
  const identity = { offX: 0, offY: 0, scaleX: 1, scaleY: 1, rotation: 0 };

  parseOrderedTreeChildren(
    spTree,
    themeContext,
    relsMap,
    partPath,
    packageModel,
    contentTypes,
    tableStylesXml,
    sourceLayer,
    (element) => {
      elements.push(element);
    }
  );

  for (const grpSpNode of ensureArray(spTree?.["p:grpSp"])) {
    elements.push(
      ...parseGroupTreeElements(
        grpSpNode,
        themeContext,
        relsMap,
        partPath,
        packageModel,
        contentTypes,
        tableStylesXml,
        sourceLayer,
        identity
      )
    );
  }

  return elements;
}

function parseMasterTextStyles(masterXml) {
  const txStyles = masterXml?.["p:sldMaster"]?.["p:txStyles"] || {};
  return {
    title: txStyles?.["p:titleStyle"] || null,
    body: txStyles?.["p:bodyStyle"] || null,
    other: txStyles?.["p:otherStyle"] || null
  };
}

export async function parsePresentationModel(openXmlPackage) {
  const presentationGraph = await buildPresentationGraph(openXmlPackage);
  const contentTypesXml = await openXmlPackage.readXml("[Content_Types].xml");
  const contentTypes = contentTypeMap(contentTypesXml);
  const tableStylesXml = openXmlPackage.hasPart("ppt/tableStyles.xml")
    ? await openXmlPackage.readXml("ppt/tableStyles.xml").catch(() => null)
    : null;

  const presentationRoot = presentationGraph.presentationRoot || {};
  const slideSizeNode = presentationRoot?.["p:sldSz"] || {};
  const notesSizeNode = presentationRoot?.["p:notesSz"] || {};
  const parsedThemeCache = new Map();
  const backgroundImageCache = new Map();
  const elementOrderCache = new Map();

  async function getElementOrderMap(partPath, treePath) {
    if (!partPath || !openXmlPackage.hasPart(partPath)) {
      return new Map();
    }
    if (!elementOrderCache.has(partPath)) {
      const xmlText = await openXmlPackage.readText(partPath).catch(() => null);
      elementOrderCache.set(partPath, buildElementOrderMap(xmlText, treePath));
    }
    return elementOrderCache.get(partPath) || new Map();
  }

  const slides = [];

  for (const slideRef of presentationGraph.slides) {
    const i = slideRef.index;
    const sldId = slideRef.slideIdNode;
    const relId = slideRef.relId;
    const slidePath = slideRef.slidePath;
    const slideXml = slideRef.slideXml || {};
    const slideRels = slideRef.slideRels || new Map();
    const layoutPath = slideRef.layoutPath || null;
    const layoutXml = slideRef.layoutXml || null;
    const layoutRels = slideRef.layoutRels || new Map();
    const masterPath = slideRef.masterPath || null;
    const masterXml = slideRef.masterXml || null;
    const masterRels = slideRef.masterRels || new Map();
    const themePath = slideRef.themePath || null;
    const themeXml = slideRef.themeXml || null;
    const themeRels = slideRef.themeRels || new Map();

    let parsedTheme = {
      name: null,
      colorScheme: {},
      fontScheme: {
        major: { latin: "Calibri", ea: null, cs: null },
        minor: { latin: "Calibri", ea: null, cs: null }
      },
      formatScheme: {
        fillStyles: [],
        lineStyles: [],
        bgFillStyles: []
      }
    };

    if (themePath && themeXml) {
      if (!parsedThemeCache.has(themePath)) {
        parsedThemeCache.set(themePath, parseTheme(themeXml));
      }
      parsedTheme = parsedThemeCache.get(themePath);
    }

    const baseColorMap = mergeColorMap(parseColorMapFromMaster(masterXml), null);
    const colorMap = mergeColorMap(baseColorMap, parseColorMapOverride(slideXml));

    const themeContext = createThemeContext(parsedTheme, colorMap);
    const masterTextStyles = parseMasterTextStyles(masterXml);

    const layoutPlaceholderMap = buildPlaceholderMap(layoutXml?.["p:sldLayout"]?.["p:cSld"]?.["p:spTree"]);
    const masterPlaceholderMap = buildPlaceholderMap(masterXml?.["p:sldMaster"]?.["p:cSld"]?.["p:spTree"]);
    const slideOrderMap = await getElementOrderMap(slidePath, ["p:sld", "p:cSld", "p:spTree"]);
    const layoutOrderMap = await getElementOrderMap(layoutPath, ["p:sldLayout", "p:cSld", "p:spTree"]);
    const masterOrderMap = await getElementOrderMap(masterPath, ["p:sldMaster", "p:cSld", "p:spTree"]);

    const slideTree = slideXml?.["p:sld"]?.["p:cSld"]?.["p:spTree"] || {};
    const elements = [];
    const slideGroupElements = [];
    const unhandledNodes = [];
    const renderElements = [];
    const identityTransform = { offX: 0, offY: 0, scaleX: 1, scaleY: 1, rotation: 0 };

    const masterTree = masterXml?.["p:sldMaster"]?.["p:cSld"]?.["p:spTree"] || {};
    const layoutTree = layoutXml?.["p:sldLayout"]?.["p:cSld"]?.["p:spTree"] || {};
    const layoutShowMasterSpRaw = layoutXml?.["p:sldLayout"]?.["@_showMasterSp"];
    const layoutShowMasterSp = layoutShowMasterSpRaw === undefined
      ? true
      : normalizeBool(layoutShowMasterSpRaw);

    if (layoutShowMasterSp) {
      renderElements.push(
        ...sortElementsByOrder(parseDecorativeTreeElements(
          masterTree,
          themeContext,
          masterRels,
          masterPath,
          openXmlPackage,
          contentTypes,
          tableStylesXml,
          "master"
        ), masterOrderMap)
      );
    }

    renderElements.push(
      ...sortElementsByOrder(parseDecorativeTreeElements(
        layoutTree,
        themeContext,
        layoutRels,
        layoutPath,
        openXmlPackage,
        contentTypes,
        tableStylesXml,
        "layout"
      ), layoutOrderMap)
    );

    for (const shapeNode of ensureArray(slideTree?.["p:sp"])) {
      const inherited = resolveInheritedShape(shapeNode, layoutPlaceholderMap, masterPlaceholderMap);
      elements.push(parseShapeElement(shapeNode, inherited, themeContext, masterTextStyles));
    }

    for (const connectorNode of ensureArray(slideTree?.["p:cxnSp"])) {
      elements.push(parseConnectorElement(connectorNode, themeContext));
    }

    for (const picNode of ensureArray(slideTree?.["p:pic"])) {
      elements.push(parsePictureElement(picNode, slideRels, slidePath, openXmlPackage, contentTypes));
    }

    for (const graphicFrameNode of ensureArray(slideTree?.["p:graphicFrame"])) {
      const table = parseTableElement(
        graphicFrameNode,
        themeContext,
        slideRels,
        slidePath,
        openXmlPackage,
        tableStylesXml
      );
      if (table) {
        elements.push(table);
      } else {
        unhandledNodes.push({ type: "p:graphicFrame", node: deepClone(graphicFrameNode) });
      }
    }

    for (const groupNode of ensureArray(slideTree?.["p:grpSp"])) {
      slideGroupElements.push(
        ...parseGroupTreeElements(
          groupNode,
          themeContext,
          slideRels,
          slidePath,
          openXmlPackage,
          contentTypes,
          tableStylesXml,
          "slide-group",
          identityTransform
        )
      );
      unhandledNodes.push({ type: "p:grpSp", node: deepClone(groupNode) });
    }

    for (const alternateNode of ensureArray(slideTree?.["mc:AlternateContent"])) {
      unhandledNodes.push({ type: "mc:AlternateContent", node: deepClone(alternateNode) });
    }

    const orderedSlideElements = sortElementsByOrder(elements, slideOrderMap);
    const orderedSlideRenderElements = sortElementsByOrder(
      [...elements, ...slideGroupElements],
      slideOrderMap
    );
    renderElements.push(...orderedSlideRenderElements);

    await resolveImageDataUris(renderElements);
    await resolveChartData(renderElements);
    await resolveDiagramData(renderElements);

    const background = await resolveBackgroundColor(slideXml, layoutXml, masterXml, themeContext, {
      slideRels,
      layoutRels,
      masterRels,
      themeRels,
      slidePath,
      layoutPath,
      masterPath,
      themePath,
      packageModel: openXmlPackage,
      contentTypes,
      imageDataCache: backgroundImageCache
    });

    const slideModel = {
      index: i,
      id: String(sldId?.["@_id"] || i + 1),
      relId,
      name: `Slide ${i + 1}`,
      sourcePath: slidePath,
      sourceRelsPath: relsPartPath(slidePath),
      layoutPath,
      masterPath,
      themePath,
      themeName: parsedTheme?.name || null,
      colorMap,
      background,
      elements: orderedSlideElements,
      renderElements,
      unhandledNodes,
      sourceRelationships: flattenRelMap(slideRels),
      _sourceXml: deepClone(slideXml)
    };
    slideModel._snapshot = createSlideSnapshot(slideModel);
    slides.push(slideModel);
  }

  return {
    version: "0.1",
    metadata: {
      slideSizeEmu: {
        cx: toInt(slideSizeNode?.["@_cx"], 9144000),
        cy: toInt(slideSizeNode?.["@_cy"], 6858000),
        type: slideSizeNode?.["@_type"] || "custom"
      },
      notesSizeEmu: {
        cx: toInt(notesSizeNode?.["@_cx"], 6858000),
        cy: toInt(notesSizeNode?.["@_cy"], 9144000)
      },
      firstSlideNum: toInt(presentationRoot?.["@_firstSlideNum"], 1),
      rtl: normalizeBool(presentationRoot?.["@_rtl"]),
      autoCompressPictures: normalizeBool(presentationRoot?.["@_autoCompressPictures"])
    },
    slides
  };
}

export function modelColorToOpenXmlFill(fill) {
  if (!fill || fill.type === "none" || !fill.color) {
    return { "a:noFill": "" };
  }

  const solidFill = {
    "a:srgbClr": {
      "@_val": rgbHexNoPrefix(fill.color)
    }
  };

  if (fill.alpha !== undefined && fill.alpha < 1) {
    solidFill["a:srgbClr"]["a:alpha"] = {
      "@_val": alphaToPct(fill.alpha)
    };
  }

  return {
    "a:solidFill": solidFill
  };
}

function modelLineEndToOpenXml(end) {
  if (!end) {
    return null;
  }
  const type = String(end.type || "none").toLowerCase();
  if (type === "none") {
    return {
      "@_type": "none"
    };
  }
  const node = {
    "@_type": type
  };
  if (end.width && end.width !== "med") {
    node["@_w"] = String(end.width).toLowerCase();
  }
  if (end.length && end.length !== "med") {
    node["@_len"] = String(end.length).toLowerCase();
  }
  return node;
}

export function modelLineToOpenXmlLn(line) {
  if (!line || line.color === null || line.dash === "none") {
    return {
      "a:ln": {
        "a:noFill": ""
      }
    };
  }

  const ln = {
    "@_w": toInt(line.width, 0),
    "a:solidFill": {
      "a:srgbClr": {
        "@_val": rgbHexNoPrefix(line.color)
      }
    }
  };

  if (line.alpha !== undefined && line.alpha < 1) {
    ln["a:solidFill"]["a:srgbClr"]["a:alpha"] = {
      "@_val": alphaToPct(line.alpha)
    };
  }

  if (line.dash === "cust" && ensureArray(line.customDash).length) {
    ln["a:custDash"] = {
      "a:ds": ensureArray(line.customDash).map((stop) => ({
        "@_d": toInt(stop?.d, 0),
        "@_sp": toInt(stop?.sp, 0)
      }))
    };
  } else if (line.dash && line.dash !== "solid") {
    ln["a:prstDash"] = {
      "@_val": line.dash
    };
  }

  if (line.cap && line.cap !== "flat") {
    ln["@_cap"] = line.cap;
  }

  const headEnd = modelLineEndToOpenXml(line.headEnd);
  if (headEnd) {
    ln["a:headEnd"] = headEnd;
  }
  const tailEnd = modelLineEndToOpenXml(line.tailEnd);
  if (tailEnd) {
    ln["a:tailEnd"] = tailEnd;
  }

  return {
    "a:ln": ln
  };
}

export function buildRunNode(run) {
  const style = run?.style || {};
  const rPr = {
    "@_lang": style.lang || "en-US",
    "@_sz": toInt((style.fontSizePt || 18) * 100, 1800),
    "@_b": style.bold ? "1" : "0",
    "@_i": style.italic ? "1" : "0",
    "@_u": style.underline ? "sng" : "none",
    "@_strike": style.strike ? "sngStrike" : "noStrike",
    "@_kern": toInt(style.kerning, 0)
  };

  if (style.baseline) {
    rPr["@_baseline"] = toInt(style.baseline, 0);
  }

  if (style.caps && style.caps !== "none") {
    rPr["@_cap"] = style.caps;
  }

  rPr["a:solidFill"] = {
    "a:srgbClr": {
      "@_val": rgbHexNoPrefix(style.color || "#000000")
    }
  };

  if (style.alpha !== undefined && style.alpha < 1) {
    rPr["a:solidFill"]["a:srgbClr"]["a:alpha"] = {
      "@_val": alphaToPct(style.alpha)
    };
  }

  if (style.fontFamily) {
    rPr["a:latin"] = {
      "@_typeface": style.fontFamily
    };
  }
  if (style.eastAsiaFont) {
    rPr["a:ea"] = {
      "@_typeface": style.eastAsiaFont
    };
  }
  if (style.complexScriptFont) {
    rPr["a:cs"] = {
      "@_typeface": style.complexScriptFont
    };
  }

  return {
    "a:rPr": rPr,
    "a:t": run?.text || ""
  };
}

export function buildTextBodyNode(textBody) {
  if (!textBody) {
    return null;
  }

  const paragraphs = ensureArray(textBody.paragraphs).map((paragraph) => {
    const pNode = {
      "a:pPr": {
        "@_algn": paragraph.alignment || "l",
        "@_lvl": toInt(paragraph.level, 0)
      },
      "a:r": ensureArray(paragraph.runs).map((run) => buildRunNode(run)),
      "a:endParaRPr": {
        "@_lang": "en-US"
      }
    };

    if (paragraph.marginLeft) pNode["a:pPr"]["@_marL"] = toInt(paragraph.marginLeft, 0);
    if (paragraph.marginRight) pNode["a:pPr"]["@_marR"] = toInt(paragraph.marginRight, 0);
    if (paragraph.indent) pNode["a:pPr"]["@_indent"] = toInt(paragraph.indent, 0);

    if (paragraph.spaceBefore) {
      pNode["a:pPr"]["a:spcBef"] = { "a:spcPts": { "@_val": toInt(paragraph.spaceBefore, 0) } };
    }
    if (paragraph.spaceAfter) {
      pNode["a:pPr"]["a:spcAft"] = { "a:spcPts": { "@_val": toInt(paragraph.spaceAfter, 0) } };
    }
    if (paragraph.lineSpacing) {
      pNode["a:pPr"]["a:lnSpc"] = { "a:spcPct": { "@_val": toInt(paragraph.lineSpacing, 0) } };
    }

    return pNode;
  });

  return {
    "a:bodyPr": {
      "@_anchor": textBody.verticalAlign || "t",
      "@_wrap": textBody.wrap || "square",
      "@_rtlCol": textBody.rtlCol ? "1" : "0",
      "@_fromWordArt": textBody.fromWordArt ? "1" : "0",
      "@_anchorCtr": textBody.anchorCtr ? "1" : "0",
      "@_forceAA": textBody.forceAA ? "1" : "0",
      "@_upright": textBody.upright ? "1" : "0",
      "@_numCol": toInt(textBody.numCol, 1),
      "@_lIns": toInt(textBody.leftInset, 45720),
      "@_tIns": toInt(textBody.topInset, 22860),
      "@_rIns": toInt(textBody.rightInset, 45720),
      "@_bIns": toInt(textBody.bottomInset, 22860)
    },
    "a:lstStyle": {},
    "a:p": paragraphs.length ? paragraphs : [{ "a:r": [{ "a:t": "" }], "a:endParaRPr": { "@_lang": "en-US" } }]
  };
}

function guideListToOpenXml(guides) {
  const list = ensureArray(guides).filter((gd) => gd?.name);
  if (!list.length) {
    return {};
  }
  return {
    "a:gd": list.map((gd) => ({
      "@_name": String(gd.name),
      "@_fmla": String(gd.fmla || "val 0")
    }))
  };
}

function pushNodeArray(node, key, value) {
  if (!node[key]) {
    node[key] = [];
  }
  if (!Array.isArray(node[key])) {
    node[key] = [node[key]];
  }
  node[key].push(value);
}

function pathPointNode(point) {
  return {
    "a:pt": {
      "@_x": String(point?.x ?? "0"),
      "@_y": String(point?.y ?? "0")
    }
  };
}

function customPathToOpenXml(path) {
  const node = {};
  if (path?.w) node["@_w"] = toInt(path.w, 0);
  if (path?.h) node["@_h"] = toInt(path.h, 0);
  if (path?.fill && path.fill !== "norm") node["@_fill"] = String(path.fill);
  if (path?.stroke !== undefined && String(path.stroke) !== "1") node["@_stroke"] = String(path.stroke);
  if (path?.extrusionOk !== undefined && String(path.extrusionOk) !== "1") {
    node["@_extrusionOk"] = String(path.extrusionOk);
  }

  for (const cmd of ensureArray(path?.commands)) {
    switch (String(cmd?.type || "")) {
      case "moveTo":
        pushNodeArray(node, "a:moveTo", pathPointNode(cmd));
        break;
      case "lnTo":
        pushNodeArray(node, "a:lnTo", pathPointNode(cmd));
        break;
      case "arcTo":
        pushNodeArray(node, "a:arcTo", {
          "@_wR": String(cmd?.wR ?? "0"),
          "@_hR": String(cmd?.hR ?? "0"),
          "@_stAng": String(cmd?.stAng ?? "0"),
          "@_swAng": String(cmd?.swAng ?? "0")
        });
        break;
      case "quadBezTo": {
        const points = ensureArray(cmd?.points).slice(0, 2).map((pt) => ({
          "@_x": String(pt?.x ?? "0"),
          "@_y": String(pt?.y ?? "0")
        }));
        if (points.length === 2) {
          pushNodeArray(node, "a:quadBezTo", { "a:pt": points });
        }
        break;
      }
      case "cubicBezTo": {
        const points = ensureArray(cmd?.points).slice(0, 3).map((pt) => ({
          "@_x": String(pt?.x ?? "0"),
          "@_y": String(pt?.y ?? "0")
        }));
        if (points.length === 3) {
          pushNodeArray(node, "a:cubicBezTo", { "a:pt": points });
        }
        break;
      }
      case "close":
        pushNodeArray(node, "a:close", "");
        break;
      default:
        break;
    }
  }

  return node;
}

function customGeometryToOpenXml(geometry) {
  const custGeom = deepClone(geometry?.raw || {});
  custGeom["a:avLst"] = guideListToOpenXml(geometry?.adjustValues);
  if (ensureArray(geometry?.guideValues).length) {
    custGeom["a:gdLst"] = guideListToOpenXml(geometry?.guideValues);
  }

  if (ensureArray(geometry?.paths).length) {
    const pathLst = {};
    if (geometry?.pathDefaults?.w) {
      pathLst["@_w"] = toInt(geometry.pathDefaults.w, 21600);
    }
    if (geometry?.pathDefaults?.h) {
      pathLst["@_h"] = toInt(geometry.pathDefaults.h, 21600);
    }
    pathLst["a:path"] = ensureArray(geometry.paths).map((path) => customPathToOpenXml(path));
    custGeom["a:pathLst"] = pathLst;
  }

  return custGeom;
}

function modelGeometryToOpenXml(element) {
  const geometry = element?.geometry;
  if (geometry?.kind === "cust") {
    return {
      "a:custGeom": customGeometryToOpenXml(geometry)
    };
  }

  const preset = geometry?.preset || element.shapeType || "rect";
  const avLst = guideListToOpenXml(geometry?.adjustValues);
  return {
    "a:prstGeom": {
      "@_prst": preset,
      "a:avLst": avLst
    }
  };
}

export function buildShapeSpPr(element) {
  const spPr = {
    "a:xfrm": {
      "a:off": {
        "@_x": toInt(element.x, 0),
        "@_y": toInt(element.y, 0)
      },
      "a:ext": {
        "@_cx": toInt(element.cx, 0),
        "@_cy": toInt(element.cy, 0)
      }
    }
  };

  Object.assign(spPr, modelGeometryToOpenXml(element));

  if (element.rotation) {
    spPr["a:xfrm"]["@_rot"] = toInt((element.rotation || 0) * 60000, 0);
  }
  if (element.flipH) {
    spPr["a:xfrm"]["@_flipH"] = "1";
  }
  if (element.flipV) {
    spPr["a:xfrm"]["@_flipV"] = "1";
  }

  Object.assign(spPr, modelColorToOpenXmlFill(element.fill));
  Object.assign(spPr, modelLineToOpenXmlLn(element.line));

  return spPr;
}
