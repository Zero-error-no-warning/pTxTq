import { ensureArray, first } from "../utils/object.js";
import { parseHexOrNull } from "../utils/color.js";

function extractColorNodeValue(node) {
  if (!node || typeof node !== "object") {
    return null;
  }
  if (node["a:srgbClr"]) {
    const srgb = first(node["a:srgbClr"]);
    return parseHexOrNull(srgb?.["@_val"]);
  }
  if (node["a:sysClr"]) {
    const sys = first(node["a:sysClr"]);
    return parseHexOrNull(sys?.["@_lastClr"] || sys?.["@_val"]);
  }
  return null;
}

function flattenStyleList(styleListNode) {
  if (!styleListNode || typeof styleListNode !== "object") {
    return [];
  }
  const flattened = [];
  for (const [key, value] of Object.entries(styleListNode)) {
    if (key.startsWith("@_")) {
      continue;
    }
    for (const item of ensureArray(value)) {
      flattened.push({ type: key, node: item });
    }
  }
  return flattened;
}

function parseScriptFonts(fontNode) {
  const scripts = {};
  for (const entry of ensureArray(fontNode?.["a:font"])) {
    const script = entry?.["@_script"];
    const typeface = entry?.["@_typeface"];
    if (!script || !typeface) {
      continue;
    }
    scripts[script] = typeface;
  }
  return scripts;
}

function langToThemeScript(lang) {
  const normalized = String(lang || "").toLowerCase();
  if (!normalized) {
    return null;
  }
  if (normalized.startsWith("ja")) {
    return "Jpan";
  }
  if (normalized.startsWith("ko")) {
    return "Hang";
  }
  if (normalized.startsWith("zh")) {
    if (normalized.includes("tw") || normalized.includes("hk") || normalized.includes("mo") || normalized.includes("hant")) {
      return "Hant";
    }
    return "Hans";
  }
  return null;
}

export function parseTheme(themeXml) {
  const root = themeXml?.["a:theme"];
  const elements = root?.["a:themeElements"] || {};
  const colorSchemeNode = elements?.["a:clrScheme"] || {};
  const fontSchemeNode = elements?.["a:fontScheme"] || {};
  const formatSchemeNode = elements?.["a:fmtScheme"] || {};

  const colorScheme = {};
  for (const [key, value] of Object.entries(colorSchemeNode)) {
    if (key.startsWith("@_")) {
      continue;
    }
    const logicalName = key.replace(/^a:/, "");
    const color = extractColorNodeValue(first(value));
    if (color) {
      colorScheme[logicalName] = color;
    }
  }

  const majorFont = fontSchemeNode?.["a:majorFont"] || {};
  const minorFont = fontSchemeNode?.["a:minorFont"] || {};

  const fontScheme = {
    major: {
      latin: majorFont?.["a:latin"]?.["@_typeface"] || null,
      ea: majorFont?.["a:ea"]?.["@_typeface"] || null,
      cs: majorFont?.["a:cs"]?.["@_typeface"] || null,
      scripts: parseScriptFonts(majorFont)
    },
    minor: {
      latin: minorFont?.["a:latin"]?.["@_typeface"] || null,
      ea: minorFont?.["a:ea"]?.["@_typeface"] || null,
      cs: minorFont?.["a:cs"]?.["@_typeface"] || null,
      scripts: parseScriptFonts(minorFont)
    }
  };

  const formatScheme = {
    fillStyles: flattenStyleList(formatSchemeNode?.["a:fillStyleLst"]),
    lineStyles: ensureArray(formatSchemeNode?.["a:lnStyleLst"]?.["a:ln"]),
    bgFillStyles: flattenStyleList(formatSchemeNode?.["a:bgFillStyleLst"])
  };

  return {
    name: root?.["@_name"] || null,
    colorScheme,
    fontScheme,
    formatScheme
  };
}

export function parseColorMapFromMaster(masterXml) {
  const clrMapNode = masterXml?.["p:sldMaster"]?.["p:clrMap"];
  if (!clrMapNode) {
    return null;
  }
  return {
    bg1: clrMapNode?.["@_bg1"] || "lt1",
    tx1: clrMapNode?.["@_tx1"] || "dk1",
    bg2: clrMapNode?.["@_bg2"] || "lt2",
    tx2: clrMapNode?.["@_tx2"] || "dk2",
    accent1: clrMapNode?.["@_accent1"] || "accent1",
    accent2: clrMapNode?.["@_accent2"] || "accent2",
    accent3: clrMapNode?.["@_accent3"] || "accent3",
    accent4: clrMapNode?.["@_accent4"] || "accent4",
    accent5: clrMapNode?.["@_accent5"] || "accent5",
    accent6: clrMapNode?.["@_accent6"] || "accent6",
    hlink: clrMapNode?.["@_hlink"] || "hlink",
    folHlink: clrMapNode?.["@_folHlink"] || "folHlink"
  };
}

export function parseColorMapOverride(slideXml) {
  const override = slideXml?.["p:sld"]?.["p:clrMapOvr"];
  if (!override) {
    return null;
  }
  const clrMapNode = override?.["a:overrideClrMapping"] || override?.["a:clrMap"];
  if (!clrMapNode) {
    return null;
  }
  return {
    bg1: clrMapNode?.["@_bg1"],
    tx1: clrMapNode?.["@_tx1"],
    bg2: clrMapNode?.["@_bg2"],
    tx2: clrMapNode?.["@_tx2"],
    accent1: clrMapNode?.["@_accent1"],
    accent2: clrMapNode?.["@_accent2"],
    accent3: clrMapNode?.["@_accent3"],
    accent4: clrMapNode?.["@_accent4"],
    accent5: clrMapNode?.["@_accent5"],
    accent6: clrMapNode?.["@_accent6"],
    hlink: clrMapNode?.["@_hlink"],
    folHlink: clrMapNode?.["@_folHlink"]
  };
}

export function mergeColorMap(baseMap, overrideMap) {
  return {
    bg1: overrideMap?.bg1 || baseMap?.bg1 || "lt1",
    tx1: overrideMap?.tx1 || baseMap?.tx1 || "dk1",
    bg2: overrideMap?.bg2 || baseMap?.bg2 || "lt2",
    tx2: overrideMap?.tx2 || baseMap?.tx2 || "dk2",
    accent1: overrideMap?.accent1 || baseMap?.accent1 || "accent1",
    accent2: overrideMap?.accent2 || baseMap?.accent2 || "accent2",
    accent3: overrideMap?.accent3 || baseMap?.accent3 || "accent3",
    accent4: overrideMap?.accent4 || baseMap?.accent4 || "accent4",
    accent5: overrideMap?.accent5 || baseMap?.accent5 || "accent5",
    accent6: overrideMap?.accent6 || baseMap?.accent6 || "accent6",
    hlink: overrideMap?.hlink || baseMap?.hlink || "hlink",
    folHlink: overrideMap?.folHlink || baseMap?.folHlink || "folHlink"
  };
}

export function createThemeContext(theme, colorMap) {
  const effectiveColorMap = mergeColorMap(colorMap, null);

  const resolveSchemeColor = (schemeKey) => {
    if (!schemeKey) {
      return null;
    }
    if (schemeKey === "phClr") {
      return theme?.colorScheme?.tx1 || "#000000";
    }
    const mapped = effectiveColorMap[schemeKey] || schemeKey;
    return theme?.colorScheme?.[mapped] || theme?.colorScheme?.[schemeKey] || null;
  };

  const resolveThemeFont = (token, lang = null) => {
    if (!token || typeof token !== "string") {
      return null;
    }
    if (!token.startsWith("+")) {
      return token;
    }

    const lower = token.toLowerCase();
    const scriptKey = langToThemeScript(lang);
    const pickScriptFont = (bucket) => {
      if (!bucket) {
        return null;
      }
      if (scriptKey && bucket?.scripts?.[scriptKey]) {
        return bucket.scripts[scriptKey];
      }
      return null;
    };

    if (lower.includes("mn")) {
      if (lower.includes("ea")) return theme?.fontScheme?.minor?.ea || pickScriptFont(theme?.fontScheme?.minor) || theme?.fontScheme?.minor?.latin || null;
      if (lower.includes("cs")) return theme?.fontScheme?.minor?.cs || pickScriptFont(theme?.fontScheme?.minor) || theme?.fontScheme?.minor?.latin || null;
      return theme?.fontScheme?.minor?.latin || null;
    }
    if (lower.includes("mj")) {
      if (lower.includes("ea")) return theme?.fontScheme?.major?.ea || pickScriptFont(theme?.fontScheme?.major) || theme?.fontScheme?.major?.latin || null;
      if (lower.includes("cs")) return theme?.fontScheme?.major?.cs || pickScriptFont(theme?.fontScheme?.major) || theme?.fontScheme?.major?.latin || null;
      return theme?.fontScheme?.major?.latin || null;
    }
    return null;
  };

  return {
    theme,
    colorMap: effectiveColorMap,
    resolveSchemeColor,
    resolveThemeFont,
    getFillStyle(index) {
      if (!index) return null;
      return theme?.formatScheme?.fillStyles?.[index - 1] || null;
    },
    getLineStyle(index) {
      if (!index) return null;
      return theme?.formatScheme?.lineStyles?.[index - 1] || null;
    },
    getBgFillStyle(index) {
      if (!index) return null;
      const numeric = Number.parseInt(String(index), 10);
      if (!Number.isFinite(numeric)) {
        return null;
      }
      const normalized = numeric >= 1001 ? numeric - 1001 : numeric - 1;
      if (normalized < 0) {
        return null;
      }
      return theme?.formatScheme?.bgFillStyles?.[normalized] || null;
    }
  };
}
