import { clamp, toInt } from "./object.js";

const PRESET_COLORS = {
  black: "000000",
  white: "FFFFFF",
  red: "FF0000",
  green: "00FF00",
  blue: "0000FF",
  yellow: "FFFF00",
  gray: "808080",
  grey: "808080",
  orange: "FFA500",
  purple: "800080",
  aqua: "00FFFF",
  fuchsia: "FF00FF",
  navy: "000080",
  maroon: "800000",
  lime: "00FF00",
  teal: "008080",
  silver: "C0C0C0"
};

function normalizeHex(color) {
  if (!color) {
    return null;
  }
  let hex = String(color).trim().replace(/^#/, "").toUpperCase();
  if (hex.length === 3) {
    hex = `${hex[0]}${hex[0]}${hex[1]}${hex[1]}${hex[2]}${hex[2]}`;
  }
  if (!/^[0-9A-F]{6}$/.test(hex)) {
    return null;
  }
  return hex;
}

function hexToRgb(hex) {
  const normalized = normalizeHex(hex);
  if (!normalized) {
    return { r: 0, g: 0, b: 0 };
  }
  return {
    r: Number.parseInt(normalized.slice(0, 2), 16),
    g: Number.parseInt(normalized.slice(2, 4), 16),
    b: Number.parseInt(normalized.slice(4, 6), 16)
  };
}

function rgbToHex({ r, g, b }) {
  const toHex = (n) => clamp(Math.round(n), 0, 255).toString(16).padStart(2, "0");
  return `${toHex(r)}${toHex(g)}${toHex(b)}`.toUpperCase();
}

function applyTint(rgb, tint) {
  const ratio = clamp(tint / 100000, 0, 1);
  return {
    r: rgb.r + (255 - rgb.r) * ratio,
    g: rgb.g + (255 - rgb.g) * ratio,
    b: rgb.b + (255 - rgb.b) * ratio
  };
}

function applyShade(rgb, shade) {
  const ratio = clamp(shade / 100000, 0, 1);
  return {
    r: rgb.r * ratio,
    g: rgb.g * ratio,
    b: rgb.b * ratio
  };
}

function applyLum(rgb, lumMod, lumOff) {
  if (lumMod === null && lumOff === null) {
    return rgb;
  }
  const hsl = rgbToHsl(rgb);
  const mod = lumMod === null ? 1 : clamp(lumMod / 100000, 0, 2);
  const off = lumOff === null ? 0 : clamp(lumOff / 100000, -1, 1);
  hsl.l = clamp(hsl.l * mod + off, 0, 1);
  return hslToRgb(hsl);
}

function applySat(rgb, satMod, satOff) {
  if (satMod === null && satOff === null) {
    return rgb;
  }
  const hsl = rgbToHsl(rgb);
  const mod = satMod === null ? 1 : clamp(satMod / 100000, 0, 2);
  const off = satOff === null ? 0 : clamp(satOff / 100000, -1, 1);
  hsl.s = clamp(hsl.s * mod + off, 0, 1);
  return hslToRgb(hsl);
}

function applyHue(rgb, hueMod, hueOff) {
  if (hueMod === null && hueOff === null) {
    return rgb;
  }
  const hsl = rgbToHsl(rgb);
  if (hueMod !== null) {
    hsl.h = (hsl.h * (hueMod / 100000)) % 360;
  }
  if (hueOff !== null) {
    hsl.h = (hsl.h + (hueOff / 60000)) % 360;
  }
  if (hsl.h < 0) {
    hsl.h += 360;
  }
  return hslToRgb(hsl);
}

function rgbToHsl({ r, g, b }) {
  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  let h = 0;
  let s = 0;
  const l = (max + min) / 2;

  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case rn:
        h = (gn - bn) / d + (gn < bn ? 6 : 0);
        break;
      case gn:
        h = (bn - rn) / d + 2;
        break;
      default:
        h = (rn - gn) / d + 4;
        break;
    }
    h *= 60;
  }
  return { h, s, l };
}

function hslToRgb({ h, s, l }) {
  if (s === 0) {
    const gray = l * 255;
    return { r: gray, g: gray, b: gray };
  }

  const hue2rgb = (p, q, t) => {
    let tt = t;
    if (tt < 0) tt += 1;
    if (tt > 1) tt -= 1;
    if (tt < 1 / 6) return p + (q - p) * 6 * tt;
    if (tt < 1 / 2) return q;
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
    return p;
  };

  const hn = h / 360;
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;

  return {
    r: hue2rgb(p, q, hn + 1 / 3) * 255,
    g: hue2rgb(p, q, hn) * 255,
    b: hue2rgb(p, q, hn - 1 / 3) * 255
  };
}

function collectTransform(node, key) {
  const candidate = node[key];
  if (!candidate) {
    return null;
  }
  if (Array.isArray(candidate)) {
    const first = candidate[0];
    return toInt(first?.["@_val"], null);
  }
  return toInt(candidate["@_val"], null);
}

function normalizePlaceholderColor(placeholderColor) {
  if (!placeholderColor) {
    return null;
  }
  if (typeof placeholderColor === "string") {
    const color = normalizeHex(placeholderColor);
    return color ? { color: `#${color}`, alpha: 1 } : null;
  }
  const color = normalizeHex(placeholderColor.color);
  if (!color) {
    return null;
  }
  return {
    color: `#${color}`,
    alpha: clamp(placeholderColor.alpha ?? 1, 0, 1)
  };
}

export function parseHexOrNull(value) {
  const normalized = normalizeHex(value);
  return normalized ? `#${normalized}` : null;
}

export function hexToCss(hex, alpha = 1) {
  const normalized = normalizeHex(hex) || "000000";
  const clampedAlpha = clamp(alpha, 0, 1);
  if (clampedAlpha >= 1) {
    return `#${normalized}`;
  }
  return `rgba(${Number.parseInt(normalized.slice(0, 2), 16)}, ${Number.parseInt(normalized.slice(2, 4), 16)}, ${Number.parseInt(normalized.slice(4, 6), 16)}, ${clampedAlpha.toFixed(4)})`;
}

export function applyOpenXmlColorTransform(baseHex, node) {
  let rgb = hexToRgb(baseHex);
  const tint = collectTransform(node, "a:tint");
  const shade = collectTransform(node, "a:shade");
  const lumMod = collectTransform(node, "a:lumMod");
  const lumOff = collectTransform(node, "a:lumOff");
  const satMod = collectTransform(node, "a:satMod");
  const satOff = collectTransform(node, "a:satOff");
  const hueMod = collectTransform(node, "a:hueMod");
  const hueOff = collectTransform(node, "a:hueOff");
  const alpha = collectTransform(node, "a:alpha");

  if (tint !== null) {
    rgb = applyTint(rgb, tint);
  }
  if (shade !== null) {
    rgb = applyShade(rgb, shade);
  }
  rgb = applyLum(rgb, lumMod, lumOff);
  rgb = applySat(rgb, satMod, satOff);
  rgb = applyHue(rgb, hueMod, hueOff);

  return {
    color: `#${rgbToHex(rgb)}`,
    alpha: alpha === null ? 1 : clamp(alpha / 100000, 0, 1)
  };
}

export function resolveDrawingColor(colorContainer, themeContext, placeholderColor = null) {
  if (!colorContainer || typeof colorContainer !== "object") {
    return null;
  }

  if (colorContainer["a:srgbClr"]) {
    const node = Array.isArray(colorContainer["a:srgbClr"])
      ? colorContainer["a:srgbClr"][0]
      : colorContainer["a:srgbClr"];
    const val = normalizeHex(node?.["@_val"]);
    if (!val) {
      return null;
    }
    return applyOpenXmlColorTransform(val, node);
  }

  if (colorContainer["a:schemeClr"]) {
    const node = Array.isArray(colorContainer["a:schemeClr"])
      ? colorContainer["a:schemeClr"][0]
      : colorContainer["a:schemeClr"];
    const scheme = node?.["@_val"];
    if (!scheme) {
      return null;
    }

    let resolvedBase = null;
    if (scheme === "phClr") {
      resolvedBase = normalizePlaceholderColor(placeholderColor)?.color || null;
    }
    if (!resolvedBase && themeContext?.resolveSchemeColor) {
      resolvedBase = themeContext.resolveSchemeColor(scheme);
    }
    if (!resolvedBase) {
      return null;
    }

    const resolved = applyOpenXmlColorTransform(resolvedBase.replace(/^#/, ""), node);
    const placeholderAlpha = normalizePlaceholderColor(placeholderColor)?.alpha ?? 1;
    if (!resolved) {
      return null;
    }
    return {
      color: resolved.color,
      alpha: clamp((resolved.alpha ?? 1) * placeholderAlpha, 0, 1)
    };
  }

  if (colorContainer["a:prstClr"]) {
    const node = Array.isArray(colorContainer["a:prstClr"])
      ? colorContainer["a:prstClr"][0]
      : colorContainer["a:prstClr"];
    const val = node?.["@_val"];
    const preset = val ? PRESET_COLORS[val.toLowerCase()] : null;
    if (!preset) {
      return null;
    }
    return applyOpenXmlColorTransform(preset, node);
  }

  if (colorContainer["a:sysClr"]) {
    const node = Array.isArray(colorContainer["a:sysClr"])
      ? colorContainer["a:sysClr"][0]
      : colorContainer["a:sysClr"];
    const val = normalizeHex(node?.["@_lastClr"] || node?.["@_val"]);
    if (!val) {
      return null;
    }
    return applyOpenXmlColorTransform(val, node);
  }

  return null;
}

export function rgbHexNoPrefix(color) {
  const normalized = normalizeHex(color);
  return normalized || "000000";
}

export function alphaToPct(alpha = 1) {
  return Math.round(clamp(alpha, 0, 1) * 100000);
}

export function lighten(hex, amount) {
  const rgb = hexToRgb(hex);
  return `#${rgbToHex(applyTint(rgb, amount * 100000))}`;
}

export function darken(hex, amount) {
  const rgb = hexToRgb(hex);
  return `#${rgbToHex(applyShade(rgb, (1 - amount) * 100000))}`;
}
