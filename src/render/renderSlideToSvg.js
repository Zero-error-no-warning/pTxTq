import { emuToPx } from "../utils/units.js";
import { clamp, ensureArray } from "../utils/object.js";
import {
  buildGeometryVars,
  evalGeomFormula,
  fitGeometryExtents,
  resolveOoxmlArcFromCurrentPoint,
  splitArcSweep
} from "../utils/geometry.js";
import { resolvePresetShapeGeometry } from "../utils/presetShapeGeometry.js";
import { buildPresetShapeParts } from "../utils/presetShape.js";

const EMU_PER_PT = 12700;

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function colorWithAlpha(color, alpha = 1) {
  if (!color) {
    return "none";
  }
  if (alpha >= 1) {
    return color;
  }
  const normalized = color.replace(/^#/, "");
  const r = Number.parseInt(normalized.slice(0, 2), 16);
  const g = Number.parseInt(normalized.slice(2, 4), 16);
  const b = Number.parseInt(normalized.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${clamp(alpha, 0, 1).toFixed(4)})`;
}

function alignmentToTextAnchor(alignment) {
  switch (alignment) {
    case "ctr":
      return "middle";
    case "r":
      return "end";
    default:
      return "start";
  }
}

function transformAttr(element) {
  if (!element.rotation && !element.flipH && !element.flipV) {
    return "";
  }
  const cx = element.x + element.cx / 2;
  const cy = element.y + element.cy / 2;
  const pieces = [`translate(${cx} ${cy})`];
  if (element.rotation) {
    pieces.push(`rotate(${element.rotation})`);
  }
  pieces.push(`scale(${element.flipH ? -1 : 1} ${element.flipV ? -1 : 1})`);
  pieces.push(`translate(${-cx} ${-cy})`);
  return ` transform=\"${pieces.join(" ")}\"`;
}

function isLineLikeShapeType(shapeType) {
  const normalized = String(shapeType || "").toLowerCase();
  return normalized === "line"
    || normalized === "straightconnector1"
    || normalized === "bentconnector2"
    || normalized === "bentconnector3"
    || normalized === "bentconnector4"
    || normalized === "bentconnector5"
    || normalized === "curvedconnector2"
    || normalized === "curvedconnector3"
    || normalized === "curvedconnector4"
    || normalized === "curvedconnector5";
}

const DASH_PRESET_FACTORS = {
  dot: [1, 2],
  sysdot: [1, 2],
  dash: [4, 3],
  sysdash: [4, 3],
  lgdash: [8, 4],
  dashdot: [4, 3, 1, 3],
  sysdashdot: [4, 3, 1, 3],
  sysdashdotdot: [4, 3, 1, 3, 1, 3],
  lgdashdot: [8, 4, 1, 4],
  lgdashdotdot: [8, 4, 1, 4, 1, 4]
};

function lineCapToSvg(cap) {
  switch (String(cap || "flat").toLowerCase()) {
    case "rnd":
      return "round";
    case "sq":
      return "square";
    default:
      return "butt";
  }
}

function lineJoinType(join) {
  return String(join?.type || join || "miter").toLowerCase();
}

function lineJoinToSvg(join) {
  switch (lineJoinType(join)) {
    case "round":
      return "round";
    case "bevel":
      return "bevel";
    default:
      return "miter";
  }
}

function lineMiterLimit(join) {
  const numeric = Number(join?.limit);
  if (!Number.isFinite(numeric) || numeric <= 0) {
    return 10;
  }
  return Math.max(1, numeric > 1000 ? numeric / 100000 : numeric);
}

function lineEndScale(value, sm, med, lg) {
  switch (String(value || "med").toLowerCase()) {
    case "sm":
      return sm;
    case "lg":
      return lg;
    default:
      return med;
  }
}

function toCustomDashEmu(value, width) {
  const numeric = Number.parseInt(String(value ?? 0), 10);
  if (!Number.isFinite(numeric) || numeric <= 0) {
    return Math.max(1, width);
  }
  return Math.max(1, (numeric / 100000) * width);
}

function lineDashArray(line, strokeWidth) {
  if (!line || line.dash === "none") {
    return "";
  }

  const effectiveWidth = Math.max(1, strokeWidth);
  if (String(line.dash || "").toLowerCase() === "cust") {
    const pattern = [];
    for (const stop of ensureArray(line.customDash)) {
      pattern.push(toCustomDashEmu(stop?.d, effectiveWidth));
      pattern.push(toCustomDashEmu(stop?.sp, effectiveWidth));
    }
    return pattern.join(",");
  }

  const factors = DASH_PRESET_FACTORS[String(line.dash || "solid").toLowerCase()];
  if (!factors) {
    return "";
  }
  return factors.map((factor) => Math.max(1, factor * effectiveWidth)).join(",");
}

function lineStrokeProps(line) {
  if (!line?.color || line?.dash === "none") {
    return null;
  }
  const strokeWidth = Math.max(1, Number(line?.width) || 0);
  return {
    stroke: colorWithAlpha(line.color, line.alpha ?? 1),
    strokeWidth,
    lineCap: lineCapToSvg(line.cap),
    lineJoin: lineJoinToSvg(line.join),
    miterLimit: lineMiterLimit(line.join),
    dashArray: lineDashArray(line, strokeWidth)
  };
}

function shapeStrokeAttrs(line) {
  const stroke = lineStrokeProps(line);
  if (!stroke) {
    return "stroke=\"none\"";
  }
  const dashAttr = stroke.dashArray ? ` stroke-dasharray=\"${stroke.dashArray}\"` : "";
  return `stroke=\"${stroke.stroke}\" stroke-width=\"${stroke.strokeWidth}\" stroke-linecap=\"${stroke.lineCap}\" stroke-linejoin=\"${stroke.lineJoin}\" stroke-miterlimit=\"${stroke.miterLimit}\"${dashAttr}`;
}

function svgSafeId(value) {
  return String(value || "").replace(/[^A-Za-z0-9_-]/g, "_");
}

function shapeGradientFill(element) {
  const fill = element?.fill;
  if (!fill || fill.type === "none") {
    return { defs: "", fill: "none" };
  }
  if (fill.type !== "gradient" || !ensureArray(fill.stops).length) {
    return {
      defs: "",
      fill: colorWithAlpha(fill.color, fill.alpha ?? 1)
    };
  }

  const gradientId = svgSafeId(`shapeGrad-${element?.id || "shape"}-${Math.round(element?.x || 0)}-${Math.round(element?.y || 0)}`);
  const stops = ensureArray(fill.stops).map((stop) => (
    `<stop offset=\"${clamp((Number(stop?.pos) || 0) / 1000, 0, 100)}%\" stop-color=\"${colorWithAlpha(stop?.color || "#FFFFFF", stop?.alpha ?? 1)}\" />`
  )).join("");

  let defs = "";
  if (fill.gradientType === "path" && fill.path === "circle") {
    defs = `<defs><radialGradient id=\"${gradientId}\" cx=\"50%\" cy=\"50%\" r=\"75%\">${stops}</radialGradient></defs>`;
  } else {
    const angle = Number(fill.angle) || 0;
    const rad = (angle * Math.PI) / 180;
    const x1 = 50 - Math.cos(rad) * 50;
    const y1 = 50 - Math.sin(rad) * 50;
    const x2 = 50 + Math.cos(rad) * 50;
    const y2 = 50 + Math.sin(rad) * 50;
    defs = `<defs><linearGradient id=\"${gradientId}\" x1=\"${x1}%\" y1=\"${y1}%\" x2=\"${x2}%\" y2=\"${y2}%\">${stops}</linearGradient></defs>`;
  }

  return {
    defs,
    fill: `url(#${gradientId})`
  };
}

function pointList(points) {
  return points.map((p) => `${p.x},${p.y}`).join(" ");
}

function loopsPathData(loops) {
  const parts = [];
  for (const loop of ensureArray(loops)) {
    const points = ensureArray(loop);
    if (!points.length) {
      continue;
    }
    parts.push(`M ${points[0].x} ${points[0].y}`);
    for (let i = 1; i < points.length; i += 1) {
      parts.push(`L ${points[i].x} ${points[i].y}`);
    }
    parts.push("Z");
  }
  return parts.join(" ");
}

function renderPresetShapePart(part, commonAttrs, transformAttrs) {
  switch (part?.kind) {
    case "polygon":
      return `<polygon points=\"${pointList(ensureArray(part.points))}\" ${commonAttrs}${transformAttrs} />`;
    case "loops":
      return `<path d=\"${loopsPathData(part.loops)}\" ${commonAttrs}${transformAttrs} />`;
    case "polyline":
      return `<polyline points=\"${pointList(ensureArray(part.points))}\" fill=\"none\" ${commonAttrs}${transformAttrs} />`;
    case "ellipse":
      return `<ellipse cx=\"${part.cx}\" cy=\"${part.cy}\" rx=\"${part.rx}\" ry=\"${part.ry}\" ${commonAttrs}${transformAttrs} />`;
    case "roundRect":
      return `<rect x=\"${part.x}\" y=\"${part.y}\" width=\"${part.w}\" height=\"${part.h}\" rx=\"${part.r}\" ry=\"${part.r}\" ${commonAttrs}${transformAttrs} />`;
    case "rect":
      return `<rect x=\"${part.x}\" y=\"${part.y}\" width=\"${part.w}\" height=\"${part.h}\" ${commonAttrs}${transformAttrs} />`;
    default:
      return "";
  }
}

function rotatePoint(point, center, rotationDeg) {
  if (!rotationDeg) {
    return point;
  }
  const rad = (rotationDeg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const dx = point.x - center.x;
  const dy = point.y - center.y;
  return {
    x: center.x + dx * cos - dy * sin,
    y: center.y + dx * sin + dy * cos
  };
}

function lineEndpoints(element) {
  const start = {
    x: element.flipH ? element.x + element.cx : element.x,
    y: element.flipV ? element.y + element.cy : element.y
  };
  const end = {
    x: element.flipH ? element.x : element.x + element.cx,
    y: element.flipV ? element.y : element.y + element.cy
  };
  if (!element.rotation) {
    return { start, end };
  }
  const center = {
    x: element.x + element.cx / 2,
    y: element.y + element.cy / 2
  };
  return {
    start: rotatePoint(start, center, element.rotation),
    end: rotatePoint(end, center, element.rotation)
  };
}

function presetAdjustValue(element, name, fallbackValue) {
  const gd = ensureArray(element?.geometry?.adjustValues).find((entry) => String(entry?.name || "").toLowerCase() === String(name || "").toLowerCase());
  const formula = String(gd?.fmla || "").trim();
  if (!formula) {
    return fallbackValue;
  }
  const match = formula.match(/^val\s+(-?\d+)$/i);
  if (!match) {
    return fallbackValue;
  }
  return clamp(Number.parseInt(match[1], 10), 0, 100000);
}

function shapePoints(shapeType, element) {
  switch (shapeType) {
    case "triangle":
      {
        const adj = presetAdjustValue(element, "adj", 50000) / 100000;
        return [
          { x: element.x + element.cx * adj, y: element.y },
          { x: element.x + element.cx, y: element.y + element.cy },
          { x: element.x, y: element.y + element.cy }
        ];
      }
    case "rttriangle":
      return [
        { x: element.x, y: element.y },
        { x: element.x + element.cx, y: element.y + element.cy },
        { x: element.x, y: element.y + element.cy }
      ];
    case "diamond":
      return [
        { x: element.x + element.cx / 2, y: element.y },
        { x: element.x + element.cx, y: element.y + element.cy / 2 },
        { x: element.x + element.cx / 2, y: element.y + element.cy },
        { x: element.x, y: element.y + element.cy / 2 }
      ];
    case "parallelogram": {
      const offset = element.cx * 0.2;
      return [
        { x: element.x + offset, y: element.y },
        { x: element.x + element.cx, y: element.y },
        { x: element.x + element.cx - offset, y: element.y + element.cy },
        { x: element.x, y: element.y + element.cy }
      ];
    }
    case "trapezoid": {
      const inset = element.cx * 0.18;
      return [
        { x: element.x + inset, y: element.y },
        { x: element.x + element.cx - inset, y: element.y },
        { x: element.x + element.cx, y: element.y + element.cy },
        { x: element.x, y: element.y + element.cy }
      ];
    }
    case "pentagon":
      return [
        { x: element.x + element.cx * 0.5, y: element.y },
        { x: element.x + element.cx, y: element.y + element.cy * 0.38 },
        { x: element.x + element.cx * 0.82, y: element.y + element.cy },
        { x: element.x + element.cx * 0.18, y: element.y + element.cy },
        { x: element.x, y: element.y + element.cy * 0.38 }
      ];
    case "hexagon":
      return [
        { x: element.x + element.cx * 0.25, y: element.y },
        { x: element.x + element.cx * 0.75, y: element.y },
        { x: element.x + element.cx, y: element.y + element.cy * 0.5 },
        { x: element.x + element.cx * 0.75, y: element.y + element.cy },
        { x: element.x + element.cx * 0.25, y: element.y + element.cy },
        { x: element.x, y: element.y + element.cy * 0.5 }
      ];
    case "chevron":
      return [
        { x: element.x, y: element.y },
        { x: element.x + element.cx * 0.6, y: element.y },
        { x: element.x + element.cx, y: element.y + element.cy * 0.5 },
        { x: element.x + element.cx * 0.6, y: element.y + element.cy },
        { x: element.x, y: element.y + element.cy },
        { x: element.x + element.cx * 0.4, y: element.y + element.cy * 0.5 }
      ];
    default:
      return null;
  }
}

function pathStrokeEnabled(path) {
  const value = String(path?.stroke ?? "1").toLowerCase();
  return value !== "0" && value !== "false" && value !== "off";
}

function pathFillEnabled(path) {
  const value = String(path?.fill ?? "norm").toLowerCase();
  return value !== "none";
}

function mapGeometryPoint(element, pathW, pathH, x, y) {
  const sx = pathW ? x / pathW : 0;
  const sy = pathH ? y / pathH : 0;
  return {
    x: element.x + sx * element.cx,
    y: element.y + sy * element.cy
  };
}

function buildCustomGeometryPathData(element, geometry, path) {
  const pathW = Math.max(1, Number(path?.w || geometry?.pathDefaults?.w || 21600));
  const pathH = Math.max(1, Number(path?.h || geometry?.pathDefaults?.h || 21600));
  const vars = buildGeometryVars(geometry, pathW, pathH, fitGeometryExtents(pathW, pathH, element.cx, element.cy));

  const parts = [];
  let currentRaw = null;
  let subPathStartRaw = null;
  for (const cmd of ensureArray(path?.commands)) {
    const type = String(cmd?.type || "");
    if (type === "moveTo" || type === "lnTo") {
      const rawX = evalGeomFormula(cmd?.x || "0", vars);
      const rawY = evalGeomFormula(cmd?.y || "0", vars);
      const p = mapGeometryPoint(element, pathW, pathH, rawX, rawY);
      parts.push(`${type === "moveTo" ? "M" : "L"} ${p.x} ${p.y}`);
      currentRaw = { x: rawX, y: rawY };
      if (type === "moveTo") {
        subPathStartRaw = { ...currentRaw };
      }
      continue;
    }

    if (type === "quadBezTo") {
      const pts = ensureArray(cmd?.points);
      if (pts.length >= 2) {
        const cp = mapGeometryPoint(
          element,
          pathW,
          pathH,
          evalGeomFormula(pts[0]?.x || "0", vars),
          evalGeomFormula(pts[0]?.y || "0", vars)
        );
        const ep = mapGeometryPoint(
          element,
          pathW,
          pathH,
          evalGeomFormula(pts[1]?.x || "0", vars),
          evalGeomFormula(pts[1]?.y || "0", vars)
        );
        parts.push(`Q ${cp.x} ${cp.y} ${ep.x} ${ep.y}`);
        currentRaw = {
          x: evalGeomFormula(pts[1]?.x || "0", vars),
          y: evalGeomFormula(pts[1]?.y || "0", vars)
        };
      }
      continue;
    }

    if (type === "cubicBezTo") {
      const pts = ensureArray(cmd?.points);
      if (pts.length >= 3) {
        const cp1 = mapGeometryPoint(
          element,
          pathW,
          pathH,
          evalGeomFormula(pts[0]?.x || "0", vars),
          evalGeomFormula(pts[0]?.y || "0", vars)
        );
        const cp2 = mapGeometryPoint(
          element,
          pathW,
          pathH,
          evalGeomFormula(pts[1]?.x || "0", vars),
          evalGeomFormula(pts[1]?.y || "0", vars)
        );
        const ep = mapGeometryPoint(
          element,
          pathW,
          pathH,
          evalGeomFormula(pts[2]?.x || "0", vars),
          evalGeomFormula(pts[2]?.y || "0", vars)
        );
        parts.push(`C ${cp1.x} ${cp1.y} ${cp2.x} ${cp2.y} ${ep.x} ${ep.y}`);
        currentRaw = {
          x: evalGeomFormula(pts[2]?.x || "0", vars),
          y: evalGeomFormula(pts[2]?.y || "0", vars)
        };
      }
      continue;
    }

    if (type === "arcTo" && currentRaw) {
      const rxRaw = Math.abs(evalGeomFormula(cmd?.wR || "0", vars));
      const ryRaw = Math.abs(evalGeomFormula(cmd?.hR || "0", vars));
      const stAng = evalGeomFormula(cmd?.stAng || "0", vars);
      const swAng = evalGeomFormula(cmd?.swAng || "0", vars);
      const arc = resolveOoxmlArcFromCurrentPoint(currentRaw.x, currentRaw.y, rxRaw, ryRaw, stAng, swAng);
      if (arc) {
        const rx = Math.abs((arc.rx / pathW) * element.cx);
        const ry = Math.abs((arc.ry / pathH) * element.cy);
        const chunks = splitArcSweep(arc.sweepParam);

        let segStart = arc.startParam;
        for (const sweepChunk of chunks) {
          const segEnd = segStart + sweepChunk;
          const endRaw = {
            x: arc.cx + arc.rx * Math.cos(segEnd),
            y: arc.cy + arc.ry * Math.sin(segEnd)
          };
          const end = mapGeometryPoint(element, pathW, pathH, endRaw.x, endRaw.y);
          const largeArcFlag = Math.abs(sweepChunk) > Math.PI ? 1 : 0;
          const sweepFlag = sweepChunk >= 0 ? 1 : 0;
          parts.push(`A ${rx} ${ry} 0 ${largeArcFlag} ${sweepFlag} ${end.x} ${end.y}`);
          currentRaw = endRaw;
          segStart = segEnd;
        }
      }
      continue;
    }

    if (type === "close") {
      parts.push("Z");
      if (subPathStartRaw) {
        currentRaw = { ...subPathStartRaw };
      }
    }
  }

  return parts.join(" ");
}

function renderCustomGeometryPrimitive(element, options = {}) {
  const geometry = element?.geometry;
  const paths = ensureArray(geometry?.paths);
  if (!paths.length) {
    return "";
  }

  const fillInfo = options.forceNoFill === true
    ? { defs: "", fill: "none" }
    : shapeGradientFill(element);
  const pieces = [];
  const forceNoFill = options.forceNoFill === true;

  for (const path of paths) {
    const d = buildCustomGeometryPathData(element, geometry, path);
    if (!d) {
      continue;
    }
    const fill = forceNoFill || !pathFillEnabled(path) ? "none" : fillInfo.fill;
    const strokeAttrs = pathStrokeEnabled(path) ? shapeStrokeAttrs(element.line) : "stroke=\"none\"";
    pieces.push(`<path d=\"${d}\" fill=\"${fill}\" ${strokeAttrs}${transformAttr(element)} />`);
  }

  return `${fillInfo.defs}${pieces.join("")}`;
}

function resolveRenderableGeometry(element) {
  if (element?.geometry?.kind === "cust") {
    return element.geometry;
  }
  if (element?.geometry?.kind === "prst") {
    return resolvePresetShapeGeometry(element.shapeType || element.geometry?.preset, element.geometry);
  }
  return null;
}

function renderLineEnd(lineEnd, tip, other, strokeColor, strokeWidth) {
  const endType = String(lineEnd?.type || "none").toLowerCase();
  if (!lineEnd || endType === "none") {
    return "";
  }

  const dx = other.x - tip.x;
  const dy = other.y - tip.y;
  const length = Math.hypot(dx, dy);
  if (length <= 0.01) {
    return "";
  }

  const ux = dx / length;
  const uy = dy / length;
  const px = -uy;
  const py = ux;

  const markerLength = Math.max(
    strokeWidth + 2,
    strokeWidth * lineEndScale(lineEnd.length, 2.2, 3.2, 4.6)
  );
  const markerHalfWidth = Math.max(
    strokeWidth * 0.6,
    (strokeWidth * lineEndScale(lineEnd.width, 1.6, 2.5, 3.3)) / 2
  );

  const backX = tip.x + ux * markerLength;
  const backY = tip.y + uy * markerLength;
  const left = { x: backX + px * markerHalfWidth, y: backY + py * markerHalfWidth };
  const right = { x: backX - px * markerHalfWidth, y: backY - py * markerHalfWidth };

  switch (endType) {
    case "arrow":
      return `<polyline points=\"${pointList([left, tip, right])}\" fill=\"none\" stroke=\"${strokeColor}\" stroke-width=\"${Math.max(1, strokeWidth * 0.9)}\" stroke-linecap=\"round\" stroke-linejoin=\"round\" />`;
    case "stealth": {
      const notch = {
        x: tip.x + ux * markerLength * 0.38,
        y: tip.y + uy * markerLength * 0.38
      };
      return `<polygon points=\"${pointList([tip, left, notch, right])}\" fill=\"${strokeColor}\" stroke=\"none\" />`;
    }
    case "diamond": {
      const mid = {
        x: tip.x + ux * markerLength * 0.5,
        y: tip.y + uy * markerLength * 0.5
      };
      return `<polygon points=\"${pointList([
        tip,
        { x: mid.x + px * markerHalfWidth, y: mid.y + py * markerHalfWidth },
        { x: backX, y: backY },
        { x: mid.x - px * markerHalfWidth, y: mid.y - py * markerHalfWidth }
      ])}\" fill=\"${strokeColor}\" stroke=\"none\" />`;
    }
    case "oval": {
      const cx = tip.x + ux * markerLength * 0.55;
      const cy = tip.y + uy * markerLength * 0.55;
      const rx = Math.max(1, markerHalfWidth * 0.95);
      const ry = Math.max(1, markerHalfWidth * 0.75);
      const angle = (Math.atan2(uy, ux) * 180) / Math.PI;
      return `<ellipse cx=\"${cx}\" cy=\"${cy}\" rx=\"${rx}\" ry=\"${ry}\" fill=\"${strokeColor}\" stroke=\"none\" transform=\"rotate(${angle} ${cx} ${cy})\" />`;
    }
    default:
      return `<polygon points=\"${pointList([tip, left, right])}\" fill=\"${strokeColor}\" stroke=\"none\" />`;
  }
}

function renderLinePrimitive(element) {
  const renderGeometry = resolveRenderableGeometry(element);
  const renderElement = renderGeometry ? { ...element, geometry: renderGeometry } : element;
  if (renderGeometry?.kind === "cust") {
    const custom = renderCustomGeometryPrimitive(renderElement, { forceNoFill: true });
    if (custom) {
      return custom;
    }
  }

  const stroke = lineStrokeProps(renderElement.line);
  if (!stroke) {
    return "";
  }

  const { start, end } = lineEndpoints(renderElement);
  const dashAttr = stroke.dashArray ? ` stroke-dasharray=\"${stroke.dashArray}\"` : "";
  const lineNode = `<line x1=\"${start.x}\" y1=\"${start.y}\" x2=\"${end.x}\" y2=\"${end.y}\" stroke=\"${stroke.stroke}\" stroke-width=\"${stroke.strokeWidth}\" stroke-linecap=\"${stroke.lineCap}\" stroke-linejoin=\"${stroke.lineJoin}\" stroke-miterlimit=\"${stroke.miterLimit}\"${dashAttr} />`;
  const headNode = renderLineEnd(renderElement.line?.headEnd, start, end, stroke.stroke, stroke.strokeWidth);
  const tailNode = renderLineEnd(renderElement.line?.tailEnd, end, start, stroke.stroke, stroke.strokeWidth);
  return `${lineNode}${headNode}${tailNode}`;
}

function renderShapePrimitive(element) {
  const renderGeometry = resolveRenderableGeometry(element);
  const renderElement = renderGeometry ? { ...element, geometry: renderGeometry } : element;
  if (renderGeometry?.kind === "cust") {
    const custom = renderCustomGeometryPrimitive(renderElement);
    if (custom) {
      return custom;
    }
  }

  const shapeType = String(renderElement.shapeType || "rect").toLowerCase();
  if (isLineLikeShapeType(shapeType)) {
    return renderLinePrimitive(renderElement);
  }

  const fillInfo = shapeGradientFill(renderElement);
  const common = `fill=\"${fillInfo.fill}\" ${shapeStrokeAttrs(renderElement.line)}`;
  const transform = transformAttr(renderElement);
  const presetParts = buildPresetShapeParts(shapeType, {
    x: renderElement.x,
    y: renderElement.y,
    cx: renderElement.cx,
    cy: renderElement.cy
  }, renderElement);
  if (presetParts?.length) {
    return `${fillInfo.defs}${presetParts.map((part) => renderPresetShapePart(part, common, transform)).join("")}`;
  }

  switch (shapeType) {
    case "ellipse": {
      const cx = renderElement.x + renderElement.cx / 2;
      const cy = renderElement.y + renderElement.cy / 2;
      const rx = renderElement.cx / 2;
      const ry = renderElement.cy / 2;
      return `${fillInfo.defs}<ellipse cx=\"${cx}\" cy=\"${cy}\" rx=\"${rx}\" ry=\"${ry}\" ${common}${transformAttr(renderElement)} />`;
    }
    case "roundrect": {
      const rx = Math.min(renderElement.cx, renderElement.cy) * 0.08;
      return `${fillInfo.defs}<rect x=\"${renderElement.x}\" y=\"${renderElement.y}\" width=\"${renderElement.cx}\" height=\"${renderElement.cy}\" rx=\"${rx}\" ry=\"${rx}\" ${common}${transformAttr(renderElement)} />`;
    }
    default:
      return `${fillInfo.defs}<rect x=\"${renderElement.x}\" y=\"${renderElement.y}\" width=\"${renderElement.cx}\" height=\"${renderElement.cy}\" ${common}${transformAttr(renderElement)} />`;
  }
}

function renderImage(element) {
  if (!element.dataUri) {
    return "";
  }
  return `<image x=\"${element.x}\" y=\"${element.y}\" width=\"${element.cx}\" height=\"${element.cy}\" href=\"${escapeXml(element.dataUri)}\" preserveAspectRatio=\"none\"${transformAttr(element)} />`;
}

function svgFontFamily(style = {}) {
  const families = [
    style.fontFamily || "Calibri",
    style.eastAsiaFont || "",
    "Yu Gothic UI",
    "Meiryo",
    "MS Gothic",
    "sans-serif"
  ].filter(Boolean);
  return families.map((name) => (/[,\s]/.test(name) ? `'${name.replace(/'/g, "\\'")}'` : name)).join(", ");
}

function svgLetterSpacing(style = {}) {
  const spacing = Number(style.spacing) || 0;
  if (!spacing || !style.fontSizePt) {
    return "";
  }
  const value = style.fontSizePt * EMU_PER_PT * (spacing / 1000);
  if (Math.abs(value) < 1) {
    return "";
  }
  return ` letter-spacing=\"${value}\"`;
}

function renderTextElement(element) {
  if (!element.text) {
    return "";
  }

  const text = element.text;
  const left = element.x + (text.leftInset || 0);
  const top = element.y + (text.topInset || 0);
  const right = element.x + element.cx - (text.rightInset || 0);
  const width = Math.max(0, right - left);

  if (String(text.direction || "horz").toLowerCase() === "eavert") {
    const paragraph = ensureArray(text.paragraphs)[0];
    const firstRun = paragraph?.runs?.[0];
    if (!firstRun) {
      return "";
    }
    const style = firstRun.style || {};
    const fontSizePt = style.fontSizePt || 18;
    const centerX = left + width / 2;
    const y = top;
    return `<text x=\"${centerX}\" y=\"${y}\" text-anchor=\"middle\" dominant-baseline=\"hanging\" writing-mode=\"vertical-rl\" text-orientation=\"upright\" font-family=\"${escapeXml(svgFontFamily(style))}\" font-size=\"${fontSizePt * EMU_PER_PT}\" font-weight=\"${style.bold ? "700" : "400"}\" font-style=\"${style.italic ? "italic" : "normal"}\" fill=\"${colorWithAlpha(style.color || "#000000", style.alpha ?? 1)}\"${svgLetterSpacing(style)}>${escapeXml(ensureArray(paragraph.runs).map((r) => r.text || "").join("") || " ")}</text>`;
  }

  let cursorY = top;
  const lines = [];

  for (const paragraph of ensureArray(text.paragraphs)) {
    const paragraphText = ensureArray(paragraph.runs).map((r) => r.text || "").join("");
    const firstRun = paragraph.runs?.[0];
    if (!firstRun) {
      continue;
    }

    const style = firstRun.style || {};
    const fontSizePt = style.fontSizePt || 18;
    const lineAdvance = Math.max(fontSizePt * EMU_PER_PT * 1.2, 18000);
    cursorY += lineAdvance;

    const anchor = alignmentToTextAnchor(paragraph.alignment);
    let x = left;
    if (anchor === "middle") {
      x = left + width / 2;
    } else if (anchor === "end") {
      x = right;
    }

    lines.push(
      `<text x=\"${x}\" y=\"${cursorY}\" text-anchor=\"${anchor}\" font-family=\"${escapeXml(svgFontFamily(style))}\" font-size=\"${fontSizePt * EMU_PER_PT}\" font-weight=\"${style.bold ? "700" : "400"}\" font-style=\"${style.italic ? "italic" : "normal"}\" fill=\"${colorWithAlpha(style.color || "#000000", style.alpha ?? 1)}\"${svgLetterSpacing(style)}>${escapeXml(paragraphText || " ")}</text>`
    );
  }

  return lines.join("");
}

function cellTextToSvg(cell, x, y, width, height) {
  const firstRun = cell?.text?.paragraphs?.[0]?.runs?.[0];
  if (!firstRun) {
    return "";
  }
  const style = firstRun.style || {};
  const fontSizePt = style.fontSizePt || 11;
  const tx = x + 2 * EMU_PER_PT;
  const ty = y + Math.min(height - 2 * EMU_PER_PT, 4 * EMU_PER_PT + fontSizePt * EMU_PER_PT);
  return `<text x=\"${tx}\" y=\"${ty}\" font-family=\"${escapeXml(svgFontFamily(style))}\" font-size=\"${fontSizePt * EMU_PER_PT}\" fill=\"${colorWithAlpha(style.color || "#000000", style.alpha ?? 1)}\"${svgLetterSpacing(style)}>${escapeXml(firstRun.text || "")}</text>`;
}

function renderTable(element) {
  const pieces = [];
  const totalWidth = element.gridCols.reduce((sum, width) => sum + width, 0) || element.cx;
  const widthScale = totalWidth ? element.cx / totalWidth : 1;

  let y = element.y;
  for (const row of element.rows) {
    const rowHeight = row.height || (element.rows.length ? element.cy / element.rows.length : element.cy);
    let x = element.x;
    for (let ci = 0; ci < row.cells.length; ci += 1) {
      const cell = row.cells[ci];
      const rawWidth = element.gridCols[ci] || (totalWidth / Math.max(1, row.cells.length));
      const cellWidth = rawWidth * widthScale;
      const cellHeight = rowHeight;
      const fill = colorWithAlpha(cell.fill?.color || "#FFFFFF", cell.fill?.alpha ?? 1);
      const borderColor = colorWithAlpha(cell.borders?.left?.color || "#808080", cell.borders?.left?.alpha ?? 1);
      const borderWidth = Math.max(1, Number(cell.borders?.left?.width || 12700));

      pieces.push(`<rect x=\"${x}\" y=\"${y}\" width=\"${cellWidth}\" height=\"${cellHeight}\" fill=\"${fill}\" stroke=\"${borderColor}\" stroke-width=\"${borderWidth}\" />`);
      pieces.push(cellTextToSvg(cell, x, y, cellWidth, cellHeight));
      x += cellWidth;
    }
    y += rowHeight;
  }

  return pieces.join("");
}

function renderElement(element) {
  if (element.hidden) {
    return "";
  }

  switch (element.type) {
    case "image":
      return renderImage(element);
    case "table":
      return renderTable(element);
    case "line":
      return renderLinePrimitive(element);
    case "shape":
      return renderShapePrimitive(element);
    case "text":
      return `${renderShapePrimitive(element)}${renderTextElement(element)}`;
    default:
      return "";
  }
}

function renderBackground(slide, slideCx, slideCy) {
  const background = slide?.background || {};
  let defs = "";
  let fill = colorWithAlpha(background?.color || "#FFFFFF", background?.alpha ?? 1);
  if (background?.type === "gradient" && ensureArray(background.stops).length) {
    const gradientId = `bgGrad-${slide?.id || "slide"}`;
    if (background.gradientType === "path" && background.path === "circle") {
      defs = `<defs><radialGradient id=\"${gradientId}\" cx=\"50%\" cy=\"50%\" r=\"75%\">${ensureArray(background.stops).map((stop) => `<stop offset=\"${clamp((Number(stop?.pos) || 0) / 1000, 0, 100)}%\" stop-color=\"${colorWithAlpha(stop?.color || "#FFFFFF", stop?.alpha ?? 1)}\" />`).join("")}</radialGradient></defs>`;
    } else {
      const angle = Number(background.angle) || 0;
      const rad = (angle * Math.PI) / 180;
      const x1 = 50 - Math.cos(rad) * 50;
      const y1 = 50 - Math.sin(rad) * 50;
      const x2 = 50 + Math.cos(rad) * 50;
      const y2 = 50 + Math.sin(rad) * 50;
      defs = `<defs><linearGradient id=\"${gradientId}\" x1=\"${x1}%\" y1=\"${y1}%\" x2=\"${x2}%\" y2=\"${y2}%\">${ensureArray(background.stops).map((stop) => `<stop offset=\"${clamp((Number(stop?.pos) || 0) / 1000, 0, 100)}%\" stop-color=\"${colorWithAlpha(stop?.color || "#FFFFFF", stop?.alpha ?? 1)}\" />`).join("")}</linearGradient></defs>`;
    }
    fill = `url(#${gradientId})`;
  }
  const baseRect = `${defs}<rect x=\"0\" y=\"0\" width=\"${slideCx}\" height=\"${slideCy}\" fill=\"${fill}\"/>`;
  if (background?.type === "image" && background?.dataUri) {
    const image = `<image x=\"0\" y=\"0\" width=\"${slideCx}\" height=\"${slideCy}\" href=\"${escapeXml(background.dataUri)}\" preserveAspectRatio=\"none\" />`;
    return `${baseRect}${image}`;
  }
  return baseRect;
}

export function renderSlideToSvg(slide, options = {}) {
  const size = options.slideSizeEmu || { cx: 9144000, cy: 6858000 };
  const slideCx = slide?.cx || size.cx;
  const slideCy = slide?.cy || size.cy;
  const pixelWidth = options.widthPx || emuToPx(slideCx);
  const pixelHeight = options.heightPx || emuToPx(slideCy);
  const background = renderBackground(slide, slideCx, slideCy);

  const renderElements = slide?.renderElements || slide?.elements || [];
  const content = renderElements.map((element) => renderElement(element)).join("\n");
  const svg = `<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" width=\"${pixelWidth}\" height=\"${pixelHeight}\" viewBox=\"0 0 ${slideCx} ${slideCy}\">${background}${content}</svg>`;

  if (options.target && typeof options.target === "object") {
    if ("innerHTML" in options.target) {
      options.target.innerHTML = svg;
    }
  }

  return svg;
}
