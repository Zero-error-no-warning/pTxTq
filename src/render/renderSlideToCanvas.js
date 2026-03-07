import { emuToPx } from "../utils/units.js";
import { clamp, ensureArray } from "../utils/object.js";
import {
  buildGeometryVars,
  evalGeomFormula,
  resolveOoxmlArcFromCurrentPoint,
  splitArcSweep
} from "../utils/geometry.js";

const PX_PER_PT = 96 / 72;

function toCanvasColor(color, alpha = 1) {
  if (!color) {
    return "rgba(0,0,0,0)";
  }
  const normalized = color.replace(/^#/, "");
  const r = Number.parseInt(normalized.slice(0, 2), 16);
  const g = Number.parseInt(normalized.slice(2, 4), 16);
  const b = Number.parseInt(normalized.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${clamp(alpha, 0, 1)})`;
}

function toPxElement(element, scaleX, scaleY) {
  return {
    x: (element.x || 0) * scaleX,
    y: (element.y || 0) * scaleY,
    cx: (element.cx || 0) * scaleX,
    cy: (element.cy || 0) * scaleY,
    rotation: element.rotation || 0,
    flipH: element.flipH || false,
    flipV: element.flipV || false
  };
}

function withElementTransform(ctx, elementPx, drawFn) {
  ctx.save();
  if (elementPx.rotation || elementPx.flipH || elementPx.flipV) {
    const cx = elementPx.x + elementPx.cx / 2;
    const cy = elementPx.y + elementPx.cy / 2;
    ctx.translate(cx, cy);
    if (elementPx.rotation) {
      ctx.rotate((elementPx.rotation * Math.PI) / 180);
    }
    ctx.scale(elementPx.flipH ? -1 : 1, elementPx.flipV ? -1 : 1);
    ctx.translate(-cx, -cy);
  }
  drawFn();
  ctx.restore();
}

function lineWidthToPx(widthEmu, scaleX, scaleY) {
  return Math.max(0, (widthEmu || 0) * ((scaleX + scaleY) / 2));
}

function isLineLikeShapeType(shapeType) {
  const normalized = String(shapeType || "").toLowerCase();
  return normalized === "line" || normalized.includes("connector");
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

function lineCapToCanvas(cap) {
  switch (String(cap || "flat").toLowerCase()) {
    case "rnd":
      return "round";
    case "sq":
      return "square";
    default:
      return "butt";
  }
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

function toCustomDashPx(value, widthPx) {
  const numeric = Number.parseInt(String(value ?? 0), 10);
  if (!Number.isFinite(numeric) || numeric <= 0) {
    return Math.max(1, widthPx);
  }
  return Math.max(1, (numeric / 100000) * widthPx);
}

function lineDashPattern(line, widthPx) {
  if (!line || line.dash === "none") {
    return [];
  }

  const effectiveWidth = Math.max(1, widthPx);
  if (String(line.dash || "").toLowerCase() === "cust") {
    const pattern = [];
    for (const stop of ensureArray(line.customDash)) {
      pattern.push(toCustomDashPx(stop?.d, effectiveWidth));
      pattern.push(toCustomDashPx(stop?.sp, effectiveWidth));
    }
    return pattern;
  }

  const factors = DASH_PRESET_FACTORS[String(line.dash || "solid").toLowerCase()];
  if (!factors) {
    return [];
  }
  return factors.map((factor) => Math.max(1, factor * effectiveWidth));
}

function lineRenderStyle(line, scaleX, scaleY) {
  const widthPx = lineWidthToPx(line?.width, scaleX, scaleY);
  if (!line?.color || widthPx <= 0.05 || line?.dash === "none") {
    return null;
  }
  const effectiveWidth = Math.max(0.5, widthPx);
  return {
    strokeStyle: toCanvasColor(line.color, line.alpha ?? 1),
    widthPx: effectiveWidth,
    dashPattern: lineDashPattern(line, effectiveWidth),
    cap: lineCapToCanvas(line.cap)
  };
}

function applyLineStyle(ctx, style) {
  if (!style) {
    return;
  }
  ctx.strokeStyle = style.strokeStyle;
  ctx.lineWidth = style.widthPx;
  ctx.lineCap = style.cap;
  ctx.lineJoin = "miter";
  ctx.setLineDash(style.dashPattern);
}

function resetLineStyle(ctx) {
  ctx.setLineDash([]);
  ctx.lineCap = "butt";
}

function strokeWithElementLine(ctx, line, scaleX, scaleY) {
  const style = lineRenderStyle(line, scaleX, scaleY);
  if (!style) {
    return null;
  }
  applyLineStyle(ctx, style);
  ctx.stroke();
  resetLineStyle(ctx);
  return style;
}

function polygonPath(ctx, points) {
  if (!points.length) {
    return;
  }
  ctx.moveTo(points[0].x, points[0].y);
  for (let i = 1; i < points.length; i += 1) {
    ctx.lineTo(points[i].x, points[i].y);
  }
  ctx.closePath();
}

function drawLineEnd(ctx, lineEnd, tip, other, strokeStyle, lineWidthPx) {
  const endType = String(lineEnd?.type || "none").toLowerCase();
  if (!lineEnd || endType === "none") {
    return;
  }

  const dx = other.x - tip.x;
  const dy = other.y - tip.y;
  const length = Math.hypot(dx, dy);
  if (length <= 0.01) {
    return;
  }

  const ux = dx / length;
  const uy = dy / length;
  const px = -uy;
  const py = ux;

  const markerLength = Math.max(
    lineWidthPx + 2,
    lineWidthPx * lineEndScale(lineEnd.length, 2.2, 3.2, 4.6)
  );
  const markerHalfWidth = Math.max(
    lineWidthPx * 0.6,
    (lineWidthPx * lineEndScale(lineEnd.width, 1.6, 2.5, 3.3)) / 2
  );

  const backX = tip.x + ux * markerLength;
  const backY = tip.y + uy * markerLength;
  const leftX = backX + px * markerHalfWidth;
  const leftY = backY + py * markerHalfWidth;
  const rightX = backX - px * markerHalfWidth;
  const rightY = backY - py * markerHalfWidth;

  ctx.save();
  ctx.setLineDash([]);
  ctx.strokeStyle = strokeStyle;
  ctx.fillStyle = strokeStyle;

  switch (endType) {
    case "arrow": {
      ctx.lineWidth = Math.max(1, lineWidthPx * 0.9);
      ctx.lineJoin = "round";
      ctx.lineCap = "round";
      ctx.beginPath();
      ctx.moveTo(leftX, leftY);
      ctx.lineTo(tip.x, tip.y);
      ctx.lineTo(rightX, rightY);
      ctx.stroke();
      break;
    }
    case "stealth": {
      const notchX = tip.x + ux * markerLength * 0.38;
      const notchY = tip.y + uy * markerLength * 0.38;
      ctx.beginPath();
      polygonPath(ctx, [
        { x: tip.x, y: tip.y },
        { x: leftX, y: leftY },
        { x: notchX, y: notchY },
        { x: rightX, y: rightY }
      ]);
      ctx.fill();
      break;
    }
    case "diamond": {
      const midX = tip.x + ux * markerLength * 0.5;
      const midY = tip.y + uy * markerLength * 0.5;
      ctx.beginPath();
      polygonPath(ctx, [
        { x: tip.x, y: tip.y },
        { x: midX + px * markerHalfWidth, y: midY + py * markerHalfWidth },
        { x: backX, y: backY },
        { x: midX - px * markerHalfWidth, y: midY - py * markerHalfWidth }
      ]);
      ctx.fill();
      break;
    }
    case "oval": {
      const cx = tip.x + ux * markerLength * 0.55;
      const cy = tip.y + uy * markerLength * 0.55;
      const rx = Math.max(1, markerHalfWidth * 0.95);
      const ry = Math.max(1, markerHalfWidth * 0.75);
      const angle = Math.atan2(uy, ux);
      ctx.beginPath();
      ctx.ellipse(cx, cy, rx, ry, angle, 0, Math.PI * 2);
      ctx.fill();
      break;
    }
    default: {
      ctx.beginPath();
      polygonPath(ctx, [
        { x: tip.x, y: tip.y },
        { x: leftX, y: leftY },
        { x: rightX, y: rightY }
      ]);
      ctx.fill();
      break;
    }
  }
  ctx.restore();
}

function drawLineEnds(ctx, line, start, end, strokeStyle) {
  if (!line || !strokeStyle) {
    return;
  }
  drawLineEnd(ctx, line.headEnd, start, end, strokeStyle.strokeStyle, strokeStyle.widthPx);
  drawLineEnd(ctx, line.tailEnd, end, start, strokeStyle.strokeStyle, strokeStyle.widthPx);
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

function lineEndpoints(box) {
  const start = {
    x: box.flipH ? box.x + box.cx : box.x,
    y: box.flipV ? box.y + box.cy : box.y
  };
  const end = {
    x: box.flipH ? box.x : box.x + box.cx,
    y: box.flipV ? box.y : box.y + box.cy
  };
  if (!box.rotation) {
    return { start, end };
  }
  const center = {
    x: box.x + box.cx / 2,
    y: box.y + box.cy / 2
  };
  return {
    start: rotatePoint(start, center, box.rotation),
    end: rotatePoint(end, center, box.rotation)
  };
}

function drawLineSegment(ctx, box, line, scaleX, scaleY) {
  const { start, end } = lineEndpoints(box);
  ctx.beginPath();
  ctx.moveTo(start.x, start.y);
  ctx.lineTo(end.x, end.y);
  const strokeStyle = strokeWithElementLine(ctx, line, scaleX, scaleY);
  drawLineEnds(ctx, line, start, end, strokeStyle);
}

function pathStrokeEnabled(path) {
  const value = String(path?.stroke ?? "1").toLowerCase();
  return value !== "0" && value !== "false" && value !== "off";
}

function pathFillEnabled(path) {
  const value = String(path?.fill ?? "norm").toLowerCase();
  return value !== "none";
}

function mapGeometryPoint(box, pathW, pathH, x, y) {
  const sx = pathW ? x / pathW : 0;
  const sy = pathH ? y / pathH : 0;
  return {
    x: box.x + sx * box.cx,
    y: box.y + sy * box.cy
  };
}

function traceCustomGeometryPath(ctx, geometry, path, box) {
  const pathW = Math.max(1, Number(path?.w || geometry?.pathDefaults?.w || 21600));
  const pathH = Math.max(1, Number(path?.h || geometry?.pathDefaults?.h || 21600));
  const vars = buildGeometryVars(geometry, pathW, pathH);

  let traced = false;
  let currentRaw = null;
  let subPathStartRaw = null;
  for (const cmd of ensureArray(path?.commands)) {
    const type = String(cmd?.type || "");
    if (type === "moveTo" || type === "lnTo") {
      const rawX = evalGeomFormula(cmd?.x || "0", vars);
      const rawY = evalGeomFormula(cmd?.y || "0", vars);
      const p = mapGeometryPoint(box, pathW, pathH, rawX, rawY);
      if (type === "moveTo") {
        ctx.moveTo(p.x, p.y);
      } else {
        ctx.lineTo(p.x, p.y);
      }
      traced = true;
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
          box,
          pathW,
          pathH,
          evalGeomFormula(pts[0]?.x || "0", vars),
          evalGeomFormula(pts[0]?.y || "0", vars)
        );
        const ep = mapGeometryPoint(
          box,
          pathW,
          pathH,
          evalGeomFormula(pts[1]?.x || "0", vars),
          evalGeomFormula(pts[1]?.y || "0", vars)
        );
        ctx.quadraticCurveTo(cp.x, cp.y, ep.x, ep.y);
        traced = true;
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
          box,
          pathW,
          pathH,
          evalGeomFormula(pts[0]?.x || "0", vars),
          evalGeomFormula(pts[0]?.y || "0", vars)
        );
        const cp2 = mapGeometryPoint(
          box,
          pathW,
          pathH,
          evalGeomFormula(pts[1]?.x || "0", vars),
          evalGeomFormula(pts[1]?.y || "0", vars)
        );
        const ep = mapGeometryPoint(
          box,
          pathW,
          pathH,
          evalGeomFormula(pts[2]?.x || "0", vars),
          evalGeomFormula(pts[2]?.y || "0", vars)
        );
        ctx.bezierCurveTo(cp1.x, cp1.y, cp2.x, cp2.y, ep.x, ep.y);
        traced = true;
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
        const center = mapGeometryPoint(box, pathW, pathH, arc.cx, arc.cy);
        const rx = Math.abs((arc.rx / pathW) * box.cx);
        const ry = Math.abs((arc.ry / pathH) * box.cy);
        const chunks = splitArcSweep(arc.sweepParam);

        let segStart = arc.startParam;
        for (const sweepChunk of chunks) {
          const segEnd = segStart + sweepChunk;
          ctx.ellipse(center.x, center.y, rx, ry, 0, segStart, segEnd, sweepChunk < 0);
          currentRaw = {
            x: arc.cx + arc.rx * Math.cos(segEnd),
            y: arc.cy + arc.ry * Math.sin(segEnd)
          };
          segStart = segEnd;
          traced = true;
        }
      }
      continue;
    }

    if (type === "close") {
      ctx.closePath();
      traced = true;
      if (subPathStartRaw) {
        currentRaw = { ...subPathStartRaw };
      }
    }
  }

  return traced;
}

function drawCustomGeometry(ctx, element, box, scaleX, scaleY) {
  const geometry = element?.geometry;
  const paths = ensureArray(geometry?.paths);
  if (!paths.length) {
    return false;
  }

  let drew = false;
  for (const path of paths) {
    ctx.beginPath();
    const traced = traceCustomGeometryPath(ctx, geometry, path, box);
    if (!traced) {
      continue;
    }

    if (element.fill?.color && pathFillEnabled(path)) {
      ctx.fillStyle = toCanvasColor(element.fill.color, element.fill.alpha ?? 1);
      ctx.fill();
      drew = true;
    }

    if (pathStrokeEnabled(path)) {
      const strokeStyle = strokeWithElementLine(ctx, element.line, scaleX, scaleY);
      drew = drew || Boolean(strokeStyle);
    }
  }

  return drew;
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

function trianglePoints(box, element) {
  const adj = presetAdjustValue(element, "adj", 50000) / 100000;
  return [
    { x: box.x + box.cx * adj, y: box.y },
    { x: box.x + box.cx, y: box.y + box.cy },
    { x: box.x, y: box.y + box.cy }
  ];
}

function setPathForShape(ctx, shapeType, box, element = null) {
  const normalized = String(shapeType || "rect").toLowerCase();
  switch (normalized) {
    case "ellipse": {
      const cx = box.x + box.cx / 2;
      const cy = box.y + box.cy / 2;
      ctx.ellipse(cx, cy, box.cx / 2, box.cy / 2, 0, 0, Math.PI * 2);
      break;
    }
    case "roundrect": {
      const radius = Math.min(box.cx, box.cy) * 0.08;
      const x = box.x;
      const y = box.y;
      const w = box.cx;
      const h = box.cy;
      ctx.moveTo(x + radius, y);
      ctx.lineTo(x + w - radius, y);
      ctx.quadraticCurveTo(x + w, y, x + w, y + radius);
      ctx.lineTo(x + w, y + h - radius);
      ctx.quadraticCurveTo(x + w, y + h, x + w - radius, y + h);
      ctx.lineTo(x + radius, y + h);
      ctx.quadraticCurveTo(x, y + h, x, y + h - radius);
      ctx.lineTo(x, y + radius);
      ctx.quadraticCurveTo(x, y, x + radius, y);
      ctx.closePath();
      break;
    }
    case "triangle": {
      polygonPath(ctx, trianglePoints(box, element));
      break;
    }
    case "rttriangle": {
      polygonPath(ctx, [
        { x: box.x, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy }
      ]);
      break;
    }
    case "diamond": {
      polygonPath(ctx, [
        { x: box.x + box.cx / 2, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy / 2 },
        { x: box.x + box.cx / 2, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy / 2 }
      ]);
      break;
    }
    case "parallelogram": {
      const offset = box.cx * 0.2;
      polygonPath(ctx, [
        { x: box.x + offset, y: box.y },
        { x: box.x + box.cx, y: box.y },
        { x: box.x + box.cx - offset, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy }
      ]);
      break;
    }
    case "trapezoid": {
      const inset = box.cx * 0.18;
      polygonPath(ctx, [
        { x: box.x + inset, y: box.y },
        { x: box.x + box.cx - inset, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy }
      ]);
      break;
    }
    case "pentagon": {
      polygonPath(ctx, [
        { x: box.x + box.cx * 0.5, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy * 0.38 },
        { x: box.x + box.cx * 0.82, y: box.y + box.cy },
        { x: box.x + box.cx * 0.18, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy * 0.38 }
      ]);
      break;
    }
    case "hexagon": {
      polygonPath(ctx, [
        { x: box.x + box.cx * 0.25, y: box.y },
        { x: box.x + box.cx * 0.75, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy * 0.5 },
        { x: box.x + box.cx * 0.75, y: box.y + box.cy },
        { x: box.x + box.cx * 0.25, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy * 0.5 }
      ]);
      break;
    }
    case "chevron": {
      polygonPath(ctx, [
        { x: box.x, y: box.y },
        { x: box.x + box.cx * 0.6, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy * 0.5 },
        { x: box.x + box.cx * 0.6, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy },
        { x: box.x + box.cx * 0.4, y: box.y + box.cy * 0.5 }
      ]);
      break;
    }
    case "line":
    case "straightconnector1":
    case "bentconnector2":
    case "bentconnector3":
    case "bentconnector4":
    case "bentconnector5":
    case "curvedconnector2":
    case "curvedconnector3":
    case "curvedconnector4":
    case "curvedconnector5": {
      ctx.moveTo(box.x, box.y);
      ctx.lineTo(box.x + box.cx, box.y + box.cy);
      break;
    }
    default:
      ctx.rect(box.x, box.y, box.cx, box.cy);
  }
}

function drawShape(ctx, element, scaleX, scaleY) {
  const box = toPxElement(element, scaleX, scaleY);
  if (isLineLikeShapeType(element.shapeType)) {
    if (element?.geometry?.kind === "cust") {
      withElementTransform(ctx, box, () => {
        drawLineSegment(ctx, box, element.line, scaleX, scaleY);
      });
      return;
    }
    drawLineSegment(ctx, box, element.line, scaleX, scaleY);
    return;
  }

  withElementTransform(ctx, box, () => {
    if (element?.geometry?.kind === "cust") {
      const drawn = drawCustomGeometry(ctx, element, box, scaleX, scaleY);
      if (drawn) {
        return;
      }
    }

    ctx.beginPath();
    setPathForShape(ctx, element.shapeType, box, element);

    if (element.fill?.color) {
      ctx.fillStyle = toCanvasColor(element.fill.color, element.fill.alpha ?? 1);
      ctx.fill();
    }

    strokeWithElementLine(ctx, element.line, scaleX, scaleY);
  });
}

function drawLineElement(ctx, element, scaleX, scaleY) {
  const box = toPxElement(element, scaleX, scaleY);
  if (element?.geometry?.kind === "cust") {
    withElementTransform(ctx, box, () => {
      const drawn = drawCustomGeometry(ctx, element, box, scaleX, scaleY);
      if (drawn) {
        return;
      }
      drawLineSegment(ctx, box, element.line, scaleX, scaleY);
    });
    return;
  }
  drawLineSegment(ctx, box, element.line, scaleX, scaleY);
}

function normalizeRunStyle(style = {}, defaultStyle = {}) {
  return {
    fontFamily: style.fontFamily || defaultStyle.fontFamily || "Calibri",
    fontSizePt: style.fontSizePt || defaultStyle.fontSizePt || 18,
    bold: style.bold || false,
    italic: style.italic || false,
    color: style.color || defaultStyle.color || "#000000",
    alpha: style.alpha ?? defaultStyle.alpha ?? 1,
    underline: style.underline || false,
    spacing: style.spacing ?? defaultStyle.spacing ?? 0
  };
}

function styleKey(style) {
  return `${style.fontFamily}|${style.fontSizePt}|${style.bold ? 1 : 0}|${style.italic ? 1 : 0}|${style.color}|${style.alpha}|${style.underline ? 1 : 0}|${style.spacing ?? 0}`;
}

function charSpacingPx(style) {
  const spacing = Number(style?.spacing) || 0;
  const fontPx = (Number(style?.fontSizePt) || 0) * PX_PER_PT;
  return (fontPx * spacing) / 1000;
}

const TEXT_METRICS_CACHE = new Map();

function measureStyleMetrics(ctx, style) {
  const key = styleKey(style);
  if (TEXT_METRICS_CACHE.has(key)) {
    return TEXT_METRICS_CACHE.get(key);
  }

  setCanvasFont(ctx, style);
  const sample = ctx.measureText("Hg明");
  const fontPx = Math.max(1, (Number(style.fontSizePt) || 0) * PX_PER_PT);
  const metrics = {
    fontPx,
    ascent: Math.max(sample.actualBoundingBoxAscent || 0, fontPx * 0.72),
    descent: Math.max(sample.actualBoundingBoxDescent || 0, fontPx * 0.08)
  };
  metrics.height = Math.max(metrics.ascent + metrics.descent, fontPx);
  TEXT_METRICS_CACHE.set(key, metrics);
  return metrics;
}

function quoteFontFamily(name) {
  if (!name) {
    return "";
  }
  if (/^[a-zA-Z0-9_-]+$/.test(name)) {
    return name;
  }
  return `"${String(name).replace(/"/g, '\\"')}"`;
}

function setCanvasFont(ctx, style) {
  const italic = style.italic ? "italic" : "normal";
  const weight = style.bold ? "bold" : "normal";
  const sizePx = Math.max(1, style.fontSizePt * PX_PER_PT);
  const families = [
    quoteFontFamily(style.fontFamily || "Calibri"),
    quoteFontFamily(style.eastAsiaFont || ""),
    "\"Yu Gothic UI\"",
    "Meiryo",
    "\"MS Gothic\"",
    "sans-serif"
  ].filter(Boolean).join(", ");
  ctx.font = `${italic} ${weight} ${sizePx}px ${families}`;
  ctx.fillStyle = toCanvasColor(style.color, style.alpha);
  return sizePx;
}

function pushTextSegment(line, text, style, width) {
  if (!text) {
    return;
  }
  const key = styleKey(style);
  const last = line.segments[line.segments.length - 1];
  if ((style.spacing ?? 0) === 0 && last && last.key === key) {
    last.text += text;
    last.width += width;
  } else {
    line.segments.push({ key, text, style, width });
  }
  line.width += width;
}

function newLine(paragraph) {
  return {
    segments: [],
    width: 0,
    maxFontSizePx: 0,
    maxAscent: 0,
    maxDescent: 0,
    alignment: paragraph?.alignment || "l",
    lineSpacing: paragraph?.lineSpacing || 0
  };
}

function flushLine(lines, line) {
  if (line.segments.length) {
    lines.push(line);
  }
}

function lineDisplayWidth(line) {
  if (!line?.segments?.length) {
    return 0;
  }
  const last = line.segments[line.segments.length - 1];
  return Math.max(0, line.width - charSpacingPx(last.style));
}

function wrapParagraphRuns(ctx, paragraph, boxWidthPx, defaultStyle) {
  const lines = [];
  let line = newLine(paragraph);

  for (const run of ensureArray(paragraph?.runs)) {
    const style = normalizeRunStyle(run.style || {}, defaultStyle);
    const metrics = measureStyleMetrics(ctx, style);
    const fontPx = metrics.fontPx;
    line.maxFontSizePx = Math.max(line.maxFontSizePx, fontPx);
    line.maxAscent = Math.max(line.maxAscent, metrics.ascent);
    line.maxDescent = Math.max(line.maxDescent, metrics.descent);
    setCanvasFont(ctx, style);

    const text = String(run.text || "");
    for (const ch of Array.from(text)) {
      if (ch === "\n") {
        flushLine(lines, line);
        line = newLine(paragraph);
        line.maxFontSizePx = Math.max(line.maxFontSizePx, fontPx);
        line.maxAscent = Math.max(line.maxAscent, metrics.ascent);
        line.maxDescent = Math.max(line.maxDescent, metrics.descent);
        continue;
      }

      const width = ctx.measureText(ch).width + charSpacingPx(style);
      if (line.width + width > boxWidthPx && line.segments.length) {
        flushLine(lines, line);
        line = newLine(paragraph);
        line.maxFontSizePx = Math.max(line.maxFontSizePx, fontPx);
        line.maxAscent = Math.max(line.maxAscent, metrics.ascent);
        line.maxDescent = Math.max(line.maxDescent, metrics.descent);
      }
      pushTextSegment(line, ch, style, width);
    }
  }

  flushLine(lines, line);
  return lines;
}

function lineHeightPx(line) {
  const natural = Math.max(line.maxAscent + line.maxDescent, line.maxFontSizePx || 0);
  const spacing = line.lineSpacing ? Math.max(1, line.lineSpacing / 100000) : 1;
  return natural * spacing;
}

function alignmentToStartX(left, width, lineWidth, alignment) {
  if (alignment === "ctr") {
    return left + Math.max(0, (width - lineWidth) / 2);
  }
  if (alignment === "r") {
    return left + Math.max(0, width - lineWidth);
  }
  return left;
}

function verticalAnchorStartY(top, height, contentHeight, anchor) {
  if (anchor === "ctr") {
    return top + Math.max(0, (height - contentHeight) / 2);
  }
  if (anchor === "b") {
    return top + Math.max(0, height - contentHeight);
  }
  return top;
}

function normalizeBulletChar(bullet) {
  const char = bullet?.char || "\u2022";
  const font = String(bullet?.fontFamily || "").toLowerCase();
  if (font.includes("wingdings") && /[\uF000-\uF0FF]/.test(char)) {
    return "\u25B6";
  }
  return char;
}

function buildVerticalGlyphs(ctx, textBody, defaultStyle) {
  const glyphs = [];
  for (const paragraph of ensureArray(textBody.paragraphs)) {
    for (const run of ensureArray(paragraph?.runs)) {
      const style = normalizeRunStyle(run.style || {}, defaultStyle);
      setCanvasFont(ctx, style);
      for (const ch of Array.from(String(run.text || ""))) {
        if (ch === "\n") {
          continue;
        }
        const measuredWidth = ctx.measureText(ch).width;
        const metrics = measureStyleMetrics(ctx, style);
        glyphs.push({
          ch,
          style,
          measuredWidth,
          ascent: metrics.ascent,
          descent: metrics.descent,
          height: metrics.height,
          advance: Math.max(metrics.height, measuredWidth) + charSpacingPx(style)
        });
      }
    }
  }
  return glyphs;
}

function drawVerticalTextBodyInBox(ctx, textBody, boxPx, defaultStyle) {
  const left = boxPx.x + (textBody.leftInset || 0);
  const top = boxPx.y + (textBody.topInset || 0);
  const width = Math.max(0, boxPx.cx - (textBody.leftInset || 0) - (textBody.rightInset || 0));
  const height = Math.max(0, boxPx.cy - (textBody.topInset || 0) - (textBody.bottomInset || 0));
  const glyphs = buildVerticalGlyphs(ctx, textBody, defaultStyle);
  if (!glyphs.length) {
    return;
  }

  const totalHeight = glyphs.reduce((sum, glyph, index) => (
    sum + glyph.advance - (index === glyphs.length - 1 ? charSpacingPx(glyph.style) : 0)
  ), 0);

  let centerX = left + width / 2;
  const firstParagraph = ensureArray(textBody.paragraphs)[0] || {};
  const alignment = firstParagraph.alignment || "ctr";
  if (alignment === "l") {
    centerX = left + width * 0.5;
  } else if (alignment === "r") {
    centerX = left + width * 0.5;
  }

  let cursorY = verticalAnchorStartY(top, height, totalHeight, textBody.verticalAlign || "t");

  ctx.save();
  ctx.beginPath();
  ctx.rect(left, top, width, height);
  ctx.clip();
  ctx.textAlign = "center";
  ctx.textBaseline = "alphabetic";

  for (const glyph of glyphs) {
    setCanvasFont(ctx, glyph.style);
    const slack = Math.max(0, glyph.advance - charSpacingPx(glyph.style) - (glyph.ascent + glyph.descent)) / 2;
    const baselineY = cursorY + slack + glyph.ascent;
    ctx.fillText(glyph.ch, centerX, baselineY);
    cursorY += glyph.advance;
  }

  ctx.restore();
}

function drawTextBodyInBox(ctx, textBody, boxPx) {
  if (!textBody) {
    return;
  }

  const left = boxPx.x + (textBody.leftInset || 0);
  const top = boxPx.y + (textBody.topInset || 0);
  const width = Math.max(0, boxPx.cx - (textBody.leftInset || 0) - (textBody.rightInset || 0));
  const height = Math.max(0, boxPx.cy - (textBody.topInset || 0) - (textBody.bottomInset || 0));

  const defaultStyle = {
    fontFamily: "Calibri",
    fontSizePt: 18,
    color: "#000000",
    alpha: 1,
    bold: false,
    italic: false,
    spacing: 0
  };

  if (String(textBody.direction || "horz").toLowerCase() === "eavert") {
    drawVerticalTextBodyInBox(ctx, textBody, boxPx, defaultStyle);
    return;
  }

  const paragraphLayouts = [];
  let totalHeight = 0;

  for (const paragraph of ensureArray(textBody.paragraphs)) {
    const before = Math.max(0, (paragraph.spaceBefore || 0) / 100) * PX_PER_PT;
    const after = Math.max(0, (paragraph.spaceAfter || 0) / 100) * PX_PER_PT;
    const marginLeft = Math.max(0, paragraph.marginLeft || 0);
    const marginRight = Math.max(0, paragraph.marginRight || 0);
    const indent = paragraph.indent || 0;
    const paragraphWidth = Math.max(0, width - marginLeft - marginRight);
    const lines = wrapParagraphRuns(ctx, paragraph, paragraphWidth, defaultStyle);
    if (!lines.length) {
      continue;
    }
    const heights = lines.map((line) => lineHeightPx(line));
    const paraHeight = before + heights.reduce((sum, h) => sum + h, 0) + after;
    paragraphLayouts.push({
      lines,
      heights,
      before,
      after,
      marginLeft,
      marginRight,
      indent,
      bullet: paragraph.bullet || null
    });
    totalHeight += paraHeight;
  }

  if (!paragraphLayouts.length) {
    return;
  }

  ctx.save();
  ctx.beginPath();
  ctx.rect(left, top, width, height);
  ctx.clip();

  let cursorY = verticalAnchorStartY(top, height, totalHeight, textBody.verticalAlign || "t");

  for (const layout of paragraphLayouts) {
    cursorY += layout.before;

    for (let li = 0; li < layout.lines.length; li += 1) {
      const line = layout.lines[li];
      const lh = layout.heights[li];
      const slack = Math.max(0, lh - (line.maxAscent + line.maxDescent)) / 2;
      const baselineY = cursorY + slack + line.maxAscent;
      const paragraphLeft = left + layout.marginLeft;
      const paragraphWidth = Math.max(0, width - layout.marginLeft - layout.marginRight);
      const startX = alignmentToStartX(paragraphLeft, paragraphWidth, lineDisplayWidth(line), line.alignment);

      if (li === 0 && layout.bullet?.type === "char") {
        const baseStyle = line.segments[0]?.style || defaultStyle;
        const bulletStyle = {
          ...baseStyle,
          fontFamily: layout.bullet.fontFamily || baseStyle.fontFamily,
          color: layout.bullet.color || baseStyle.color,
          alpha: layout.bullet.alpha ?? baseStyle.alpha,
          fontSizePt: layout.bullet.sizePt || (baseStyle.fontSizePt * (layout.bullet.sizePct || 1))
        };
        setCanvasFont(ctx, bulletStyle);
        ctx.textBaseline = "alphabetic";
        ctx.fillText(normalizeBulletChar(layout.bullet), paragraphLeft + Math.min(0, layout.indent), baselineY);
      }

      let x = startX;
      for (const seg of line.segments) {
        setCanvasFont(ctx, seg.style);
        ctx.textBaseline = "alphabetic";
        ctx.fillText(seg.text, x, baselineY);

        if (seg.style.underline) {
          const underlineY = baselineY + seg.style.fontSizePt * PX_PER_PT * 0.08;
          ctx.strokeStyle = toCanvasColor(seg.style.color, seg.style.alpha);
          ctx.lineWidth = Math.max(0.5, seg.style.fontSizePt * PX_PER_PT * 0.025);
          ctx.beginPath();
          ctx.moveTo(x, underlineY);
          ctx.lineTo(x + seg.width, underlineY);
          ctx.stroke();
        }

        x += seg.width;
      }

      cursorY += lh;
    }

    cursorY += layout.after;
  }

  ctx.restore();
}

function drawText(ctx, element, scaleX, scaleY) {
  const box = toPxElement(element, scaleX, scaleY);
  const paragraphs = ensureArray(element.text?.paragraphs).map((paragraph) => ({
    ...paragraph,
    marginLeft: (paragraph.marginLeft || 0) * scaleX,
    marginRight: (paragraph.marginRight || 0) * scaleX,
    indent: (paragraph.indent || 0) * scaleX
  }));

  withElementTransform(ctx, box, () => {
    drawTextBodyInBox(ctx, {
      ...element.text,
      paragraphs,
      leftInset: (element.text?.leftInset || 0) * scaleX,
      rightInset: (element.text?.rightInset || 0) * scaleX,
      topInset: (element.text?.topInset || 0) * scaleY,
      bottomInset: (element.text?.bottomInset || 0) * scaleY
    }, box);
  });
}

async function loadImage(dataUri) {
  if (!dataUri || typeof Image === "undefined") {
    return null;
  }

  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = dataUri;
  });
}

async function drawImageElement(ctx, element, scaleX, scaleY) {
  if (!element.dataUri) {
    return;
  }
  const image = await loadImage(element.dataUri);
  if (!image) {
    return;
  }

  const box = toPxElement(element, scaleX, scaleY);
  withElementTransform(ctx, box, () => {
    ctx.drawImage(image, box.x, box.y, box.cx, box.cy);
  });
}

async function drawSlideBackground(ctx, slide, widthPx, heightPx) {
  const background = slide?.background || {};
  let fillStyle = toCanvasColor("#FFFFFF", 1);
  if (background?.type === "solid") {
    fillStyle = toCanvasColor(background.color || "#FFFFFF", background.alpha ?? 1);
  } else if (background?.type === "gradient" && ensureArray(background.stops).length) {
    const angleRad = ((Number(background.angle) || 0) * Math.PI) / 180;
    const dx = Math.cos(angleRad);
    const dy = Math.sin(angleRad);
    const cx = widthPx / 2;
    const cy = heightPx / 2;
    const tx = Math.abs(dx) < 0.0001 ? Number.POSITIVE_INFINITY : (widthPx / 2) / Math.abs(dx);
    const ty = Math.abs(dy) < 0.0001 ? Number.POSITIVE_INFINITY : (heightPx / 2) / Math.abs(dy);
    const span = Math.min(tx, ty);
    const x0 = cx - dx * span;
    const y0 = cy - dy * span;
    const x1 = cx + dx * span;
    const y1 = cy + dy * span;
    const gradient = ctx.createLinearGradient(x0, y0, x1, y1);
    for (const stop of ensureArray(background.stops)) {
      gradient.addColorStop(
        clamp((Number(stop?.pos) || 0) / 100000, 0, 1),
        toCanvasColor(stop?.color || "#FFFFFF", stop?.alpha ?? 1)
      );
    }
    fillStyle = gradient;
  }

  ctx.fillStyle = fillStyle;
  ctx.fillRect(0, 0, widthPx, heightPx);

  if (background?.type === "image" && background.dataUri) {
    const image = await loadImage(background.dataUri);
    if (image) {
      ctx.drawImage(image, 0, 0, widthPx, heightPx);
    }
  }
}

function drawTableCellText(ctx, cell, x, y, width, height) {
  if (!cell?.text) {
    return;
  }
  const textBody = {
    ...cell.text,
    leftInset: cell.marginLeft || 2,
    rightInset: cell.marginRight || 2,
    topInset: cell.marginTop || 1,
    bottomInset: cell.marginBottom || 1
  };
  drawTextBodyInBox(ctx, textBody, { x, y, cx: width, cy: height });
}

function drawTable(ctx, element, scaleX, scaleY) {
  const box = toPxElement(element, scaleX, scaleY);

  withElementTransform(ctx, box, () => {
    const totalWidth = element.gridCols.reduce((sum, width) => sum + width, 0) || element.cx;
    const widthScale = totalWidth ? box.cx / totalWidth : 1;

    let y = box.y;
    for (const row of element.rows) {
      const rowHeight = (row.height || (element.rows.length ? element.cy / element.rows.length : element.cy)) * scaleY;
      let x = box.x;

      for (let ci = 0; ci < row.cells.length; ci += 1) {
        const cell = row.cells[ci];
        const rawWidth = element.gridCols[ci] || (totalWidth / Math.max(1, row.cells.length));
        const cellWidth = rawWidth * widthScale;

        ctx.fillStyle = toCanvasColor(cell.fill?.color || "#FFFFFF", cell.fill?.alpha ?? 1);
        ctx.fillRect(x, y, cellWidth, rowHeight);

        const border = cell.borders?.left || cell.borders?.top || cell.borders?.right || cell.borders?.bottom;
        if (border?.color) {
          ctx.strokeStyle = toCanvasColor(border.color, border.alpha ?? 1);
          ctx.lineWidth = Math.max(0.5, lineWidthToPx(border.width || 12700, scaleX, scaleY));
          ctx.strokeRect(x, y, cellWidth, rowHeight);
        }

        drawTableCellText(ctx, cell, x, y, cellWidth, rowHeight);
        x += cellWidth;
      }

      y += rowHeight;
    }
  });
}

function drawChart(ctx, element, scaleX, scaleY) {
  if (!element?.chart || element.chart.chartType !== "pie") {
    return;
  }

  const box = toPxElement(element, scaleX, scaleY);
  const series = ensureArray(element.chart.series)[0];
  if (!series) {
    return;
  }
  const values = ensureArray(series.values).map((v) => Number(v) || 0);
  const total = values.reduce((sum, v) => sum + Math.max(0, v), 0);
  if (total <= 0) {
    return;
  }

  const centerX = box.x + box.cx * 0.5;
  const centerY = box.y + box.cy * 0.58;
  const radius = Math.max(8, Math.min(box.cx, box.cy) * 0.28);
  const titleSize = Math.max(12, box.cy * 0.09);
  const labelSize = Math.max(12, box.cy * 0.08);
  const valueSize = Math.max(10, box.cy * 0.07);

  withElementTransform(ctx, box, () => {
    if (element.chart.title) {
      ctx.fillStyle = "rgba(89, 89, 89, 1)";
      ctx.textAlign = "center";
      ctx.textBaseline = "top";
      ctx.font = `700 ${titleSize}px "Yu Gothic UI", Meiryo, "MS Gothic", sans-serif`;
      ctx.fillText(String(element.chart.title), centerX, box.y + box.cy * 0.03);
    }

    let startAngle = -Math.PI / 2;
    const categories = ensureArray(series.categories);
    const colors = ensureArray(series.colors);

    for (let i = 0; i < values.length; i += 1) {
      const value = Math.max(0, values[i]);
      if (!value) {
        continue;
      }
      const ratio = value / total;
      const angle = ratio * Math.PI * 2;
      const endAngle = startAngle + angle;
      const color = colors[i] || "#92278F";

      ctx.beginPath();
      ctx.moveTo(centerX, centerY);
      ctx.arc(centerX, centerY, radius, startAngle, endAngle, false);
      ctx.closePath();
      ctx.fillStyle = toCanvasColor(color, 1);
      ctx.fill();

      const mid = startAngle + angle / 2;
      const labelX = centerX + Math.cos(mid) * radius * 1.35;
      const labelY = centerY + Math.sin(mid) * radius * 1.35;
      const align = Math.cos(mid) >= 0 ? "left" : "right";

      ctx.textAlign = align;
      ctx.textBaseline = "alphabetic";
      ctx.fillStyle = toCanvasColor(color, 1);
      ctx.font = `700 ${labelSize}px "Yu Gothic UI", Meiryo, "MS Gothic", sans-serif`;
      ctx.fillText(String(categories[i] || ""), labelX, labelY);

      ctx.font = `700 ${valueSize}px "Yu Gothic UI", Meiryo, "MS Gothic", sans-serif`;
      ctx.fillText(`${Math.round(ratio * 100)}%`, labelX, labelY + valueSize * 1.45);

      startAngle = endAngle;
    }
  });
}

async function drawRenderableElement(ctx, element, scaleX, scaleY) {
  if (!element || element.hidden) {
    return;
  }

  if (element.type === "image") {
    await drawImageElement(ctx, element, scaleX, scaleY);
    return;
  }

  if (element.type === "table") {
    drawTable(ctx, element, scaleX, scaleY);
    return;
  }

  if (element.type === "chart") {
    drawChart(ctx, element, scaleX, scaleY);
    return;
  }

  if (element.type === "diagram") {
    for (const child of ensureArray(element.diagramElements)) {
      await drawRenderableElement(ctx, child, scaleX, scaleY);
    }
    return;
  }

  if (element.type === "line") {
    drawLineElement(ctx, element, scaleX, scaleY);
    return;
  }

  if (element.type === "shape" || element.type === "text") {
    drawShape(ctx, element, scaleX, scaleY);
    if (element.type === "text") {
      drawText(ctx, element, scaleX, scaleY);
    }
  }
}

export async function renderSlideToCanvas(slide, canvasOrContext, options = {}) {
  if (!canvasOrContext) {
    throw new Error("canvasOrContext is required");
  }

  const size = options.slideSizeEmu || { cx: 9144000, cy: 6858000 };
  const slideCx = slide?.cx || size.cx;
  const slideCy = slide?.cy || size.cy;

  const canvas = typeof canvasOrContext.getContext === "function" ? canvasOrContext : canvasOrContext.canvas;
  const ctx = typeof canvasOrContext.getContext === "function" ? canvasOrContext.getContext("2d") : canvasOrContext;

  if (!ctx) {
    throw new Error("Canvas 2D context is not available");
  }

  const widthPx = Math.round(options.widthPx || emuToPx(slideCx));
  const heightPx = Math.round(options.heightPx || emuToPx(slideCy));
  const scaleX = widthPx / slideCx;
  const scaleY = heightPx / slideCy;

  canvas.width = widthPx;
  canvas.height = heightPx;

  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, widthPx, heightPx);
  await drawSlideBackground(ctx, slide, widthPx, heightPx);

  const renderElements = slide?.renderElements || slide?.elements || [];
  for (const element of renderElements) {
    await drawRenderableElement(ctx, element, scaleX, scaleY);
  }

  return canvas;
}
