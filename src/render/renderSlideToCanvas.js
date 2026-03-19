import { emuToPx } from "../utils/units.js";
import { clamp, ensureArray } from "../utils/object.js";
import { darken, lighten } from "../utils/color.js";
import {
  buildGeometryVars,
  evalGeomFormula,
  fitGeometryExtents,
  resolveGeometryGuides,
  resolveOoxmlArcFromCurrentPoint,
  splitArcSweep
} from "../utils/geometry.js";
import { shouldAspectFitPresetGeometry } from "../utils/aspectFitPresetGeometry.js";
import { resolvePresetShapeGeometry } from "../utils/presetShapeGeometry.js";
import { buildPresetShapeParts } from "../utils/presetShape.js";

const PX_PER_PT = 96 / 72;
const BASE_EMU_TO_PX = emuToPx(1);

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

function lineJoinType(join) {
  return String(join?.type || join || "miter").toLowerCase();
}

function lineJoinToCanvas(join) {
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
    cap: lineCapToCanvas(line.cap),
    join: lineJoinToCanvas(line.join),
    miterLimit: lineMiterLimit(line.join),
    compound: String(line?.cmpd || "sng").toLowerCase()
  };
}

function applyLineStyle(ctx, style) {
  if (!style) {
    return;
  }
  ctx.strokeStyle = style.strokeStyle;
  ctx.lineWidth = style.widthPx;
  ctx.lineCap = style.cap;
  ctx.lineJoin = style.join;
  ctx.miterLimit = style.miterLimit;
  ctx.setLineDash(style.dashPattern);
}

function resetLineStyle(ctx) {
  ctx.setLineDash([]);
  ctx.lineCap = "butt";
  ctx.lineJoin = "miter";
  ctx.miterLimit = 10;
}

function compoundStrokeSpec(style) {
  switch (style?.compound) {
    case "dbl":
      return { gapRatio: 0.5, innerRatio: 0 };
    case "tri":
      return { gapRatio: 0.58, innerRatio: 0.18 };
    case "thickthin":
      return { gapRatio: 0.56, innerRatio: 0.14 };
    case "thinthick":
      return { gapRatio: 0.56, innerRatio: 0.22 };
    default:
      return null;
  }
}

function strokeCurrentPath(ctx, style) {
  ctx.stroke();

  const compound = compoundStrokeSpec(style);
  if (!compound) {
    return;
  }

  ctx.save();
  ctx.globalCompositeOperation = "destination-out";
  ctx.lineWidth = Math.max(0.5, style.widthPx * compound.gapRatio);
  ctx.stroke();
  ctx.restore();

  if (compound.innerRatio > 0) {
    ctx.lineWidth = Math.max(0.5, style.widthPx * compound.innerRatio);
    ctx.stroke();
  }
}

function strokeWithElementLine(ctx, line, scaleX, scaleY) {
  const style = lineRenderStyle(line, scaleX, scaleY);
  if (!style) {
    return null;
  }
  applyLineStyle(ctx, style);
  strokeCurrentPath(ctx, style);
  resetLineStyle(ctx);
  return style;
}

function gradientVectorForBox(box, angleDeg = 0) {
  const angleRad = ((Number(angleDeg) || 0) * Math.PI) / 180;
  const dx = Math.cos(angleRad);
  const dy = Math.sin(angleRad);
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  const tx = Math.abs(dx) < 0.0001 ? Number.POSITIVE_INFINITY : (box.cx / 2) / Math.abs(dx);
  const ty = Math.abs(dy) < 0.0001 ? Number.POSITIVE_INFINITY : (box.cy / 2) / Math.abs(dy);
  const span = Math.min(tx, ty);
  return {
    x0: cx - dx * span,
    y0: cy - dy * span,
    x1: cx + dx * span,
    y1: cy + dy * span
  };
}

function createCanvasFillStyle(ctx, fill, box) {
  if (!fill || fill.type === "none") {
    return null;
  }
  if (fill.type === "gradient" && ensureArray(fill.stops).length) {
    let gradient = null;
    if (fill.gradientType === "path" && fill.path === "circle") {
      const cx = box.x + box.cx / 2;
      const cy = box.y + box.cy / 2;
      const radius = Math.max(1, Math.max(box.cx, box.cy) * 0.75);
      gradient = ctx.createRadialGradient(cx, cy, 0, cx, cy, radius);
    } else {
      const { x0, y0, x1, y1 } = gradientVectorForBox(box, fill.angle);
      gradient = ctx.createLinearGradient(x0, y0, x1, y1);
    }
    for (const stop of ensureArray(fill.stops)) {
      gradient.addColorStop(
        clamp((Number(stop?.pos) || 0) / 100000, 0, 1),
        toCanvasColor(stop?.color || "#FFFFFF", stop?.alpha ?? 1)
      );
    }
    return gradient;
  }
  if (!fill.color) {
    return null;
  }
  return toCanvasColor(fill.color, fill.alpha ?? 1);
}

function mapPathFillMode(fillMode) {
  switch (String(fillMode || "norm").toLowerCase()) {
    case "lightenless":
      return { type: "lighten", amount: 0.84 };
    case "lighten":
      return { type: "lighten", amount: 0.7 };
    case "darkenless":
      return { type: "darken", amount: 0.2 };
    case "darken":
      return { type: "darken", amount: 0.35 };
    default:
      return null;
  }
}

function transformFillColor(color, fillMode) {
  const mode = mapPathFillMode(fillMode);
  if (!mode || !color) {
    return color;
  }
  return mode.type === "lighten"
    ? lighten(color, mode.amount)
    : darken(color, mode.amount);
}

function resolvePathFill(fill, fillMode) {
  const mode = String(fillMode || "norm").toLowerCase();
  if (!fill || mode === "none") {
    return null;
  }
  if (mode === "norm") {
    return fill;
  }
  if (fill.type === "gradient") {
    return {
      ...fill,
      stops: ensureArray(fill.stops).map((stop) => ({
        ...stop,
        color: transformFillColor(stop?.color, mode)
      }))
    };
  }
  if (!fill.color) {
    return fill;
  }
  return {
    ...fill,
    color: transformFillColor(fill.color, mode)
  };
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
    lineWidthPx * 2,
    lineWidthPx * lineEndScale(lineEnd.length, 3.0, 4.0, 5.0)
  );
  const markerHalfWidth = Math.max(
    lineWidthPx * 0.75,
    (lineWidthPx * lineEndScale(lineEnd.width, 2.0, 3.0, 4.5)) / 2
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
      ctx.lineWidth = Math.max(1, lineWidthPx * lineEndScale(lineEnd.width, 0.8, 1.0, 1.2));
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

function drawPolylineSegment(ctx, points, line, scaleX, scaleY) {
  if (!points || points.length < 2) {
    return;
  }
  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  for (let i = 1; i < points.length; i += 1) {
    ctx.lineTo(points[i].x, points[i].y);
  }
  const strokeStyle = strokeWithElementLine(ctx, line, scaleX, scaleY);
  if (strokeStyle) {
    drawLineEnd(ctx, line?.headEnd, points[0], points[1], strokeStyle.strokeStyle, strokeStyle.widthPx);
    drawLineEnd(
      ctx,
      line?.tailEnd,
      points[points.length - 1],
      points[points.length - 2],
      strokeStyle.strokeStyle,
      strokeStyle.widthPx
    );
  }
}

function connectorPolylinePoints(box, shapeType) {
  const normalized = String(shapeType || "").toLowerCase();
  const start = { x: box.x, y: box.y };
  const end = { x: box.x + box.cx, y: box.y + box.cy };
  switch (normalized) {
    case "bentconnector2":
      return [start, { x: end.x, y: start.y }, end];
    case "bentconnector3": {
      const midX = box.x + box.cx * 0.5;
      return [start, { x: midX, y: start.y }, { x: midX, y: end.y }, end];
    }
    case "bentconnector4": {
      const x1 = box.x + box.cx * 0.35;
      const midY = box.y + box.cy * 0.5;
      return [start, { x: x1, y: start.y }, { x: x1, y: midY }, { x: end.x, y: midY }, end];
    }
    case "bentconnector5": {
      const x1 = box.x + box.cx * 0.33;
      const x2 = box.x + box.cx * 0.66;
      const midY = box.y + box.cy * 0.5;
      return [
        start,
        { x: x1, y: start.y },
        { x: x1, y: midY },
        { x: x2, y: midY },
        { x: x2, y: end.y },
        end
      ];
    }
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

function mapGeometryPoint(box, pathW, pathH, x, y) {
  const sx = pathW ? x / pathW : 0;
  const sy = pathH ? y / pathH : 0;
  return {
    x: box.x + sx * box.cx,
    y: box.y + sy * box.cy
  };
}

function resolveGeometryCoordSpace(geometry, pathW, pathH, boxW, boxH) {
  const presetKey = String(geometry?.preset || "").toLowerCase();
  const isPresetGeometry = shouldAspectFitPresetGeometry(presetKey);
  if (!isPresetGeometry) {
    return {
      vars: buildGeometryVars(geometry, pathW, pathH),
      coordW: pathW,
      coordH: pathH
    };
  }

  const fit = fitGeometryExtents(pathW, pathH, boxW, boxH);
  return {
    vars: buildGeometryVars(geometry, pathW, pathH, fit),
    coordW: Math.max(1, Number(fit.shapeW || pathW)),
    coordH: Math.max(1, Number(fit.shapeH || pathH))
  };
}

function traceCustomGeometryPath(ctx, geometry, path, box) {
  const pathW = Math.max(1, Number(path?.w || geometry?.pathDefaults?.w || 21600));
  const pathH = Math.max(1, Number(path?.h || geometry?.pathDefaults?.h || 21600));
  const { vars, coordW, coordH } = resolveGeometryCoordSpace(geometry, pathW, pathH, box.cx, box.cy);

  let traced = false;
  let currentRaw = null;
  let subPathStartRaw = null;
  for (const cmd of ensureArray(path?.commands)) {
    const type = String(cmd?.type || "");
    if (type === "moveTo" || type === "lnTo") {
      const rawX = evalGeomFormula(cmd?.x || "0", vars);
      const rawY = evalGeomFormula(cmd?.y || "0", vars);
      const p = mapGeometryPoint(box, coordW, coordH, rawX, rawY);
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
          coordW,
          coordH,
          evalGeomFormula(pts[0]?.x || "0", vars),
          evalGeomFormula(pts[0]?.y || "0", vars)
        );
        const ep = mapGeometryPoint(
          box,
          coordW,
          coordH,
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
          coordW,
          coordH,
          evalGeomFormula(pts[0]?.x || "0", vars),
          evalGeomFormula(pts[0]?.y || "0", vars)
        );
        const cp2 = mapGeometryPoint(
          box,
          coordW,
          coordH,
          evalGeomFormula(pts[1]?.x || "0", vars),
          evalGeomFormula(pts[1]?.y || "0", vars)
        );
        const ep = mapGeometryPoint(
          box,
          coordW,
          coordH,
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
        const center = mapGeometryPoint(box, coordW, coordH, arc.cx, arc.cy);
        const rx = Math.abs((arc.rx / coordW) * box.cx);
        const ry = Math.abs((arc.ry / coordH) * box.cy);
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

    const fillStyle = createCanvasFillStyle(ctx, resolvePathFill(element.fill, path?.fill), box);
    if (fillStyle && pathFillEnabled(path)) {
      ctx.fillStyle = fillStyle;
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

function resolveRenderableGeometry(element) {
  if (element?.geometry?.kind === "cust") {
    return element.geometry;
  }
  if (element?.geometry?.kind === "prst") {
    return resolvePresetShapeGeometry(element.shapeType || element.geometry?.preset, element.geometry);
  }
  return null;
}

function resolveGeometryTextBox(box, geometry) {
  const textRect = geometry?.textRect;
  if (!textRect) {
    return box;
  }

  const pathW = Math.max(1, Number(geometry?.pathDefaults?.w || 21600));
  const pathH = Math.max(1, Number(geometry?.pathDefaults?.h || 21600));
  const { vars, coordW, coordH } = resolveGeometryCoordSpace(geometry, pathW, pathH, box.cx, box.cy);
  const left = evalGeomFormula(textRect.l ?? "l", vars);
  const top = evalGeomFormula(textRect.t ?? "t", vars);
  const right = evalGeomFormula(textRect.r ?? "r", vars);
  const bottom = evalGeomFormula(textRect.b ?? "b", vars);

  const mappedLeft = box.x + (left / coordW) * box.cx;
  const mappedTop = box.y + (top / coordH) * box.cy;
  const mappedRight = box.x + (right / coordW) * box.cx;
  const mappedBottom = box.y + (bottom / coordH) * box.cy;

  return {
    x: Math.min(mappedLeft, mappedRight),
    y: Math.min(mappedTop, mappedBottom),
    cx: Math.max(0, Math.abs(mappedRight - mappedLeft)),
    cy: Math.max(0, Math.abs(mappedBottom - mappedTop))
  };
}

function presetAdjustValue(element, name, fallbackValue, min = 0, max = 100000) {
  const gd = ensureArray(element?.geometry?.adjustValues).find((entry) => String(entry?.name || "").toLowerCase() === String(name || "").toLowerCase());
  const formula = String(gd?.fmla || "").trim();
  if (!formula) {
    return fallbackValue;
  }
  const match = formula.match(/^val\s+(-?\d+)$/i);
  if (!match) {
    return fallbackValue;
  }
  return clamp(Number.parseInt(match[1], 10), min, max);
}

function trianglePoints(box, element) {
  const adj = presetAdjustValue(element, "adj", 50000) / 100000;
  return [
    { x: box.x + box.cx * adj, y: box.y },
    { x: box.x + box.cx, y: box.y + box.cy },
    { x: box.x, y: box.y + box.cy }
  ];
}

function pointFromRatio(box, xRatio, yRatio) {
  return {
    x: box.x + box.cx * xRatio,
    y: box.y + box.cy * yRatio
  };
}

function pointsFromRatios(box, ratios) {
  return ratios.map(([xRatio, yRatio]) => pointFromRatio(box, xRatio, yRatio));
}

function regularPolygonPoints(box, sides, rotationDeg = -90) {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  const rx = box.cx / 2;
  const ry = box.cy / 2;
  const start = (rotationDeg * Math.PI) / 180;
  const points = [];
  for (let i = 0; i < sides; i += 1) {
    const angle = start + (i * Math.PI * 2) / sides;
    points.push({
      x: cx + Math.cos(angle) * rx,
      y: cy + Math.sin(angle) * ry
    });
  }
  return points;
}

function starPolygonPoints(box, points, innerRatio = 0.45, rotationDeg = -90) {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  const rx = box.cx / 2;
  const ry = box.cy / 2;
  const start = (rotationDeg * Math.PI) / 180;
  const clampedInner = clamp(innerRatio, 0.1, 0.95);
  const vertices = [];
  for (let i = 0; i < points * 2; i += 1) {
    const angle = start + (i * Math.PI) / points;
    const radiusRatio = i % 2 === 0 ? 1 : clampedInner;
    vertices.push({
      x: cx + Math.cos(angle) * rx * radiusRatio,
      y: cy + Math.sin(angle) * ry * radiusRatio
    });
  }
  return vertices;
}

function blockArrowPoints(box, direction, bodyRatio = 0.36, headRatio = 0.34) {
  const clampedBody = clamp(bodyRatio, 0.12, 0.82);
  const clampedHead = clamp(headRatio, 0.18, 0.75);
  if (direction === "right") {
    const halfBody = (box.cy * clampedBody) / 2;
    const head = box.cx * clampedHead;
    const bodyEnd = box.x + box.cx - head;
    const midY = box.y + box.cy / 2;
    return [
      { x: box.x, y: midY - halfBody },
      { x: bodyEnd, y: midY - halfBody },
      { x: bodyEnd, y: box.y },
      { x: box.x + box.cx, y: midY },
      { x: bodyEnd, y: box.y + box.cy },
      { x: bodyEnd, y: midY + halfBody },
      { x: box.x, y: midY + halfBody }
    ];
  }
  if (direction === "left") {
    const halfBody = (box.cy * clampedBody) / 2;
    const head = box.cx * clampedHead;
    const bodyStart = box.x + head;
    const midY = box.y + box.cy / 2;
    return [
      { x: box.x, y: midY },
      { x: bodyStart, y: box.y },
      { x: bodyStart, y: midY - halfBody },
      { x: box.x + box.cx, y: midY - halfBody },
      { x: box.x + box.cx, y: midY + halfBody },
      { x: bodyStart, y: midY + halfBody },
      { x: bodyStart, y: box.y + box.cy }
    ];
  }
  if (direction === "up") {
    const halfBody = (box.cx * clampedBody) / 2;
    const head = box.cy * clampedHead;
    const bodyStart = box.y + head;
    const midX = box.x + box.cx / 2;
    return [
      { x: midX, y: box.y },
      { x: box.x + box.cx, y: bodyStart },
      { x: midX + halfBody, y: bodyStart },
      { x: midX + halfBody, y: box.y + box.cy },
      { x: midX - halfBody, y: box.y + box.cy },
      { x: midX - halfBody, y: bodyStart },
      { x: box.x, y: bodyStart }
    ];
  }

  const halfBody = (box.cx * clampedBody) / 2;
  const head = box.cy * clampedHead;
  const bodyEnd = box.y + box.cy - head;
  const midX = box.x + box.cx / 2;
  return [
    { x: midX - halfBody, y: box.y },
    { x: midX + halfBody, y: box.y },
    { x: midX + halfBody, y: bodyEnd },
    { x: box.x + box.cx, y: bodyEnd },
    { x: midX, y: box.y + box.cy },
    { x: box.x, y: bodyEnd },
    { x: midX - halfBody, y: bodyEnd }
  ];
}

function doubleArrowPoints(box, orientation, bodyRatio = 0.32, headRatio = 0.24) {
  const clampedBody = clamp(bodyRatio, 0.12, 0.82);
  const clampedHead = clamp(headRatio, 0.15, 0.35);
  if (orientation === "vertical") {
    const halfBody = (box.cx * clampedBody) / 2;
    const head = box.cy * clampedHead;
    const midX = box.x + box.cx / 2;
    const bodyTop = box.y + head;
    const bodyBottom = box.y + box.cy - head;
    return [
      { x: midX, y: box.y },
      { x: box.x + box.cx, y: bodyTop },
      { x: midX + halfBody, y: bodyTop },
      { x: midX + halfBody, y: bodyBottom },
      { x: box.x + box.cx, y: bodyBottom },
      { x: midX, y: box.y + box.cy },
      { x: box.x, y: bodyBottom },
      { x: midX - halfBody, y: bodyBottom },
      { x: midX - halfBody, y: bodyTop },
      { x: box.x, y: bodyTop }
    ];
  }

  const halfBody = (box.cy * clampedBody) / 2;
  const head = box.cx * clampedHead;
  const midY = box.y + box.cy / 2;
  const bodyLeft = box.x + head;
  const bodyRight = box.x + box.cx - head;
  return [
    { x: box.x, y: midY },
    { x: bodyLeft, y: box.y },
    { x: bodyLeft, y: midY - halfBody },
    { x: bodyRight, y: midY - halfBody },
    { x: bodyRight, y: box.y },
    { x: box.x + box.cx, y: midY },
    { x: bodyRight, y: box.y + box.cy },
    { x: bodyRight, y: midY + halfBody },
    { x: bodyLeft, y: midY + halfBody },
    { x: bodyLeft, y: box.y + box.cy }
  ];
}

function quadArrowPoints(box, bodyRatio = 0.28, headRatio = 0.24) {
  const halfBodyX = (box.cx * clamp(bodyRatio, 0.12, 0.6)) / 2;
  const halfBodyY = (box.cy * clamp(bodyRatio, 0.12, 0.6)) / 2;
  const headX = box.cx * clamp(headRatio, 0.12, 0.35);
  const headY = box.cy * clamp(headRatio, 0.12, 0.35);
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  return [
    { x: cx, y: box.y },
    { x: cx + headX, y: box.y + headY },
    { x: cx + halfBodyX, y: box.y + headY },
    { x: cx + halfBodyX, y: cy - halfBodyY },
    { x: box.x + box.cx - headX, y: cy - halfBodyY },
    { x: box.x + box.cx - headX, y: box.y + headY },
    { x: box.x + box.cx, y: cy },
    { x: box.x + box.cx - headX, y: box.y + box.cy - headY },
    { x: box.x + box.cx - headX, y: cy + halfBodyY },
    { x: cx + halfBodyX, y: cy + halfBodyY },
    { x: cx + halfBodyX, y: box.y + box.cy - headY },
    { x: cx, y: box.y + box.cy },
    { x: cx - headX, y: box.y + box.cy - headY },
    { x: cx - halfBodyX, y: box.y + box.cy - headY },
    { x: cx - halfBodyX, y: cy + halfBodyY },
    { x: box.x + headX, y: cy + halfBodyY },
    { x: box.x + headX, y: box.y + box.cy - headY },
    { x: box.x, y: cy },
    { x: box.x + headX, y: box.y + headY },
    { x: box.x + headX, y: cy - halfBodyY },
    { x: cx - halfBodyX, y: cy - halfBodyY },
    { x: cx - halfBodyX, y: box.y + headY },
    { x: cx - headX, y: box.y + headY }
  ];
}

function cumulativeOffsets(values) {
  const offsets = [0];
  for (const value of values) {
    offsets.push(offsets[offsets.length - 1] + Math.max(0, Number(value) || 0));
  }
  return offsets;
}

function offsetSpan(offsets, startIndex, span) {
  const start = clamp(startIndex, 0, Math.max(0, offsets.length - 1));
  const end = clamp(start + Math.max(1, span), 0, Math.max(0, offsets.length - 1));
  return Math.max(0, offsets[end] - offsets[start]);
}

const PRESET_WEDGE_RECT_CALLOUT_ADJUSTS = [
  { name: "adj1", fmla: "val -20833" },
  { name: "adj2", fmla: "val 62500" }
];

const PRESET_WEDGE_RECT_CALLOUT_GUIDES = [
  { name: "dxPos", fmla: "*/ w adj1 100000" },
  { name: "dyPos", fmla: "*/ h adj2 100000" },
  { name: "xPos", fmla: "+- hc dxPos 0" },
  { name: "yPos", fmla: "+- vc dyPos 0" },
  { name: "dx", fmla: "+- xPos 0 hc" },
  { name: "dy", fmla: "+- yPos 0 vc" },
  { name: "dq", fmla: "*/ dxPos h w" },
  { name: "ady", fmla: "abs dyPos" },
  { name: "adq", fmla: "abs dq" },
  { name: "dz", fmla: "+- ady 0 adq" },
  { name: "xg1", fmla: "?: dxPos 7 2" },
  { name: "xg2", fmla: "?: dxPos 10 5" },
  { name: "x1", fmla: "*/ w xg1 12" },
  { name: "x2", fmla: "*/ w xg2 12" },
  { name: "yg1", fmla: "?: dyPos 7 2" },
  { name: "yg2", fmla: "?: dyPos 10 5" },
  { name: "y1", fmla: "*/ h yg1 12" },
  { name: "y2", fmla: "*/ h yg2 12" },
  { name: "t1", fmla: "?: dxPos l xPos" },
  { name: "xl", fmla: "?: dz l t1" },
  { name: "t2", fmla: "?: dyPos x1 xPos" },
  { name: "xt", fmla: "?: dz t2 x1" },
  { name: "t3", fmla: "?: dxPos xPos r" },
  { name: "xr", fmla: "?: dz r t3" },
  { name: "t4", fmla: "?: dyPos xPos x1" },
  { name: "xb", fmla: "?: dz t4 x1" },
  { name: "t5", fmla: "?: dxPos y1 yPos" },
  { name: "yl", fmla: "?: dz y1 t5" },
  { name: "t6", fmla: "?: dyPos t yPos" },
  { name: "yt", fmla: "?: dz t6 t" },
  { name: "t7", fmla: "?: dxPos yPos y1" },
  { name: "yr", fmla: "?: dz y1 t7" },
  { name: "t8", fmla: "?: dyPos yPos b" },
  { name: "yb", fmla: "?: dz t8 b" }
];

function mergePresetAdjustValues(defaults, overrides) {
  const overrideMap = new Map(
    ensureArray(overrides)
      .filter((entry) => entry?.name)
      .map((entry) => [String(entry.name), entry])
  );
  return defaults.map((entry) => overrideMap.get(entry.name) || entry);
}

function resolvePresetGuideVars(box, defaultAdjusts, guideValues, element) {
  const vars = {
    w: box.cx,
    h: box.cy,
    l: 0,
    t: 0,
    r: box.cx,
    b: box.cy,
    hc: box.cx / 2,
    vc: box.cy / 2
  };
  const mergedAdjusts = mergePresetAdjustValues(defaultAdjusts, element?.geometry?.adjustValues);
  const withAdjusts = resolveGeometryGuides(mergedAdjusts, vars);
  return resolveGeometryGuides(guideValues, withAdjusts);
}

function setPathForWedgeRectCallout(ctx, box, element) {
  const vars = resolvePresetGuideVars(
    box,
    PRESET_WEDGE_RECT_CALLOUT_ADJUSTS,
    PRESET_WEDGE_RECT_CALLOUT_GUIDES,
    element
  );
  const points = [
    { x: vars.l, y: vars.t },
    { x: vars.x1, y: vars.t },
    { x: vars.xt, y: vars.yt },
    { x: vars.x2, y: vars.t },
    { x: vars.r, y: vars.t },
    { x: vars.r, y: vars.y1 },
    { x: vars.xr, y: vars.yr },
    { x: vars.r, y: vars.y2 },
    { x: vars.r, y: vars.b },
    { x: vars.x2, y: vars.b },
    { x: vars.xb, y: vars.yb },
    { x: vars.x1, y: vars.b },
    { x: vars.l, y: vars.b },
    { x: vars.l, y: vars.y2 },
    { x: vars.xl, y: vars.yl },
    { x: vars.l, y: vars.y1 }
  ];

  ctx.moveTo(box.x + points[0].x, box.y + points[0].y);
  for (let i = 1; i < points.length; i += 1) {
    ctx.lineTo(box.x + points[i].x, box.y + points[i].y);
  }
  ctx.closePath();
}

function tracePresetShapePart(ctx, part) {
  switch (part?.kind) {
    case "polygon":
      polygonPath(ctx, ensureArray(part.points));
      break;
    case "loops":
      for (const loop of ensureArray(part.loops)) {
        polygonPath(ctx, ensureArray(loop));
      }
      break;
    case "polyline": {
      const points = ensureArray(part.points);
      if (!points.length) {
        break;
      }
      ctx.moveTo(points[0].x, points[0].y);
      for (let i = 1; i < points.length; i += 1) {
        ctx.lineTo(points[i].x, points[i].y);
      }
      if (part.closed) {
        ctx.closePath();
      }
      break;
    }
    case "ellipse":
      ctx.moveTo(part.cx + part.rx, part.cy);
      ctx.ellipse(part.cx, part.cy, part.rx, part.ry, 0, 0, Math.PI * 2);
      break;
    case "roundRect": {
      const x = part.x;
      const y = part.y;
      const w = part.w;
      const h = part.h;
      const radius = Math.max(0, Math.min(part.r || 0, w / 2, h / 2));
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
    case "rect":
      ctx.rect(part.x, part.y, part.w, part.h);
      break;
    default:
      break;
  }
}

function setPathForShape(ctx, shapeType, box, element = null) {
  const presetParts = buildPresetShapeParts(shapeType, box, element);
  if (presetParts?.length) {
    for (const part of presetParts) {
      tracePresetShapePart(ctx, part);
    }
    return;
  }

  const normalized = String(shapeType || "rect").toLowerCase();
  const regularPolygonSides = {
    heptagon: 7,
    octagon: 8,
    decagon: 10,
    dodecagon: 12
  };
  if (regularPolygonSides[normalized]) {
    polygonPath(ctx, regularPolygonPoints(box, regularPolygonSides[normalized]));
    return;
  }

  const starMatch = normalized.match(/^star(\d+)$/);
  if (starMatch) {
    const starPoints = Number.parseInt(starMatch[1], 10);
    const innerRatio = presetAdjustValue(element, "adj", 38000, 5000, 95000) / 100000;
    polygonPath(ctx, starPolygonPoints(box, starPoints, innerRatio));
    return;
  }

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
    case "plus": {
      polygonPath(ctx, pointsFromRatios(box, [
        [0.35, 0],
        [0.65, 0],
        [0.65, 0.35],
        [1, 0.35],
        [1, 0.65],
        [0.65, 0.65],
        [0.65, 1],
        [0.35, 1],
        [0.35, 0.65],
        [0, 0.65],
        [0, 0.35],
        [0.35, 0.35]
      ]));
      break;
    }
    case "homeplate": {
      polygonPath(ctx, pointsFromRatios(box, [
        [0, 0],
        [0.72, 0],
        [1, 0.5],
        [0.72, 1],
        [0, 1]
      ]));
      break;
    }
    case "rightarrow": {
      polygonPath(ctx, blockArrowPoints(box, "right"));
      break;
    }
    case "leftarrow": {
      polygonPath(ctx, blockArrowPoints(box, "left"));
      break;
    }
    case "uparrow": {
      polygonPath(ctx, blockArrowPoints(box, "up"));
      break;
    }
    case "downarrow": {
      polygonPath(ctx, blockArrowPoints(box, "down"));
      break;
    }
    case "leftrightarrow": {
      polygonPath(ctx, doubleArrowPoints(box, "horizontal"));
      break;
    }
    case "updownarrow": {
      polygonPath(ctx, doubleArrowPoints(box, "vertical"));
      break;
    }
    case "quadarrow": {
      polygonPath(ctx, quadArrowPoints(box));
      break;
    }
    case "notchedrightarrow": {
      polygonPath(ctx, pointsFromRatios(box, [
        [0, 0.28],
        [0.62, 0.28],
        [0.62, 0],
        [1, 0.5],
        [0.62, 1],
        [0.62, 0.72],
        [0.16, 0.72],
        [0, 0.5]
      ]));
      break;
    }
    case "wedgerectcallout": {
      setPathForWedgeRectCallout(ctx, box, element);
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
  const renderGeometry = resolveRenderableGeometry(element);
  const renderElement = renderGeometry ? { ...element, geometry: renderGeometry } : element;
  if (isLineLikeShapeType(element.shapeType)) {
    if (renderGeometry?.kind === "cust") {
      withElementTransform(ctx, box, () => {
        drawLineSegment(ctx, box, renderElement.line, scaleX, scaleY);
      });
      return;
    }
    drawLineSegment(ctx, box, element.line, scaleX, scaleY);
    return;
  }

  withElementTransform(ctx, box, () => {
    if (renderGeometry?.kind === "cust") {
      const drawn = drawCustomGeometry(ctx, renderElement, box, scaleX, scaleY);
      if (drawn) {
        return;
      }
    }

    ctx.beginPath();
    setPathForShape(ctx, element.shapeType, box, element);

    const fillStyle = createCanvasFillStyle(ctx, renderElement.fill, box);
    if (fillStyle) {
      ctx.fillStyle = fillStyle;
      ctx.fill();
    }

    strokeWithElementLine(ctx, renderElement.line, scaleX, scaleY);
  });
}

function drawLineElement(ctx, element, scaleX, scaleY) {
  const box = toPxElement(element, scaleX, scaleY);
  const renderGeometry = resolveRenderableGeometry(element);
  const renderElement = renderGeometry ? { ...element, geometry: renderGeometry } : element;
  const connectorPoints = connectorPolylinePoints(box, element?.shapeType);
  if (connectorPoints) {
    withElementTransform(ctx, box, () => {
      drawPolylineSegment(ctx, connectorPoints, element.line, scaleX, scaleY);
    });
    return;
  }
  if (renderGeometry?.kind === "cust") {
    withElementTransform(ctx, box, () => {
      const drawn = drawCustomGeometry(ctx, renderElement, box, scaleX, scaleY);
      if (drawn) {
        return;
      }
      drawLineSegment(ctx, box, renderElement.line, scaleX, scaleY);
    });
    return;
  }
  drawLineSegment(ctx, box, element.line, scaleX, scaleY);
}

const FONT_FAMILY_ALIASES = new Map([
  ["ＭＳ Ｐゴシック", "MS PGothic"],
  ["ＭＳ ゴシック", "MS Gothic"],
  ["ＭＳ Ｐ明朝", "MS PMincho"],
  ["ＭＳ 明朝", "MS Mincho"]
]);

function normalizeFontFamilyName(name) {
  const normalized = String(name || "").trim();
  if (!normalized) {
    return null;
  }
  return FONT_FAMILY_ALIASES.get(normalized) || normalized;
}

function normalizeRunStyle(style = {}, defaultStyle = {}) {
  return {
    fontFamily: normalizeFontFamilyName(style.fontFamily || defaultStyle.fontFamily || "Calibri"),
    eastAsiaFont: normalizeFontFamilyName(style.eastAsiaFont || defaultStyle.eastAsiaFont || null),
    complexScriptFont: normalizeFontFamilyName(style.complexScriptFont || defaultStyle.complexScriptFont || null),
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
  return [
    style.fontFamily,
    style.eastAsiaFont,
    style.complexScriptFont,
    style.fontSizePt,
    style.bold ? 1 : 0,
    style.italic ? 1 : 0,
    style.color,
    style.alpha,
    style.underline ? 1 : 0,
    style.spacing ?? 0
  ].join("|");
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
    quoteFontFamily(style.complexScriptFont || ""),
    "\"Yu Gothic UI\"",
    "\"Yu Gothic\"",
    "\"Meiryo UI\"",
    "Meiryo",
    "\"MS PGothic\"",
    "\"MS Gothic\"",
    "sans-serif"
  ].filter(Boolean).join(", ");
  ctx.font = `${italic} ${weight} ${sizePx}px ${families}`;
  ctx.fillStyle = toCanvasColor(style.color, style.alpha);
  return sizePx;
}

function measureStyledTextWidth(ctx, text, style) {
  if (!text) {
    return 0;
  }
  setCanvasFont(ctx, style);
  const glyphCount = Array.from(text).length;
  return ctx.measureText(text).width + Math.max(0, glyphCount - 1) * charSpacingPx(style);
}

function pushTextSegment(line, text, style, width) {
  if (!text) {
    return;
  }
  line.explicitBreak = false;
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
    lineSpacing: paragraph?.lineSpacing || 0,
    lineSpacingPt: paragraph?.lineSpacingPt || 0,
    explicitBreak: false
  };
}

function flushLine(lines, line, allowEmpty = false) {
  if (line.segments.length || allowEmpty) {
    lines.push(line);
  }
}

function lineDisplayWidth(line) {
  if (!line?.segments?.length) {
    return 0;
  }
  return Math.max(0, line.width);
}

function finalizeLineMeasurements(ctx, line) {
  if (!line?.segments?.length) {
    line.width = 0;
    return line;
  }
  let width = 0;
  for (const seg of line.segments) {
    seg.width = measureStyledTextWidth(ctx, seg.text, seg.style);
    width += seg.width;
  }
  line.width = width;
  return line;
}

function wrapParagraphRuns(ctx, paragraph, boxWidthPx, defaultStyle, options = {}) {
  const lines = [];
  let line = newLine(paragraph);
  let hadVisibleText = false;
  const wrapEnabled = options.wrapEnabled !== false;

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
        line.explicitBreak = true;
        flushLine(lines, line, true);
        line = newLine(paragraph);
        line.maxFontSizePx = Math.max(line.maxFontSizePx, fontPx);
        line.maxAscent = Math.max(line.maxAscent, metrics.ascent);
        line.maxDescent = Math.max(line.maxDescent, metrics.descent);
        line.explicitBreak = true;
        continue;
      }

      hadVisibleText = true;
      const width = ctx.measureText(ch).width + charSpacingPx(style);
      if (wrapEnabled && line.width + width > boxWidthPx && line.segments.length) {
        flushLine(lines, line);
        line = newLine(paragraph);
        line.maxFontSizePx = Math.max(line.maxFontSizePx, fontPx);
        line.maxAscent = Math.max(line.maxAscent, metrics.ascent);
        line.maxDescent = Math.max(line.maxDescent, metrics.descent);
      }
      pushTextSegment(line, ch, style, width);
    }
  }

  if (!hadVisibleText && !line.segments.length) {
    line.explicitBreak = true;
  }
  flushLine(lines, line, line.explicitBreak);
  for (const item of lines) {
    finalizeLineMeasurements(ctx, item);
  }
  return lines;
}

function lineHeightPx(line) {
  const natural = Math.max(line.maxAscent + line.maxDescent, (line.maxFontSizePx || 0) * 1.05);
  if (line.lineSpacingPt) {
    return Math.max(natural, line.lineSpacingPt * PX_PER_PT);
  }
  const spacing = line.lineSpacing ? Math.max(1, line.lineSpacing / 100000) : 1;
  return natural * spacing;
}

function alignmentToStartX(left, width, lineWidth, alignment, options = {}) {
  const allowOverflow = options.allowOverflow === true;
  if (alignment === "ctr") {
    const offset = (width - lineWidth) / 2;
    return left + (allowOverflow ? offset : Math.max(0, offset));
  }
  if (alignment === "r") {
    const offset = width - lineWidth;
    return left + (allowOverflow ? offset : Math.max(0, offset));
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
  if (font.includes("wingdings")) {
    if (char === "l") {
      return "\u25cf";
    }
    if (/[\uF000-\uF0FF]/.test(char)) {
      return "\u2022";
    }
  }
  if (!font.includes("wingdings") && /[\uF000-\uF0FF]/.test(char)) {
    return "\u2022";
  }
  return char;
}

function mappedBulletSizeScale(bullet, bulletChar) {
  const font = String(bullet?.fontFamily || "").toLowerCase();
  if (font.includes("wingdings") && bulletChar === "\u25cf") {
    return 0.72;
  }
  return 1;
}

function toAlphaSequence(index, upper = false) {
  let value = Math.max(1, index);
  let out = "";
  while (value > 0) {
    value -= 1;
    out = String.fromCharCode(97 + (value % 26)) + out;
    value = Math.floor(value / 26);
  }
  return upper ? out.toUpperCase() : out;
}

function toRoman(index, upper = false) {
  const numerals = [
    [1000, "m"],
    [900, "cm"],
    [500, "d"],
    [400, "cd"],
    [100, "c"],
    [90, "xc"],
    [50, "l"],
    [40, "xl"],
    [10, "x"],
    [9, "ix"],
    [5, "v"],
    [4, "iv"],
    [1, "i"]
  ];
  let value = Math.max(1, index);
  let out = "";
  for (const [n, token] of numerals) {
    while (value >= n) {
      out += token;
      value -= n;
    }
  }
  return upper ? out.toUpperCase() : out;
}

function formatAutoNumber(bullet, index) {
  const value = Math.max(1, index);
  switch (String(bullet?.format || "arabicPeriod")) {
    case "arabicParenBoth":
      return `(${value})`;
    case "arabicParenR":
      return `${value})`;
    case "alphaLcPeriod":
      return `${toAlphaSequence(value, false)}.`;
    case "alphaLcParenR":
      return `${toAlphaSequence(value, false)})`;
    case "alphaUcPeriod":
      return `${toAlphaSequence(value, true)}.`;
    case "alphaUcParenR":
      return `${toAlphaSequence(value, true)})`;
    case "romanLcPeriod":
      return `${toRoman(value, false)}.`;
    case "romanUcPeriod":
      return `${toRoman(value, true)}.`;
    default:
      return `${value}.`;
  }
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

function scaleRunStyleForCanvas(style, fontScale) {
  if (!style) {
    return style;
  }
  return {
    ...style,
    fontSizePt: (style.fontSizePt || 0) * fontScale
  };
}

function scaleParagraphForCanvas(paragraph, scaleX, fontScale) {
  return {
    ...paragraph,
    marginLeft: (paragraph?.marginLeft || 0) * scaleX,
    marginRight: (paragraph?.marginRight || 0) * scaleX,
    indent: (paragraph?.indent || 0) * scaleX,
    defaultTabSize: (paragraph?.defaultTabSize || 0) * scaleX,
    spaceBefore: (paragraph?.spaceBefore || 0) * fontScale,
    spaceAfter: (paragraph?.spaceAfter || 0) * fontScale,
    lineSpacingPt: (paragraph?.lineSpacingPt || 0) * fontScale,
    bullet: paragraph?.bullet ? {
      ...paragraph.bullet,
      sizePt: paragraph.bullet.sizePt ? paragraph.bullet.sizePt * fontScale : paragraph.bullet.sizePt
    } : paragraph?.bullet,
    runs: ensureArray(paragraph?.runs).map((run) => ({
      ...run,
      style: scaleRunStyleForCanvas(run?.style, fontScale)
    }))
  };
}

function scaleTextBodyForCanvas(textBody, scaleX, scaleY, extraInsets = {}) {
  if (!textBody) {
    return textBody;
  }
  const fontScale = ((scaleX + scaleY) / 2) / BASE_EMU_TO_PX;
  return {
    ...textBody,
    leftInset: ((textBody.leftInset || 0) + (extraInsets.left || 0)) * scaleX,
    rightInset: ((textBody.rightInset || 0) + (extraInsets.right || 0)) * scaleX,
    topInset: ((textBody.topInset || 0) + (extraInsets.top || 0)) * scaleY,
    bottomInset: ((textBody.bottomInset || 0) + (extraInsets.bottom || 0)) * scaleY,
    paragraphs: ensureArray(textBody.paragraphs).map((paragraph) => (
      scaleParagraphForCanvas(paragraph, scaleX, fontScale)
    ))
  };
}

function scaleParagraphForAutoFit(paragraph, scale) {
  return {
    ...paragraph,
    lineSpacingPt: (paragraph?.lineSpacingPt || 0) * scale,
    runs: ensureArray(paragraph?.runs).map((run) => ({
      ...run,
      style: run?.style ? {
        ...run.style,
        fontSizePt: (run.style.fontSizePt || 0) * scale
      } : run?.style
    })),
    bullet: paragraph?.bullet ? {
      ...paragraph.bullet,
      sizePt: paragraph.bullet.sizePt ? paragraph.bullet.sizePt * scale : paragraph.bullet.sizePt
    } : paragraph?.bullet
  };
}

function autoFitNoWrapTextBody(ctx, textBody, width, height, defaultStyle) {
  const paragraphs = ensureArray(textBody?.paragraphs);
  if (!paragraphs.length) {
    return textBody;
  }

  const layouts = [];
  let maxWidth = 0;
  let totalHeight = 0;
  for (const paragraph of paragraphs) {
    const before = Math.max(0, (paragraph.spaceBefore || 0) / 100) * PX_PER_PT;
    const after = Math.max(0, (paragraph.spaceAfter || 0) / 100) * PX_PER_PT;
    const marginLeft = Math.max(0, paragraph.marginLeft || 0);
    const marginRight = Math.max(0, paragraph.marginRight || 0);
    const paragraphWidth = Math.max(0, width - marginLeft - marginRight);
    const lines = wrapParagraphRuns(ctx, paragraph, paragraphWidth, defaultStyle, { wrapEnabled: false });
    const heights = lines.map((line) => lineHeightPx(line));
    const widest = lines.reduce((acc, line) => Math.max(acc, lineDisplayWidth(line)), 0);
    maxWidth = Math.max(maxWidth, widest);
    totalHeight += before + heights.reduce((sum, lineHeight) => sum + lineHeight, 0) + after;
    layouts.push({ paragraphWidth });
  }

  const widthScale = maxWidth > 0 ? width / maxWidth : 1;
  const heightScale = totalHeight > 0 ? height / totalHeight : 1;
  const scale = Math.min(1, widthScale, heightScale);
  if (!(scale > 0) || scale >= 0.999) {
    return textBody;
  }

  return {
    ...textBody,
    paragraphs: paragraphs.map((paragraph) => scaleParagraphForAutoFit(paragraph, scale))
  };
}

function buildTextBodyLayout(ctx, textBody, contentWidth, defaultStyle) {
  const paragraphLayouts = [];
  let totalHeight = 0;
  let maxContentWidth = 0;
  const autoNumberState = new Map();
  const wrapEnabled = String(textBody.wrap || "square").toLowerCase() !== "none";

  for (const paragraph of ensureArray(textBody.paragraphs)) {
    const before = Math.max(0, (paragraph.spaceBefore || 0) / 100) * PX_PER_PT;
    const after = Math.max(0, (paragraph.spaceAfter || 0) / 100) * PX_PER_PT;
    const marginLeft = Math.max(0, paragraph.marginLeft || 0);
    const marginRight = Math.max(0, paragraph.marginRight || 0);
    const indent = paragraph.indent || 0;
    const paragraphWidth = Math.max(0, contentWidth - marginLeft - marginRight);
    const lines = wrapParagraphRuns(ctx, paragraph, paragraphWidth, defaultStyle, { wrapEnabled });
    if (!lines.length) {
      continue;
    }
    const hasVisibleText = ensureArray(paragraph.runs).some((run) => String(run?.text || "").trim().length > 0);
    let bulletLabel = null;
    if (paragraph.bullet?.type === "autoNum" && hasVisibleText) {
      const key = `${paragraph.level || 0}|${paragraph.bullet.format || "arabicPeriod"}|${paragraph.bullet.startAt || 1}`;
      const current = autoNumberState.has(key)
        ? autoNumberState.get(key)
        : (paragraph.bullet.startAt || 1);
      bulletLabel = formatAutoNumber(paragraph.bullet, current);
      autoNumberState.set(key, current + 1);
    }
    const heights = lines.map((line) => lineHeightPx(line));
    const paraHeight = before + heights.reduce((sum, h) => sum + h, 0) + after;
    const widestLine = lines.reduce((acc, line) => Math.max(acc, lineDisplayWidth(line)), 0);
    paragraphLayouts.push({
      lines,
      heights,
      before,
      after,
      marginLeft,
      marginRight,
      indent,
      bullet: paragraph.bullet || null,
      bulletLabel
    });
    totalHeight += paraHeight;
    maxContentWidth = Math.max(maxContentWidth, marginLeft + widestLine + marginRight);
  }

  return {
    paragraphLayouts,
    totalHeight,
    maxContentWidth,
    wrapEnabled
  };
}

const DEFAULT_TEXT_STYLE = {
  fontFamily: "Calibri",
  eastAsiaFont: null,
  complexScriptFont: null,
  fontSizePt: 18,
  color: "#000000",
  alpha: 1,
  bold: false,
  italic: false,
  spacing: 0
};

function shouldExpandTextBoxForAutoFit(element, textBody) {
  if (!element || !textBody || textBody.autoFit !== "shape" || element.type !== "text") {
    return false;
  }
  const fillType = element.fill?.type || null;
  return (fillType === null || fillType === "none") && !element.line;
}

function expandTextBoxForAutoFit(ctx, element, boxPx, textBoxPx, textBody) {
  if (!shouldExpandTextBoxForAutoFit(element, textBody)) {
    return { boxPx, textBoxPx };
  }

  const contentWidth = Math.max(0, textBoxPx.cx - (textBody.leftInset || 0) - (textBody.rightInset || 0));
  const { totalHeight, maxContentWidth, wrapEnabled } = buildTextBodyLayout(ctx, textBody, contentWidth, DEFAULT_TEXT_STYLE);
  const requiredWidth = (textBody.leftInset || 0) + (textBody.rightInset || 0) + maxContentWidth;
  const requiredHeight = (textBody.topInset || 0) + (textBody.bottomInset || 0) + totalHeight;
  const expandedTextBox = {
    ...textBoxPx,
    cx: wrapEnabled ? textBoxPx.cx : Math.max(textBoxPx.cx, requiredWidth),
    cy: Math.max(textBoxPx.cy, requiredHeight)
  };
  return {
    boxPx: {
      ...boxPx,
      cx: boxPx.cx + (expandedTextBox.cx - textBoxPx.cx),
      cy: boxPx.cy + (expandedTextBox.cy - textBoxPx.cy)
    },
    textBoxPx: expandedTextBox
  };
}

function shouldClipTextBody(element, textBody) {
  return !(element?.isTextBox && textBody?.autoFit === "none");
}

function drawTextBodyInBox(ctx, textBody, boxPx, options = {}) {
  if (!textBody) {
    return;
  }

  const left = boxPx.x + (textBody.leftInset || 0);
  const top = boxPx.y + (textBody.topInset || 0);
  const width = Math.max(0, boxPx.cx - (textBody.leftInset || 0) - (textBody.rightInset || 0));
  const height = Math.max(0, boxPx.cy - (textBody.topInset || 0) - (textBody.bottomInset || 0));

  if (String(textBody.direction || "horz").toLowerCase() === "eavert") {
    drawVerticalTextBodyInBox(ctx, textBody, boxPx, DEFAULT_TEXT_STYLE);
    return;
  }

  let effectiveTextBody = textBody;
  if (String(textBody.wrap || "square").toLowerCase() === "none" && textBody.autoFit === "norm") {
    effectiveTextBody = autoFitNoWrapTextBody(ctx, textBody, width, height, DEFAULT_TEXT_STYLE);
  }

  const {
    paragraphLayouts,
    totalHeight,
    wrapEnabled
  } = buildTextBodyLayout(ctx, effectiveTextBody, width, DEFAULT_TEXT_STYLE);

  if (!paragraphLayouts.length) {
    return;
  }

  const clipText = options.clip !== false;
  ctx.save();
  if (wrapEnabled && clipText) {
    ctx.beginPath();
    ctx.rect(left, top, width, height);
    ctx.clip();
  }

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
      const startX = alignmentToStartX(
        paragraphLeft,
        paragraphWidth,
        lineDisplayWidth(line),
        line.alignment,
        { allowOverflow: !wrapEnabled }
      );

      if (li === 0 && (layout.bullet?.type === "char" || layout.bulletLabel)) {
        const baseStyle = line.segments[0]?.style || DEFAULT_TEXT_STYLE;
        const bulletChar = layout.bullet?.type === "char"
          ? normalizeBulletChar(layout.bullet)
          : null;
        const useMappedBulletFont = Boolean(
          layout.bullet?.type === "char"
          && bulletChar
          && bulletChar !== (layout.bullet?.char || "\u2022")
        );
        const bulletStyle = {
          ...baseStyle,
          fontFamily: useMappedBulletFont ? baseStyle.fontFamily : (layout.bullet?.fontFamily || baseStyle.fontFamily),
          eastAsiaFont: useMappedBulletFont ? baseStyle.eastAsiaFont : (layout.bullet?.fontFamily ? null : baseStyle.eastAsiaFont),
          complexScriptFont: useMappedBulletFont ? baseStyle.complexScriptFont : (layout.bullet?.fontFamily ? null : baseStyle.complexScriptFont),
          color: layout.bullet?.color || baseStyle.color,
          alpha: layout.bullet?.alpha ?? baseStyle.alpha,
          fontSizePt: (layout.bullet?.sizePt || (baseStyle.fontSizePt * (layout.bullet?.sizePct || 1)))
            * mappedBulletSizeScale(layout.bullet, bulletChar)
        };
        setCanvasFont(ctx, bulletStyle);
        ctx.textBaseline = "alphabetic";
        const bulletText = layout.bullet?.type === "char" ? bulletChar : layout.bulletLabel;
        ctx.fillText(bulletText, paragraphLeft + Math.min(0, layout.indent), baselineY);
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
  const renderGeometry = resolveRenderableGeometry(element);
  const textBox = renderGeometry ? resolveGeometryTextBox(box, renderGeometry) : box;
  const textBody = scaleTextBodyForCanvas(element.text, scaleX, scaleY);
  const autoFitBox = expandTextBoxForAutoFit(ctx, element, box, textBox, textBody);

  withElementTransform(ctx, autoFitBox.boxPx, () => {
    drawTextBodyInBox(ctx, textBody, autoFitBox.textBoxPx, {
      clip: shouldClipTextBody(element, textBody)
    });
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
  ctx.save();
  ctx.globalAlpha = 1;
  ctx.globalCompositeOperation = "source-over";
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
  ctx.restore();
}

function drawTableCellText(ctx, cell, x, y, width, height, scaleX, scaleY) {
  if (!cell?.text) {
    return;
  }
  const textBody = {
    ...scaleTextBodyForCanvas(cell.text, scaleX, scaleY, {
      left: cell.marginLeft || 0,
      right: cell.marginRight || 0,
      top: cell.marginTop || 0,
      bottom: cell.marginBottom || 0
    }),
    verticalAlign: cell.verticalAlign || cell.text.verticalAlign,
  };
  drawTextBodyInBox(ctx, textBody, { x, y, cx: width, cy: height });
}

function drawTableCellBorder(ctx, side, border, x, y, width, height, scaleX, scaleY) {
  if (!border?.color || !border?.width) {
    return;
  }

  ctx.save();
  ctx.strokeStyle = toCanvasColor(border.color, border.alpha ?? 1);
  ctx.lineWidth = Math.max(0.75, lineWidthToPx(border.width, scaleX, scaleY));
  ctx.lineCap = "butt";
  ctx.beginPath();
  switch (side) {
    case "left":
      ctx.moveTo(x, y);
      ctx.lineTo(x, y + height);
      break;
    case "right":
      ctx.moveTo(x + width, y);
      ctx.lineTo(x + width, y + height);
      break;
    case "top":
      ctx.moveTo(x, y);
      ctx.lineTo(x + width, y);
      break;
    case "bottom":
      ctx.moveTo(x, y + height);
      ctx.lineTo(x + width, y + height);
      break;
    default:
      break;
  }
  ctx.stroke();
  ctx.restore();
}

function drawTable(ctx, element, scaleX, scaleY) {
  const box = toPxElement(element, scaleX, scaleY);

  withElementTransform(ctx, box, () => {
    const columnCount = Math.max(
      element.gridCols.length || 0,
      ...element.rows.map((row) => row.cells.length || 0),
      1
    );
    const fallbackColumnWidth = columnCount ? element.cx / columnCount : element.cx;
    const rawColumnWidths = Array.from({ length: columnCount }, (_, index) => element.gridCols[index] || fallbackColumnWidth);
    const totalWidth = rawColumnWidths.reduce((sum, width) => sum + width, 0) || element.cx;
    const widthScale = totalWidth ? box.cx / totalWidth : 1;
    const rowHeights = element.rows.map((row) => row.height || (element.rows.length ? element.cy / element.rows.length : element.cy));
    const columnOffsets = cumulativeOffsets(rawColumnWidths.map((width) => width * widthScale));
    const rowOffsets = cumulativeOffsets(rowHeights.map((height) => height * scaleY));
    const occupied = new Set();

    for (let ri = 0; ri < element.rows.length; ri += 1) {
      const row = element.rows[ri];
      let logicalCol = 0;
      for (let ci = 0; ci < row.cells.length; ci += 1) {
        const cell = row.cells[ci];
        if (cell.hMerge || cell.vMerge) {
          logicalCol += 1;
          continue;
        }

        while (occupied.has(`${ri}:${logicalCol}`)) {
          logicalCol += 1;
        }

        const colIndex = logicalCol;

        const spanCols = Math.max(1, cell.gridSpan || 1);
        const spanRows = Math.max(1, cell.rowSpan || 1);
        const cellX = box.x + columnOffsets[Math.min(colIndex, columnOffsets.length - 1)];
        const cellY = box.y + rowOffsets[Math.min(ri, rowOffsets.length - 1)];
        const cellWidth = offsetSpan(columnOffsets, colIndex, spanCols);
        const cellHeight = offsetSpan(rowOffsets, ri, spanRows);

        ctx.fillStyle = toCanvasColor(cell.fill?.color || "#FFFFFF", cell.fill?.alpha ?? 1);
        ctx.fillRect(cellX, cellY, cellWidth, cellHeight);

        for (const side of ["left", "right", "top", "bottom"]) {
          drawTableCellBorder(ctx, side, cell.borders?.[side], cellX, cellY, cellWidth, cellHeight, scaleX, scaleY);
        }

        drawTableCellText(ctx, cell, cellX, cellY, cellWidth, cellHeight, scaleX, scaleY);

        for (let spanColIndex = colIndex + 1; spanColIndex < Math.min(columnCount, colIndex + spanCols); spanColIndex += 1) {
          occupied.add(`${ri}:${spanColIndex}`);
        }
        for (let spanRowIndex = ri + 1; spanRowIndex < Math.min(element.rows.length, ri + spanRows); spanRowIndex += 1) {
          for (let spanColIndex = colIndex; spanColIndex < Math.min(columnCount, colIndex + spanCols); spanColIndex += 1) {
            occupied.add(`${spanRowIndex}:${spanColIndex}`);
          }
        }

        logicalCol = colIndex + 1;
      }
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
