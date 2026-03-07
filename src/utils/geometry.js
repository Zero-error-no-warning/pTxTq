import { clamp, ensureArray } from "./object.js";

const ANGLE_TO_RAD = Math.PI / (180 * 60000);
const RAD_TO_ANGLE = (180 * 60000) / Math.PI;
const DEG_TO_RAD = Math.PI / 180;
const RAD_TO_DEG = 180 / Math.PI;
const OOXML_DEGREE = 60000;

function toNumber(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function toRadians(angle60000) {
  return toNumber(angle60000, 0) * ANGLE_TO_RAD;
}

function unresolvedNumber(strict) {
  return strict ? Number.NaN : 0;
}

function addFractionVars(target, prefix, value) {
  target[`${prefix}2`] = value / 2;
  target[`${prefix}3`] = value / 3;
  target[`${prefix}4`] = value / 4;
  target[`${prefix}5`] = value / 5;
  target[`${prefix}6`] = value / 6;
  target[`${prefix}8`] = value / 8;
  target[`${prefix}10`] = value / 10;
  target[`${prefix}12`] = value / 12;
  target[`${prefix}16`] = value / 16;
  target[`${prefix}32`] = value / 32;
}

function evalFormulaValue(index, args, vars, options) {
  return evalGeomToken(args[index], vars, options);
}

export function evalGeomToken(token, vars = {}, options = {}) {
  const strict = options.strict === true;
  if (token === undefined || token === null || token === "") {
    return unresolvedNumber(strict);
  }
  const str = String(token).trim();
  if (!str) {
    return unresolvedNumber(strict);
  }
  const numeric = Number(str);
  if (Number.isFinite(numeric)) {
    return numeric;
  }
  if (Object.prototype.hasOwnProperty.call(vars, str)) {
    return toNumber(vars[str], 0);
  }
  return unresolvedNumber(strict);
}

export function evalGeomFormula(fmla, vars = {}, options = {}) {
  const strict = options.strict === true;
  if (fmla === undefined || fmla === null) {
    return unresolvedNumber(strict);
  }
  const parts = String(fmla).trim().split(/\s+/).filter(Boolean);
  if (!parts.length) {
    return unresolvedNumber(strict);
  }
  if (parts.length === 1) {
    return evalGeomToken(parts[0], vars, options);
  }

  const [op, ...args] = parts;
  const v = (idx) => evalFormulaValue(idx, args, vars, options);
  const hasNaN = (...values) => values.some((value) => !Number.isFinite(value));

  switch (op) {
    case "val": {
      const value = v(0);
      return strict && hasNaN(value) ? Number.NaN : value;
    }
    case "*/": {
      const a = v(0);
      const b = v(1);
      const d = v(2);
      if (strict && hasNaN(a, b, d)) {
        return Number.NaN;
      }
      return d ? (a * b) / d : 0;
    }
    case "+-": {
      const a = v(0);
      const b = v(1);
      const c = v(2);
      return strict && hasNaN(a, b, c) ? Number.NaN : a + b - c;
    }
    case "+/": {
      const a = v(0);
      const b = v(1);
      const d = v(2);
      if (strict && hasNaN(a, b, d)) {
        return Number.NaN;
      }
      return d ? (a + b) / d : 0;
    }
    case "?:": {
      const a = v(0);
      const b = v(1);
      const c = v(2);
      return strict && hasNaN(a, b, c) ? Number.NaN : (a > 0 ? b : c);
    }
    case "abs": {
      const value = v(0);
      return strict && hasNaN(value) ? Number.NaN : Math.abs(value);
    }
    case "sqrt": {
      const value = v(0);
      return strict && hasNaN(value) ? Number.NaN : Math.sqrt(Math.max(0, value));
    }
    case "max": {
      const a = v(0);
      const b = v(1);
      return strict && hasNaN(a, b) ? Number.NaN : Math.max(a, b);
    }
    case "min": {
      const a = v(0);
      const b = v(1);
      return strict && hasNaN(a, b) ? Number.NaN : Math.min(a, b);
    }
    case "mod": {
      const a = v(0);
      const b = v(1);
      const c = v(2);
      return strict && hasNaN(a, b, c) ? Number.NaN : Math.sqrt(a * a + b * b + c * c);
    }
    case "sin": {
      const a = v(0);
      const b = v(1);
      return strict && hasNaN(a, b) ? Number.NaN : a * Math.sin(toRadians(b));
    }
    case "cos": {
      const a = v(0);
      const b = v(1);
      return strict && hasNaN(a, b) ? Number.NaN : a * Math.cos(toRadians(b));
    }
    case "tan": {
      const a = v(0);
      const b = v(1);
      return strict && hasNaN(a, b) ? Number.NaN : a * Math.tan(toRadians(b));
    }
    case "at2": {
      const a = v(0);
      const b = v(1);
      return strict && hasNaN(a, b) ? Number.NaN : Math.atan2(b, a) * RAD_TO_ANGLE;
    }
    case "cat2": {
      const a = v(0);
      const b = v(1);
      const c = v(2);
      if (strict && hasNaN(a, b, c)) {
        return Number.NaN;
      }
      const angle = Math.atan2(c, b);
      return a * Math.cos(angle);
    }
    case "sat2": {
      const a = v(0);
      const b = v(1);
      const c = v(2);
      if (strict && hasNaN(a, b, c)) {
        return Number.NaN;
      }
      const angle = Math.atan2(c, b);
      return a * Math.sin(angle);
    }
    case "pin": {
      const a = v(0);
      const b = v(1);
      const c = v(2);
      return strict && hasNaN(a, b, c) ? Number.NaN : clamp(b, a, c);
    }
    default:
      return evalGeomToken(parts[0], vars, options);
  }
}

export function resolveGeometryGuides(guides, vars = {}) {
  const resolved = { ...vars };
  let pending = ensureArray(guides).filter((gd) => gd?.name);
  let pass = 0;

  while (pending.length && pass < Math.max(4, pending.length * 2)) {
    const nextPending = [];
    let progressed = false;

    for (const gd of pending) {
      const value = evalGeomFormula(gd.fmla || "0", resolved, { strict: true });
      if (Number.isFinite(value)) {
        resolved[gd.name] = value;
        progressed = true;
      } else {
        nextPending.push(gd);
      }
    }

    if (!progressed) {
      pending = nextPending;
      break;
    }

    pending = nextPending;
    pass += 1;
  }

  for (const gd of pending) {
    resolved[gd.name] = evalGeomFormula(gd.fmla || "0", resolved);
  }

  return resolved;
}

export function buildGeometryVars(geometry, pathW, pathH) {
  const w = Math.max(1, toNumber(pathW, 21600));
  const h = Math.max(1, toNumber(pathH, 21600));
  const ss = Math.min(w, h);
  const ls = Math.max(w, h);

  const vars = {
    w,
    h,
    l: 0,
    t: 0,
    r: w,
    b: h,
    hc: w / 2,
    vc: h / 2,
    ssd2: ss / 2,
    ssd4: ss / 4,
    ssd6: ss / 6,
    ssd8: ss / 8,
    ssd16: ss / 16,
    ssd32: ss / 32,
    cd2: 180 * OOXML_DEGREE,
    cd4: 90 * OOXML_DEGREE,
    cd8: 45 * OOXML_DEGREE,
    "3cd4": 270 * OOXML_DEGREE,
    "3cd8": 135 * OOXML_DEGREE,
    "5cd8": 225 * OOXML_DEGREE,
    "7cd8": 315 * OOXML_DEGREE,
    ss,
    ls
  };

  addFractionVars(vars, "wd", w);
  addFractionVars(vars, "hd", h);
  addFractionVars(vars, "ssd", ss);

  const withGuides = resolveGeometryGuides(geometry?.guideValues, vars);
  return resolveGeometryGuides(geometry?.adjustValues, withGuides);
}

export function convertOoxmlToAwtAngleDeg(ooAngleDeg, width, height) {
  const safeWidth = Math.abs(toNumber(width, 0));
  const safeHeight = Math.abs(toNumber(height, 0));
  if (safeWidth <= 0.000001 || safeHeight <= 0.000001) {
    return -ooAngleDeg;
  }

  const aspect = safeHeight / safeWidth;
  let awtAngle = -toNumber(ooAngleDeg, 0);

  let awtAngle2 = awtAngle % 360;
  let awtAngle3 = awtAngle - awtAngle2;

  switch (Math.trunc(awtAngle2 / 90)) {
    case -3:
      awtAngle3 -= 360;
      awtAngle2 += 360;
      break;
    case -2:
    case -1:
      awtAngle3 -= 180;
      awtAngle2 += 180;
      break;
    case 2:
    case 1:
      awtAngle3 += 180;
      awtAngle2 -= 180;
      break;
    case 3:
      awtAngle3 += 360;
      awtAngle2 -= 360;
      break;
    default:
      break;
  }

  awtAngle = (Math.atan2(Math.tan(awtAngle2 * DEG_TO_RAD), aspect) * RAD_TO_DEG) + awtAngle3;
  return awtAngle;
}

export function resolveOoxmlArcFromCurrentPoint(currentX, currentY, rx, ry, stAng, swAng) {
  const safeRx = Math.abs(toNumber(rx, 0));
  const safeRy = Math.abs(toNumber(ry, 0));
  if (safeRx <= 0.000001 || safeRy <= 0.000001) {
    return null;
  }

  const startDeg = toNumber(stAng, 0) / OOXML_DEGREE;
  const endDeg = startDeg + (toNumber(swAng, 0) / OOXML_DEGREE);

  const awtStart = convertOoxmlToAwtAngleDeg(startDeg, safeRx, safeRy);
  const awtEnd = convertOoxmlToAwtAngleDeg(endDeg, safeRx, safeRy);
  const awtSweep = awtEnd - awtStart;

  // Canvas/SVG use y-down coordinates where increasing parameter angles sweep clockwise.
  const startParam = -awtStart * DEG_TO_RAD;
  const sweepParam = -awtSweep * DEG_TO_RAD;

  const startRad = startDeg * DEG_TO_RAD;
  const invStart = Math.atan2(safeRx * Math.sin(startRad), safeRy * Math.cos(startRad));
  const centerX = toNumber(currentX, 0) - safeRx * Math.cos(invStart);
  const centerY = toNumber(currentY, 0) - safeRy * Math.sin(invStart);

  return {
    cx: centerX,
    cy: centerY,
    rx: safeRx,
    ry: safeRy,
    startParam,
    sweepParam,
    endParam: startParam + sweepParam
  };
}

export function splitArcSweep(sweepParam, maxSegmentAbs = Math.PI * 1.5) {
  const sweep = toNumber(sweepParam, 0);
  const maxAbs = Math.max(0.01, Math.abs(toNumber(maxSegmentAbs, Math.PI * 1.5)));
  if (!Number.isFinite(sweep) || Math.abs(sweep) < 0.0000001) {
    return [];
  }

  const chunks = [];
  const direction = sweep > 0 ? 1 : -1;
  let remaining = sweep;

  while (Math.abs(remaining) > maxAbs) {
    chunks.push(direction * maxAbs);
    remaining -= direction * maxAbs;
  }

  if (Math.abs(remaining) > 0.0000001) {
    chunks.push(remaining);
  }

  return chunks;
}
