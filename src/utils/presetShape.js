import { resolveGeometryGuides } from "./geometry.js";
import { clamp, ensureArray } from "./object.js";

export function presetAdjustValue(element, name, fallbackValue, min = 0, max = 100000) {
  const gd = ensureArray(element?.geometry?.adjustValues).find((entry) => (
    String(entry?.name || "").toLowerCase() === String(name || "").toLowerCase()
  ));
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

function trianglePoints(box, element) {
  const adj = presetAdjustValue(element, "adj", 50000) / 100000;
  return [
    { x: box.x + box.cx * adj, y: box.y },
    { x: box.x + box.cx, y: box.y + box.cy },
    { x: box.x, y: box.y + box.cy }
  ];
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

function insetOctagonPoints(box, insetRatio) {
  const inset = Math.min(box.cx, box.cy) * insetRatio;
  return [
    { x: box.x + inset, y: box.y },
    { x: box.x + box.cx - inset, y: box.y },
    { x: box.x + box.cx, y: box.y + inset },
    { x: box.x + box.cx, y: box.y + box.cy - inset },
    { x: box.x + box.cx - inset, y: box.y + box.cy },
    { x: box.x + inset, y: box.y + box.cy },
    { x: box.x, y: box.y + box.cy - inset },
    { x: box.x, y: box.y + inset }
  ];
}

function reversePoints(points) {
  return [...ensureArray(points)].reverse().map((point) => ({ ...point }));
}

function ellipsePoints(cx, cy, rx, ry, startDeg = -90, sweepDeg = 360, steps = 48) {
  const points = [];
  const count = Math.max(8, steps);
  for (let i = 0; i <= count; i += 1) {
    const angle = ((startDeg + (sweepDeg * i) / count) * Math.PI) / 180;
    points.push({
      x: cx + Math.cos(angle) * rx,
      y: cy + Math.sin(angle) * ry
    });
  }
  return points;
}

function ringLoops(box, innerRatio = 0.55) {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  const outer = ellipsePoints(cx, cy, box.cx / 2, box.cy / 2);
  const inner = reversePoints(
    ellipsePoints(cx, cy, (box.cx / 2) * innerRatio, (box.cy / 2) * innerRatio)
  );
  return [outer, inner];
}

function frameLoops(box, insetRatio = 0.18) {
  const insetX = box.cx * insetRatio;
  const insetY = box.cy * insetRatio;
  return [
    pointsFromRatios(box, [[0, 0], [1, 0], [1, 1], [0, 1]]),
    reversePoints(pointsFromRatios(box, [
      [insetX / box.cx, insetY / box.cy],
      [1 - insetX / box.cx, insetY / box.cy],
      [1 - insetX / box.cx, 1 - insetY / box.cy],
      [insetX / box.cx, 1 - insetY / box.cy]
    ]))
  ];
}

function ringSegmentPoints(box, startDeg, sweepDeg, innerRatio = 0.58, steps = 40) {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  const outer = ellipsePoints(cx, cy, box.cx / 2, box.cy / 2, startDeg, sweepDeg, steps);
  const inner = reversePoints(
    ellipsePoints(cx, cy, (box.cx / 2) * innerRatio, (box.cy / 2) * innerRatio, startDeg, sweepDeg, steps)
  );
  return [...outer, ...inner];
}

function pieSlicePoints(box, startDeg = -70, sweepDeg = 300, steps = 40) {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  return [
    { x: cx, y: cy },
    ...ellipsePoints(cx, cy, box.cx / 2, box.cy / 2, startDeg, sweepDeg, steps)
  ];
}

function heartPoints(box, steps = 64) {
  const points = [];
  for (let i = 0; i < steps; i += 1) {
    const t = (Math.PI * 2 * i) / steps;
    const x = 16 * Math.pow(Math.sin(t), 3);
    const y = 13 * Math.cos(t) - 5 * Math.cos(2 * t) - 2 * Math.cos(3 * t) - Math.cos(4 * t);
    points.push({
      x: box.x + box.cx * (0.5 + x / 36),
      y: box.y + box.cy * (0.56 - y / 34)
    });
  }
  return points;
}

function cloudPoints(box, bumps = 7, steps = 84) {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy * 0.54;
  const rx = box.cx / 2;
  const ry = box.cy * 0.42;
  const points = [];
  for (let i = 0; i < steps; i += 1) {
    const a = (Math.PI * 2 * i) / steps;
    const radius = 0.84 + 0.12 * Math.sin(a * bumps) + 0.06 * Math.sin(a * (bumps + 3));
    points.push({
      x: cx + Math.cos(a) * rx * radius,
      y: cy + Math.sin(a) * ry * radius
    });
  }
  return points;
}

function gearPoints(box, teeth = 6) {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  const outer = Math.min(box.cx, box.cy) * 0.5;
  const inner = outer * 0.72;
  const valley = outer * 0.54;
  const points = [];
  for (let i = 0; i < teeth * 2; i += 1) {
    const angle = (-Math.PI / 2) + (i * Math.PI) / teeth;
    const radius = i % 2 === 0 ? outer : valley;
    points.push({
      x: cx + Math.cos(angle) * radius,
      y: cy + Math.sin(angle) * radius
    });
    const midAngle = angle + Math.PI / (teeth * 2);
    points.push({
      x: cx + Math.cos(midAngle) * inner,
      y: cy + Math.sin(midAngle) * inner
    });
  }
  return points;
}

function wavePolygonPoints(box, lobes = 2, thicknessRatio = 0.34, steps = 48) {
  const top = [];
  const bottom = [];
  const amplitude = box.cy * 0.16;
  const centerY = box.y + box.cy * 0.5;
  const halfThickness = box.cy * thicknessRatio * 0.5;
  for (let i = 0; i <= steps; i += 1) {
    const ratio = i / steps;
    const phase = ratio * Math.PI * 2 * lobes;
    const y = centerY + Math.sin(phase) * amplitude;
    const x = box.x + box.cx * ratio;
    top.push({ x, y: y - halfThickness });
    bottom.push({ x, y: y + halfThickness });
  }
  return [...top, ...reversePoints(bottom)];
}

function teardropPoints(box, steps = 56) {
  const points = [];
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy * 0.58;
  const rx = box.cx * 0.34;
  const ry = box.cy * 0.34;
  points.push({ x: box.x + box.cx / 2, y: box.y });
  for (let i = 0; i <= steps; i += 1) {
    const angle = ((20 + (320 * i) / steps) * Math.PI) / 180;
    points.push({
      x: cx + Math.cos(angle) * rx,
      y: cy + Math.sin(angle) * ry
    });
  }
  return points;
}

function crescentLoops(box) {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  const outer = ellipsePoints(cx, cy, box.cx / 2, box.cy / 2);
  const inner = reversePoints(
    ellipsePoints(
      box.x + box.cx * 0.62,
      cy,
      box.cx * 0.32,
      box.cy * 0.42
    )
  );
  return [outer, inner];
}

function mathBarRect(box, yRatio) {
  const barHeight = Math.max(box.cy * 0.16, 1);
  const y = box.y + box.cy * yRatio - barHeight / 2;
  return {
    kind: "rect",
    x: box.x + box.cx * 0.08,
    y,
    w: box.cx * 0.84,
    h: barHeight
  };
}

function roundRectPart(box, radiusRatio = 0.08) {
  return {
    kind: "roundRect",
    x: box.x,
    y: box.y,
    w: box.cx,
    h: box.cy,
    r: Math.min(box.cx, box.cy) * radiusRatio
  };
}

function pillPart(box) {
  return {
    kind: "roundRect",
    x: box.x,
    y: box.y,
    w: box.cx,
    h: box.cy,
    r: Math.min(box.cx, box.cy) / 2
  };
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

function wedgeRectCalloutPoints(box, element) {
  const vars = resolvePresetGuideVars(
    box,
    PRESET_WEDGE_RECT_CALLOUT_ADJUSTS,
    PRESET_WEDGE_RECT_CALLOUT_GUIDES,
    element
  );
  return [
    { x: box.x + vars.l, y: box.y + vars.t },
    { x: box.x + vars.x1, y: box.y + vars.t },
    { x: box.x + vars.xt, y: box.y + vars.yt },
    { x: box.x + vars.x2, y: box.y + vars.t },
    { x: box.x + vars.r, y: box.y + vars.t },
    { x: box.x + vars.r, y: box.y + vars.y1 },
    { x: box.x + vars.xr, y: box.y + vars.yr },
    { x: box.x + vars.r, y: box.y + vars.y2 },
    { x: box.x + vars.r, y: box.y + vars.b },
    { x: box.x + vars.x2, y: box.y + vars.b },
    { x: box.x + vars.xb, y: box.y + vars.yb },
    { x: box.x + vars.x1, y: box.y + vars.b },
    { x: box.x + vars.l, y: box.y + vars.b },
    { x: box.x + vars.l, y: box.y + vars.y2 },
    { x: box.x + vars.xl, y: box.y + vars.yl },
    { x: box.x + vars.l, y: box.y + vars.y1 }
  ];
}

function normalizePresetType(shapeType) {
  const normalized = String(shapeType || "rect").toLowerCase();
  const aliases = {
    cross: "plus",
    mathplus: "plus",
    actionbuttonbackprevious: "roundrect",
    actionbuttonbeginning: "roundrect",
    actionbuttonblank: "roundrect",
    actionbuttondocument: "roundrect",
    actionbuttonend: "roundrect",
    actionbuttonforwardnext: "roundrect",
    actionbuttonhelp: "roundrect",
    actionbuttonhome: "roundrect",
    actionbuttoninformation: "roundrect",
    actionbuttonmovie: "roundrect",
    actionbuttonreturn: "roundrect",
    actionbuttonsound: "roundrect",
    chartplus: "plus",
    chartstar: "star5",
    chartx: "mathmultiply",
    bentarrow: "rightarrow",
    bentuparrow: "uparrow",
    chord: "ellipse",
    circulararrow: "blockarc",
    curveddownarrow: "downarrow",
    curvedleftarrow: "leftarrow",
    curvedrightarrow: "rightarrow",
    curveduparrow: "uparrow",
    ellipseribbon: "ribbon",
    ellipseribbon2: "ribbon2",
    flowchartprocess: "rect",
    flowchartpredefinedprocess: "rect",
    flowchartinternalstorage: "rect",
    flowchartmultidocument: "rect",
    flowchartdocument: "rect",
    flowchartdecision: "diamond",
    flowchartdata: "parallelogram",
    flowchartinputoutput: "parallelogram",
    flowchartpreparation: "hexagon",
    flowchartconnector: "ellipse",
    flowchartoffpageconnector: "homeplate",
    flowchartalternateprocess: "roundrect",
    flowchartstoreddata: "can",
    flowchartonlinestorage: "can",
    flowchartofflinestorage: "can",
    flowchartmagneticdrum: "can",
    flowchartmagneticdisk: "can",
    flowchartterminator: "pill",
    flowchartmanualinput: "manualinput",
    flowchartor: "ellipse",
    flowchartsort: "diamond",
    flowchartsummingjunction: "ellipse",
    irregularseal1: "star12",
    irregularseal2: "star16",
    leftarrowcallout: "leftarrow",
    leftbrace: "leftbrace",
    leftbracket: "leftbracket",
    leftcirculararrow: "blockarc",
    leftrightcirculararrow: "blockarc",
    leftrightribbon: "leftrightribbon",
    leftrightuparrow: "quadarrow",
    leftuparrow: "leftarrow",
    rightarrowcallout: "rightarrow",
    rightbrace: "rightbrace",
    rightbracket: "rightbracket",
    uparrowcallout: "uparrow",
    downarrowcallout: "downarrow",
    leftrightarrowcallout: "leftrightarrow",
    updownarrowcallout: "updownarrow",
    quadarrowcallout: "quadarrow",
    bordercallout1: "wedgerectcallout",
    bordercallout2: "wedgerectcallout",
    bordercallout3: "wedgerectcallout",
    accentcallout1: "wedgerectcallout",
    accentcallout2: "wedgerectcallout",
    accentcallout3: "wedgerectcallout",
    accentbordercallout1: "wedgerectcallout",
    accentbordercallout2: "wedgerectcallout",
    accentbordercallout3: "wedgerectcallout",
    callout1: "wedgerectcallout",
    callout2: "wedgerectcallout",
    callout3: "wedgerectcallout",
    bracepair: "bracepair",
    bracketpair: "bracketpair",
    lineinv: "line",
    mathnotequal: "mathequal",
    moon: "moon",
    nonisoscelestrapezoid: "trapezoid",
    oval: "ellipse",
    piewedge: "triangle",
    smileyface: "ellipse",
    stripedrightarrow: "rightarrow",
    swoosharrow: "rightarrow",
    ribbon: "ribbon",
    ribbon2: "ribbon2",
    uturnarrow: "uparrow",
    wedgeellipsecallout: "ellipse",
    verticalscroll: "verticalscroll",
    horizontalscroll: "horizontalscroll",
    wedgeroundrectcallout: "roundrect",
    sun: "sun"
  };
  return aliases[normalized] || normalized;
}

export function buildPresetShapeParts(shapeType, box, element = null) {
  const normalized = normalizePresetType(shapeType);
  const regularPolygonSides = {
    heptagon: 7,
    octagon: 8,
    decagon: 10,
    dodecagon: 12
  };
  if (regularPolygonSides[normalized]) {
    return [{ kind: "polygon", points: regularPolygonPoints(box, regularPolygonSides[normalized]) }];
  }

  const starMatch = normalized.match(/^star(\d+)$/);
  if (starMatch) {
    const starPoints = Number.parseInt(starMatch[1], 10);
    const innerRatio = presetAdjustValue(element, "adj", 38000, 5000, 95000) / 100000;
    return [{ kind: "polygon", points: starPolygonPoints(box, starPoints, innerRatio) }];
  }

  switch (normalized) {
    case "rect":
      return [{ kind: "rect", x: box.x, y: box.y, w: box.cx, h: box.cy }];
    case "ellipse":
      return [{
        kind: "ellipse",
        cx: box.x + box.cx / 2,
        cy: box.y + box.cy / 2,
        rx: box.cx / 2,
        ry: box.cy / 2
      }];
    case "roundrect":
      return [roundRectPart(box)];
    case "pill":
      return [pillPart(box)];
    case "line":
      return [{ kind: "polyline", points: [
        { x: box.x, y: box.y + box.cy },
        { x: box.x + box.cx, y: box.y }
      ] }];
    case "triangle":
      return [{ kind: "polygon", points: trianglePoints(box, element) }];
    case "rttriangle":
      return [{ kind: "polygon", points: [
        { x: box.x, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy }
      ] }];
    case "diamond":
      return [{ kind: "polygon", points: [
        { x: box.x + box.cx / 2, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy / 2 },
        { x: box.x + box.cx / 2, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy / 2 }
      ] }];
    case "parallelogram":
      return [{ kind: "polygon", points: [
        { x: box.x + box.cx * 0.2, y: box.y },
        { x: box.x + box.cx, y: box.y },
        { x: box.x + box.cx * 0.8, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy }
      ] }];
    case "trapezoid":
      return [{ kind: "polygon", points: [
        { x: box.x + box.cx * 0.18, y: box.y },
        { x: box.x + box.cx * 0.82, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy }
      ] }];
    case "manualinput":
      return [{ kind: "polygon", points: [
        { x: box.x + box.cx * 0.2, y: box.y },
        { x: box.x + box.cx, y: box.y },
        { x: box.x + box.cx * 0.8, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy }
      ] }];
    case "cube":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.2, 0],
        [0.78, 0],
        [1, 0.22],
        [1, 0.78],
        [0.42, 0.78],
        [0.2, 1],
        [0, 0.78],
        [0, 0.22]
      ]) }];
    case "can":
      return [
        {
          kind: "rect",
          x: box.x + box.cx * 0.08,
          y: box.y + box.cy * 0.12,
          w: box.cx * 0.84,
          h: box.cy * 0.76
        },
        {
          kind: "ellipse",
          cx: box.x + box.cx / 2,
          cy: box.y + box.cy * 0.12,
          rx: box.cx * 0.42,
          ry: box.cy * 0.12
        },
        {
          kind: "ellipse",
          cx: box.x + box.cx / 2,
          cy: box.y + box.cy * 0.88,
          rx: box.cx * 0.42,
          ry: box.cy * 0.12
        }
      ];
    case "pentagon":
      return [{ kind: "polygon", points: [
        { x: box.x + box.cx * 0.5, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy * 0.38 },
        { x: box.x + box.cx * 0.82, y: box.y + box.cy },
        { x: box.x + box.cx * 0.18, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy * 0.38 }
      ] }];
    case "hexagon":
      return [{ kind: "polygon", points: [
        { x: box.x + box.cx * 0.25, y: box.y },
        { x: box.x + box.cx * 0.75, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy * 0.5 },
        { x: box.x + box.cx * 0.75, y: box.y + box.cy },
        { x: box.x + box.cx * 0.25, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy * 0.5 }
      ] }];
    case "chevron":
      return [{ kind: "polygon", points: [
        { x: box.x, y: box.y },
        { x: box.x + box.cx * 0.6, y: box.y },
        { x: box.x + box.cx, y: box.y + box.cy * 0.5 },
        { x: box.x + box.cx * 0.6, y: box.y + box.cy },
        { x: box.x, y: box.y + box.cy },
        { x: box.x + box.cx * 0.4, y: box.y + box.cy * 0.5 }
      ] }];
    case "plus":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
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
      ]) }];
    case "homeplate":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0],
        [0.72, 0],
        [1, 0.5],
        [0.72, 1],
        [0, 1]
      ]) }];
    case "rightarrow":
      return [{ kind: "polygon", points: blockArrowPoints(box, "right") }];
    case "leftarrow":
      return [{ kind: "polygon", points: blockArrowPoints(box, "left") }];
    case "uparrow":
      return [{ kind: "polygon", points: blockArrowPoints(box, "up") }];
    case "downarrow":
      return [{ kind: "polygon", points: blockArrowPoints(box, "down") }];
    case "leftrightarrow":
      return [{ kind: "polygon", points: doubleArrowPoints(box, "horizontal") }];
    case "updownarrow":
      return [{ kind: "polygon", points: doubleArrowPoints(box, "vertical") }];
    case "quadarrow":
      return [{ kind: "polygon", points: quadArrowPoints(box) }];
    case "notchedrightarrow":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0.28],
        [0.62, 0.28],
        [0.62, 0],
        [1, 0.5],
        [0.62, 1],
        [0.62, 0.72],
        [0.16, 0.72],
        [0, 0.5]
      ]) }];
    case "wedgerectcallout":
      return [{ kind: "polygon", points: wedgeRectCalloutPoints(box, element) }];
    case "ribbon":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0.18],
        [0.16, 0.18],
        [0.28, 0],
        [0.72, 0],
        [0.84, 0.18],
        [1, 0.18],
        [0.9, 0.5],
        [1, 0.82],
        [0.84, 0.82],
        [0.72, 1],
        [0.28, 1],
        [0.16, 0.82],
        [0, 0.82],
        [0.1, 0.5]
      ]) }];
    case "ribbon2":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.18, 0],
        [0.82, 0],
        [0.82, 0.18],
        [1, 0.3],
        [1, 0.7],
        [0.82, 0.82],
        [0.82, 1],
        [0.18, 1],
        [0.18, 0.82],
        [0, 0.7],
        [0, 0.3],
        [0.18, 0.18]
      ]) }];
    case "leftrightribbon":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0.5],
        [0.12, 0.18],
        [0.32, 0.18],
        [0.44, 0],
        [0.56, 0],
        [0.68, 0.18],
        [0.88, 0.18],
        [1, 0.5],
        [0.88, 0.82],
        [0.68, 0.82],
        [0.56, 1],
        [0.44, 1],
        [0.32, 0.82],
        [0.12, 0.82]
      ]) }];
    case "corner":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0],
        [1, 0],
        [1, 0.18],
        [0.24, 0.18],
        [0.24, 1],
        [0, 1]
      ]) }];
    case "cornertabs":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0],
        [0.86, 0],
        [1, 0.16],
        [1, 0.34],
        [0.72, 0.34],
        [0.72, 1],
        [0, 1]
      ]) }];
    case "diagstripe":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.14, 0],
        [0.42, 0],
        [0.86, 1],
        [0.58, 1]
      ]) }];
    case "bevel":
      return [{ kind: "polygon", points: insetOctagonPoints(box, 0.18) }];
    case "plaque":
      return [{ kind: "polygon", points: insetOctagonPoints(box, 0.1) }];
    case "plaquetabs":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.08, 0],
        [0.92, 0],
        [1, 0.12],
        [1, 0.34],
        [0.92, 0.46],
        [1, 0.58],
        [1, 0.88],
        [0.92, 1],
        [0.08, 1],
        [0, 0.88],
        [0, 0.58],
        [0.08, 0.46],
        [0, 0.34],
        [0, 0.12]
      ]) }];
    case "squaretabs":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.14, 0],
        [0.86, 0],
        [0.86, 0.18],
        [1, 0.18],
        [1, 0.82],
        [0.86, 0.82],
        [0.86, 1],
        [0.14, 1],
        [0.14, 0.82],
        [0, 0.82],
        [0, 0.18],
        [0.14, 0.18]
      ]) }];
    case "foldedcorner":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0],
        [0.76, 0],
        [1, 0.24],
        [1, 1],
        [0, 1]
      ]) }];
    case "frame":
      return [{ kind: "loops", loops: frameLoops(box, 0.18) }];
    case "halfframe":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0],
        [1, 0],
        [1, 0.18],
        [0.22, 0.18],
        [0.22, 1],
        [0, 1]
      ]) }];
    case "funnel":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0],
        [1, 0],
        [0.64, 0.42],
        [0.64, 1],
        [0.36, 1],
        [0.36, 0.42]
      ]) }];
    case "gear6":
      return [{ kind: "polygon", points: gearPoints(box, 6) }];
    case "gear9":
      return [{ kind: "polygon", points: gearPoints(box, 9) }];
    case "heart":
      return [{ kind: "polygon", points: heartPoints(box) }];
    case "horizontalscroll":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.1, 0],
        [0.9, 0],
        [1, 0.18],
        [0.9, 0.34],
        [1, 0.5],
        [0.9, 0.66],
        [1, 0.82],
        [0.9, 1],
        [0.1, 1],
        [0, 0.82],
        [0.1, 0.66],
        [0, 0.5],
        [0.1, 0.34],
        [0, 0.18]
      ]) }];
    case "leftbracket":
      return [{ kind: "polyline", points: pointsFromRatios(box, [
        [0.82, 0],
        [0.28, 0],
        [0.28, 1],
        [0.82, 1]
      ]) }];
    case "rightbracket":
      return [{ kind: "polyline", points: pointsFromRatios(box, [
        [0.18, 0],
        [0.72, 0],
        [0.72, 1],
        [0.18, 1]
      ]) }];
    case "leftbrace":
      return [{ kind: "polyline", points: pointsFromRatios(box, [
        [0.78, 0],
        [0.44, 0.14],
        [0.44, 0.38],
        [0.22, 0.5],
        [0.44, 0.62],
        [0.44, 0.86],
        [0.78, 1]
      ]) }];
    case "rightbrace":
      return [{ kind: "polyline", points: pointsFromRatios(box, [
        [0.22, 0],
        [0.56, 0.14],
        [0.56, 0.38],
        [0.78, 0.5],
        [0.56, 0.62],
        [0.56, 0.86],
        [0.22, 1]
      ]) }];
    case "bracketpair":
      return [
        { kind: "polyline", points: pointsFromRatios(box, [[0.24, 0], [0.08, 0], [0.08, 1], [0.24, 1]]) },
        { kind: "polyline", points: pointsFromRatios(box, [[0.76, 0], [0.92, 0], [0.92, 1], [0.76, 1]]) }
      ];
    case "bracepair":
      return [
        { kind: "polyline", points: pointsFromRatios(box, [[0.24, 0], [0.1, 0.18], [0.1, 0.38], [0.02, 0.5], [0.1, 0.62], [0.1, 0.82], [0.24, 1]]) },
        { kind: "polyline", points: pointsFromRatios(box, [[0.76, 0], [0.9, 0.18], [0.9, 0.38], [0.98, 0.5], [0.9, 0.62], [0.9, 0.82], [0.76, 1]]) }
      ];
    case "verticalscroll":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0.1],
        [0.18, 0],
        [0.34, 0.1],
        [0.5, 0],
        [0.66, 0.1],
        [0.82, 0],
        [1, 0.1],
        [1, 0.9],
        [0.82, 1],
        [0.66, 0.9],
        [0.5, 1],
        [0.34, 0.9],
        [0.18, 1],
        [0, 0.9]
      ]) }];
    case "mathminus":
      return [mathBarRect(box, 0.5)];
    case "mathequal":
      return [
        mathBarRect(box, 0.34),
        mathBarRect(box, 0.66)
      ];
    case "mathdivide":
      return [
        mathBarRect(box, 0.5),
        {
          kind: "ellipse",
          cx: box.x + box.cx / 2,
          cy: box.y + box.cy * 0.22,
          rx: Math.max(1, box.cx * 0.07),
          ry: Math.max(1, box.cy * 0.07)
        },
        {
          kind: "ellipse",
          cx: box.x + box.cx / 2,
          cy: box.y + box.cy * 0.78,
          rx: Math.max(1, box.cx * 0.07),
          ry: Math.max(1, box.cy * 0.07)
        }
      ];
    case "mathmultiply":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.18, 0],
        [0.5, 0.32],
        [0.82, 0],
        [1, 0.18],
        [0.68, 0.5],
        [1, 0.82],
        [0.82, 1],
        [0.5, 0.68],
        [0.18, 1],
        [0, 0.82],
        [0.32, 0.5],
        [0, 0.18]
      ]) }];
    case "arc":
      return [{ kind: "polyline", points: ellipsePoints(
        box.x + box.cx / 2,
        box.y + box.cy / 2,
        box.cx / 2,
        box.cy / 2,
        30,
        300,
        32
      ) }];
    case "blockarc":
      return [{ kind: "polygon", points: ringSegmentPoints(box, 25, 300, 0.62) }];
    case "donut":
      return [{ kind: "loops", loops: ringLoops(box, 0.55) }];
    case "pie":
      return [{ kind: "polygon", points: pieSlicePoints(box, -65, 290) }];
    case "cloud":
      return [{ kind: "polygon", points: cloudPoints(box) }];
    case "cloudcallout":
      return [{ kind: "polygon", points: [
        ...cloudPoints(box, 7, 72),
        { x: box.x + box.cx * 0.34, y: box.y + box.cy * 0.92 },
        { x: box.x + box.cx * 0.18, y: box.y + box.cy * 1.08 },
        { x: box.x + box.cx * 0.42, y: box.y + box.cy * 0.84 }
      ] }];
    case "flowchartpunchedcard":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.18, 0],
        [1, 0],
        [0.82, 1],
        [0, 1]
      ]) }];
    case "flowchartcollate":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0],
        [1, 0],
        [0.54, 0.5],
        [1, 1],
        [0, 1],
        [0.46, 0.5]
      ]) }];
    case "flowchartdelay":
      return [pillPart(box)];
    case "flowchartdisplay":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.16, 0],
        [1, 0],
        [0.82, 1],
        [0, 1],
        [0.16, 0.5]
      ]) }];
    case "flowchartextract":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.5, 0],
        [1, 1],
        [0, 1]
      ]) }];
    case "flowchartmanualoperation":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0],
        [1, 0],
        [0.8, 1],
        [0.2, 1]
      ]) }];
    case "flowchartmerge":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0, 0],
        [1, 0],
        [0.5, 1]
      ]) }];
    case "flowchartpunchedtape":
      return [{ kind: "polygon", points: wavePolygonPoints(box, 2, 0.88, 24) }];
    case "flowchartmagnettape":
      return [{ kind: "polygon", points: wavePolygonPoints(box, 2, 0.72, 24) }];
    case "flowchartmagnetictape":
      return [{ kind: "polygon", points: wavePolygonPoints(box, 2, 0.72, 24) }];
    case "flowchartonlinestorage":
      return buildPresetShapeParts("can", box, element);
    case "flowchartsummingjunction":
      return [{
        kind: "ellipse",
        cx: box.x + box.cx / 2,
        cy: box.y + box.cy / 2,
        rx: box.cx / 2,
        ry: box.cy / 2
      }];
    case "lightningbolt":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.42, 0],
        [0.7, 0],
        [0.52, 0.38],
        [0.8, 0.38],
        [0.34, 1],
        [0.46, 0.58],
        [0.2, 0.58]
      ]) }];
    case "moon":
      return [{ kind: "loops", loops: crescentLoops(box) }];
    case "nosmoking":
      return [{
        kind: "ellipse",
        cx: box.x + box.cx / 2,
        cy: box.y + box.cy / 2,
        rx: box.cx / 2,
        ry: box.cy / 2
      }];
    case "round1rect":
      return [roundRectPart(box, 0.16)];
    case "round2diagrect":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.1, 0],
        [1, 0],
        [1, 0.9],
        [0.9, 1],
        [0, 1],
        [0, 0.1]
      ]) }];
    case "round2samerect":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.14, 0],
        [1, 0],
        [1, 0.86],
        [0.86, 1],
        [0, 1],
        [0, 0.14]
      ]) }];
    case "snip1rect":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.12, 0],
        [1, 0],
        [1, 1],
        [0, 1],
        [0, 0.12]
      ]) }];
    case "snip2diagrect":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.12, 0],
        [1, 0],
        [1, 0.88],
        [0.88, 1],
        [0, 1],
        [0, 0.12]
      ]) }];
    case "snip2samerect":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.12, 0],
        [0.88, 0],
        [1, 0.12],
        [1, 1],
        [0.12, 1],
        [0, 0.88],
        [0, 0.12]
      ]) }];
    case "sniproundrect":
      return [{ kind: "polygon", points: pointsFromRatios(box, [
        [0.16, 0],
        [1, 0],
        [1, 0.84],
        [0.84, 1],
        [0.12, 1],
        [0, 0.88],
        [0, 0.16]
      ]) }];
    case "teardrop":
      return [{ kind: "polygon", points: teardropPoints(box) }];
    case "wave":
      return [{ kind: "polygon", points: wavePolygonPoints(box, 2, 0.36, 48) }];
    case "doublewave":
      return [{ kind: "polygon", points: wavePolygonPoints(box, 3, 0.28, 64) }];
    case "sun":
      return [{ kind: "polygon", points: starPolygonPoints(box, 16, 0.68) }];
    default:
      return null;
  }
}
