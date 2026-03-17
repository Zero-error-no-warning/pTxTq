import { ensureArray } from "./object.js";
import { presetShapeGeometry } from "./generatedPresetShapeGeometry.js";

function lowerCaseShapeType(shapeType) {
  return String(shapeType || "").trim().toLowerCase();
}

const POWERPOINT_DEFAULT_ADJUST_OVERRIDES = {
  leftrightarrow: [
    { name: "adj1", fmla: "val 70000" },
    { name: "adj2", fmla: "val 30000" }
  ],
  curveduparrow: [
    { name: "adj1", fmla: "val 30000" },
    { name: "adj2", fmla: "val 40000" },
    { name: "adj3", fmla: "val 25000" }
  ],
  curveddownarrow: [
    { name: "adj1", fmla: "val 30000" },
    { name: "adj2", fmla: "val 40000" },
    { name: "adj3", fmla: "val 25000" }
  ],
  curvedleftarrow: [
    { name: "adj1", fmla: "val 30000" },
    { name: "adj2", fmla: "val 40000" },
    { name: "adj3", fmla: "val 25000" }
  ],
  curvedrightarrow: [
    { name: "adj1", fmla: "val 30000" },
    { name: "adj2", fmla: "val 40000" },
    { name: "adj3", fmla: "val 25000" }
  ],
  halfframe: [
    { name: "adj1", fmla: "val 45000" },
    { name: "adj2", fmla: "val 45000" }
  ]
};

function mergeAdjustValues(defaults, overrides) {
  const overrideMap = new Map(
    ensureArray(overrides)
      .filter((entry) => entry?.name)
      .map((entry) => [String(entry.name).toLowerCase(), entry])
  );

  const merged = ensureArray(defaults).map((entry) => (
    overrideMap.get(String(entry?.name || "").toLowerCase()) || entry
  ));

  for (const entry of ensureArray(overrides)) {
    const key = String(entry?.name || "").toLowerCase();
    if (!key || merged.some((defaultEntry) => String(defaultEntry?.name || "").toLowerCase() === key)) {
      continue;
    }
    merged.push(entry);
  }

  return merged;
}

function resolveDefaultAdjustValues(shapeType, defaults, overrides) {
  if (ensureArray(overrides).length) {
    return defaults;
  }
  return POWERPOINT_DEFAULT_ADJUST_OVERRIDES[lowerCaseShapeType(shapeType)] || defaults;
}

export function hasPresetShapeGeometry(shapeType) {
  return Boolean(presetShapeGeometry[lowerCaseShapeType(shapeType)]);
}

export function resolvePresetShapeGeometry(shapeType, geometry = null) {
  const definition = presetShapeGeometry[lowerCaseShapeType(shapeType)];
  if (!definition) {
    return null;
  }

  return {
    kind: "cust",
    preset: geometry?.preset || definition.preset || shapeType || "rect",
    adjustValues: mergeAdjustValues(
      resolveDefaultAdjustValues(shapeType, definition.adjustValues, geometry?.adjustValues),
      geometry?.adjustValues
    ),
    guideValues: definition.guideValues,
    pathDefaults: definition.pathDefaults,
    paths: definition.paths,
    textRect: geometry?.textRect || definition.textRect || null,
    raw: null
  };
}
