import { ensureArray } from "./object.js";
import { presetShapeGeometry } from "./generatedPresetShapeGeometry.js";

function lowerCaseShapeType(shapeType) {
  return String(shapeType || "").trim().toLowerCase();
}

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
    adjustValues: mergeAdjustValues(definition.adjustValues, geometry?.adjustValues),
    guideValues: definition.guideValues,
    pathDefaults: definition.pathDefaults,
    paths: definition.paths,
    textRect: geometry?.textRect || definition.textRect || null,
    raw: null
  };
}
