export function ensureArray(value) {
  if (value === undefined || value === null) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

export function first(value, fallback = null) {
  if (value === undefined || value === null) {
    return fallback;
  }
  if (Array.isArray(value)) {
    return value.length ? value[0] : fallback;
  }
  return value;
}

export function toInt(value, fallback = 0) {
  if (value === undefined || value === null || value === "") {
    return fallback;
  }
  const parsed = Number.parseInt(String(value), 10);
  return Number.isNaN(parsed) ? fallback : parsed;
}

export function toFloat(value, fallback = 0) {
  if (value === undefined || value === null || value === "") {
    return fallback;
  }
  const parsed = Number.parseFloat(String(value));
  return Number.isNaN(parsed) ? fallback : parsed;
}

export function deepClone(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

export function pickDefined(obj) {
  return Object.fromEntries(Object.entries(obj).filter(([, value]) => value !== undefined));
}

export function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

export function assignIfDefined(target, key, value) {
  if (value !== undefined && value !== null) {
    target[key] = value;
  }
}
