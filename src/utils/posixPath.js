function normalizeInput(value) {
  return String(value || "").replace(/\\/g, "/");
}

function splitPath(value) {
  const normalized = normalizeInput(value);
  const absolute = normalized.startsWith("/");
  const segments = [];

  for (const segment of normalized.split("/")) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      if (segments.length && segments[segments.length - 1] !== "..") {
        segments.pop();
      } else if (!absolute) {
        segments.push("..");
      }
      continue;
    }
    segments.push(segment);
  }

  return { absolute, segments };
}

export function normalizePosixPath(value) {
  const { absolute, segments } = splitPath(value);
  if (!segments.length) {
    return absolute ? "/" : ".";
  }
  return `${absolute ? "/" : ""}${segments.join("/")}`;
}

export function joinPosixPath(...values) {
  const parts = values
    .map((value) => normalizeInput(value))
    .filter((value) => value.length > 0);

  if (!parts.length) {
    return ".";
  }

  return normalizePosixPath(parts.join("/"));
}

export function dirnamePosixPath(value) {
  const normalized = normalizePosixPath(value);
  if (normalized === "/" || normalized === ".") {
    return normalized;
  }

  const absolute = normalized.startsWith("/");
  const segments = normalized.replace(/^\/+/, "").split("/");
  segments.pop();

  if (!segments.length) {
    return absolute ? "/" : ".";
  }

  return `${absolute ? "/" : ""}${segments.join("/")}`;
}

export function basenamePosixPath(value) {
  const normalized = normalizePosixPath(value);
  if (normalized === "/" || normalized === ".") {
    return normalized;
  }

  const segments = normalized.replace(/^\/+/, "").split("/");
  return segments[segments.length - 1] || "";
}

export function extnamePosixPath(value) {
  const base = basenamePosixPath(value);
  const lastDot = base.lastIndexOf(".");
  if (lastDot <= 0) {
    return "";
  }
  return base.slice(lastDot);
}

export function relativePosixPath(from, to) {
  const fromNormalized = normalizePosixPath(from);
  const toNormalized = normalizePosixPath(to);

  const fromSegments = fromNormalized === "."
    ? []
    : fromNormalized.replace(/^\/+/, "").split("/").filter(Boolean);
  const toSegments = toNormalized === "."
    ? []
    : toNormalized.replace(/^\/+/, "").split("/").filter(Boolean);

  let common = 0;
  while (
    common < fromSegments.length
    && common < toSegments.length
    && fromSegments[common] === toSegments[common]
  ) {
    common += 1;
  }

  const result = [
    ...Array(Math.max(fromSegments.length - common, 0)).fill(".."),
    ...toSegments.slice(common)
  ];

  return result.join("/");
}
