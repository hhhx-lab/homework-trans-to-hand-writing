export const assetUrl = (path) => {
  const base = process.env.BASE_URL || "/";
  const normalizedBase = base.endsWith("/") ? base : `${base}/`;
  const normalizedPath = String(path || "").replace(/^\//, "");
  return `${normalizedBase}${normalizedPath}`;
};
