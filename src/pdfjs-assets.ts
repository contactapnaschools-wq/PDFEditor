/** Absolute URL for a file shipped beside this library (worker, cmaps, fonts, wasm). */
export function pdfjsAsset(path: string): string {
  const base = import.meta.env.DEV
    ? new URL(`${import.meta.env.BASE_URL}pdfjs/`, window.location.origin)
    : new URL(libraryAssetBase(), import.meta.url);
  return new URL(path, base).href;
}

/** Kept as a function so Vite does not treat the relative path as a missing asset. */
function libraryAssetBase(): string {
  return './pdfjs/';
}
