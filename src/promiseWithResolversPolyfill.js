// pdfjs-dist calls Promise.withResolvers() (ES2024), which some browsers
// don't have yet. This file has no imports of its own, so importing it
// first — before pdfjs-dist — guarantees this polyfill runs before
// pdfjs-dist's module code does. (A polyfill written inline in the same
// file as the pdfjs-dist import does NOT reliably run first: ES modules
// fully evaluate all of a file's imports, in the order they're written,
// before running any of that file's own top-level statements — so
// pdfjs-dist would still execute before an inline polyfill below it.)
if (typeof Promise.withResolvers !== "function") {
  Promise.withResolvers = function () {
    let resolve, reject;
    const promise = new Promise((res, rej) => { resolve = res; reject = rej; });
    return { promise, resolve, reject };
  };
}
