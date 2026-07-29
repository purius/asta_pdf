import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const viewer = await readFile(
  new URL("../src/PdfMergeTool/Assets/PdfViewerOfficial/web/viewer.mjs", import.meta.url),
  "utf8"
);

assert.match(viewer, /#selectThumbnailPages\(pageNumber, event\)/);
assert.match(viewer, /if \(target instanceof HTMLInputElement\) \{\s*this\.\#selectThumbnailPages\(pageNumber, e\);/s);
assert.match(viewer, /if \(e\.shiftKey \|\| e\.ctrlKey \|\| e\.metaKey\) \{\s*e\.preventDefault\(\);\s*this\.\#selectThumbnailPages\(pageNumber, e\);/s);
assert.match(viewer, /this\.\#goToPage\(e\);/);

const adapter = await readFile(
  new URL("../src/PdfMergeTool/Assets/PdfViewerOfficial/web/app-adapter.js", import.meta.url),
  "utf8"
);

assert.match(adapter, /function undoPageEdit\(\)/);
assert.match(adapter, /document\.getElementById\("viewsManagerStatusUndoButton"\)/);
assert.match(adapter, /case "undo":\s*if \(!undoPageEdit\(\) && !window\.EditorAdapter\?\.undo\?\.\(\)\)/s);
