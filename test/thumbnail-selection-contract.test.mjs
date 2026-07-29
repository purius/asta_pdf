import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const viewer = await readFile(
  new URL("../src/PdfMergeTool/Assets/PdfViewerOfficial/web/viewer.mjs", import.meta.url),
  "utf8"
);

assert.match(viewer, /#selectThumbnailPages\(pageNumber, event\)/);
assert.match(viewer, /target instanceof HTMLInputElement \|\| e\.shiftKey \|\| e\.ctrlKey \|\| e\.metaKey\) \{\s*e\.preventDefault\(\);\s*this\.\#selectThumbnailPages\(pageNumber, e\);/s);
assert.match(viewer, /this\.\#goToPage\(e\);/);
