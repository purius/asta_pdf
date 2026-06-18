Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$adapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\pdf-lib-adapter.js'

if (-not (Test-Path $adapterPath)) {
    throw 'pdf-lib adapter must exist before verifying overlay geometry.'
}

$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    throw 'node is required to verify overlay geometry.'
}

$script = @'
const fs = require("fs");
const vm = require("vm");

const sourcePath = process.argv[2];
let source = fs.readFileSync(sourcePath, "utf8");
source = source.replace(
  "  window.PdfLibAdapter = {",
  `  window.__overlayGeometryTest = {
    editRotation,
    editRotationDegrees,
    rotatePointAroundCenter,
    getBox,
    getRotatedBoxPoint,
    getRotatedBoxOrigin,
    getLineEndpoints
  };

  window.PdfLibAdapter = {`
);

const sandbox = {
  window: {
    PDFLib: {
      PDFDocument: {},
      degrees: value => value
    }
  },
  atob: value => Buffer.from(value, "base64").toString("binary"),
  btoa: value => Buffer.from(value, "binary").toString("base64")
};

vm.runInNewContext(source, sandbox, { filename: sourcePath });
const geometry = sandbox.window.__overlayGeometryTest;
if (!geometry) {
  throw new Error("overlay geometry helpers were not exported for verification.");
}

function assertAlmost(actual, expected, label) {
  if (Math.abs(actual - expected) > 0.001) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}

function assertPoint(actual, expected, label) {
  assertAlmost(actual.x, expected.x, `${label}.x`);
  assertAlmost(actual.y, expected.y, `${label}.y`);
}

const pageHeight = 200;
const line = geometry.getLineEndpoints({
  x: 100,
  y: 50,
  width: 80,
  height: 20,
  rotate: 90
}, pageHeight);
assertPoint(line.start, { x: 140, y: 180 }, "line.start");
assertPoint(line.end, { x: 140, y: 100 }, "line.end");

const boxEdit = {
  x: 100,
  y: 50,
  width: 80,
  height: 40,
  rotate: 90
};
assertPoint(
  geometry.getRotatedBoxOrigin(boxEdit, pageHeight),
  { x: 120, y: 170 },
  "rotatedBox.origin"
);
assertPoint(
  geometry.getRotatedBoxPoint(boxEdit, pageHeight, 40, 20),
  { x: 140, y: 130 },
  "rotatedBox.center"
);

assertAlmost(geometry.editRotation(sandbox.window.PDFLib, { rotate: 45 }), -45, "editRotation");
if (geometry.editRotation(sandbox.window.PDFLib, { rotate: 0 }) !== undefined) {
  throw new Error("zero edit rotation should be omitted.");
}

console.log("overlay geometry checks passed.");
'@

$tempScript = Join-Path ([System.IO.Path]::GetTempPath()) "verify-overlay-geometry-$([System.Guid]::NewGuid()).js"
try {
    Set-Content -LiteralPath $tempScript -Value $script -Encoding UTF8
    & $node.Source $tempScript $adapterPath
    if ($LASTEXITCODE -ne 0) {
        throw "overlay geometry checks failed with exit code $LASTEXITCODE."
    }
}
finally {
    if (Test-Path $tempScript) {
        Remove-Item -LiteralPath $tempScript -Force
    }
}
