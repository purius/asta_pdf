Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$pdfLibPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\vendor\pdf-lib.min.js'
$fontkitPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\vendor\fontkit.umd.min.js'
$adapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\pdf-lib-adapter.js'
$fontPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\standard_fonts\LiberationSans-Regular.ttf'

foreach ($path in @($pdfLibPath, $fontkitPath, $adapterPath, $fontPath)) {
    if (-not (Test-Path $path)) {
        throw "Required overlay export test asset is missing: $path"
    }
}

$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    throw 'node is required to verify pdf-lib overlay export.'
}

$script = @'
const fs = require("fs");
const vm = require("vm");

const [pdfLibPath, fontkitPath, adapterPath, fontPath] = process.argv.slice(2);
const fontBase64 = fs.readFileSync(fontPath).toString("base64");
const sandbox = {
  console,
  Buffer,
  Uint8Array,
  ArrayBuffer,
  TextEncoder,
  TextDecoder,
  setTimeout,
  clearTimeout,
  atob: value => Buffer.from(value, "base64").toString("binary"),
  btoa: value => Buffer.from(value, "binary").toString("base64")
};
sandbox.window = sandbox;
sandbox.self = sandbox;
sandbox.globalThis = sandbox;

vm.runInNewContext(fs.readFileSync(pdfLibPath, "utf8"), sandbox, { filename: pdfLibPath });
vm.runInNewContext(fs.readFileSync(fontkitPath, "utf8"), sandbox, { filename: fontkitPath });
vm.runInNewContext(fs.readFileSync(adapterPath, "utf8"), sandbox, { filename: adapterPath });
sandbox.fontBase64 = fontBase64;

const verification = `
(async () => {
  const pdfDoc = await PDFLib.PDFDocument.create();
  pdfDoc.addPage([420, 560]);
  const sourceBytes = await pdfDoc.save({ useObjectStreams: false });
  const noEditBase64 = await PdfLibAdapter.createOverlayPdf({ sourceBytes, edits: [] });
  const noEditLength = Buffer.from(noEditBase64, "base64").length;

  const edits = [
    {
      type: "text",
      page: 1,
      x: 32,
      y: 40,
      width: 180,
      height: 36,
      text: "한글 ABC",
      size: 16,
      fontName: "LiberationSans",
      color: [0.05, 0.05, 0.05],
      zIndex: 10
    },
    {
      type: "textReplace",
      page: 1,
      x: 32,
      y: 92,
      width: 150,
      height: 28,
      text: "Replacement",
      size: 12,
      fontName: "LiberationSans",
      whiteoutPadding: 2,
      lineHeight: 15,
      zIndex: 20
    },
    {
      type: "whiteout",
      variant: "redact",
      page: 1,
      x: 220,
      y: 40,
      width: 70,
      height: 24,
      zIndex: 30
    },
    {
      type: "rectangle",
      page: 1,
      x: 32,
      y: 150,
      width: 80,
      height: 48,
      borderWidth: 2,
      fillColor: [0.9, 0.95, 1],
      rotate: 10,
      zIndex: 40
    },
    {
      type: "rectangle",
      shape: "ellipse",
      page: 1,
      x: 130,
      y: 150,
      width: 80,
      height: 48,
      borderWidth: 2,
      zIndex: 50
    },
    {
      type: "line",
      page: 1,
      x: 32,
      y: 230,
      width: 140,
      height: 20,
      borderWidth: 3,
      rotate: -8,
      zIndex: 60
    },
    {
      type: "arrow",
      page: 1,
      x: 210,
      y: 230,
      width: 120,
      height: 20,
      borderWidth: 3,
      opacity: 0.8,
      zIndex: 70
    },
    {
      type: "ink",
      tool: "highlight",
      page: 1,
      points: [{ x: 34, y: 305 }, { x: 80, y: 312 }, { x: 140, y: 302 }],
      borderWidth: 10,
      zIndex: 80
    },
    {
      type: "image",
      page: 1,
      x: 300,
      y: 300,
      width: 32,
      height: 32,
      imageMimeType: "image/png",
      imageDataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
      zIndex: 90
    },
    {
      type: "stamp",
      page: 1,
      x: 32,
      y: 360,
      width: 110,
      height: 42,
      text: "확인",
      fontName: "LiberationSans",
      size: 14,
      zIndex: 100
    },
    {
      type: "text",
      page: 99,
      x: 0,
      y: 0,
      width: 20,
      height: 20,
      text: "stale page edit must be skipped",
      zIndex: 110
    }
  ];

  const editedBase64 = await PdfLibAdapter.createOverlayPdf({
    sourceBytes,
    fonts: { LiberationSans: fontBase64 },
    edits
  });
  const editedBytes = Uint8Array.from(Buffer.from(editedBase64, "base64"));
  const loaded = await PDFLib.PDFDocument.load(editedBytes);
  return {
    sourceLength: sourceBytes.length,
    noEditLength,
    editedLength: editedBytes.length,
    pageCount: loaded.getPageCount()
  };
})()
`;

vm.runInNewContext(verification, sandbox, { filename: "overlay-export-verification.js" })
  .then(result => {
    if (result.pageCount !== 1) {
      throw new Error(`Expected exported PDF to keep 1 page, got ${result.pageCount}.`);
    }
    if (result.noEditLength !== result.sourceLength) {
      throw new Error(`No-edit export changed the PDF length from ${result.sourceLength} to ${result.noEditLength}.`);
    }
    if (result.editedLength <= result.noEditLength + 2500) {
      throw new Error(`Edited export did not grow enough to prove overlays were embedded: ${JSON.stringify(result)}`);
    }
    console.log(JSON.stringify(result));
  })
  .catch(error => {
    console.error(error && error.stack ? error.stack : error);
    process.exit(1);
  });
'@

$tempScript = Join-Path ([System.IO.Path]::GetTempPath()) "verify-pdf-lib-overlay-export-$([System.Guid]::NewGuid()).js"
try {
    Set-Content -LiteralPath $tempScript -Value $script -Encoding UTF8
    $output = & $node.Source $tempScript $pdfLibPath $fontkitPath $adapterPath $fontPath
    if ($LASTEXITCODE -ne 0) {
        throw "pdf-lib overlay export verification failed with exit code $LASTEXITCODE."
    }

    $result = $output | Select-Object -Last 1 | ConvertFrom-Json
    if ($result.pageCount -ne 1) {
        throw "pdf-lib overlay export page-count verification failed: $($result.pageCount)"
    }

    Write-Output "pdf-lib overlay export checks passed. source=$($result.sourceLength) edited=$($result.editedLength)"
}
finally {
    if (Test-Path $tempScript) {
        Remove-Item -LiteralPath $tempScript -Force
    }
}
