# PDF.js Official Viewer and Editor Implementation Plan

Status: shipped in v1.0.93.

The app has moved from the legacy custom PDF viewer to the official PDF.js viewer. WPF/WebView2 remains the shell, `Assets\PdfViewerOfficial` is the active viewer bundle, and `MainWindow.xaml.cs` routes WebView2 to `https://pdfviewer.local/web/viewer.html`.

Legacy custom viewer removed. The old `Assets\PdfViewer` folder is no longer packaged or kept as rollback code. `scripts\verify-viewer-thumbnails.ps1` enforces this by failing if the old asset folder or old project content include returns.

## Completed Scope

1. Official PDF.js viewer bundle is packaged under `src\PdfMergeTool\Assets\PdfViewerOfficial`.
2. `app-adapter.js` bridges WPF commands to PDF.js viewer APIs.
3. Page navigation, zoom, page order, selection, rotate, delete, drag/drop insertion, and save state are preserved through the adapter and WPF host contract.
4. Thumbnail/page navigation instability is guarded by explicit navigation tracking and stale `pagechanging` suppression.
5. Editor overlays are available through `editor-adapter.js` with text, visual text replacement, whiteout, redaction, underline, strikeout, shapes, line, arrow, pen, highlight, image, stamp, signature, copy/paste/duplicate, layer order, opacity, and rotation controls.
6. `pdf-lib-adapter.js` persists app-owned overlay edits into saved PDFs, including embedded Windows fonts for Korean text through `WindowsFontService`.
7. Release builds run viewer/stability verification and JavaScript syntax checks before packaging.
8. Release publication verifies the GitHub latest release and `PdfMergeToolSetup.exe` download URL with retries.

## Verification Gates

Run these before shipping viewer/editor changes:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\verify-viewer-thumbnails.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\verify-pdf-lib-overlay-export.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\verify-stability.ps1
node --check src\PdfMergeTool\Assets\PdfViewerOfficial\web\app-adapter.js
node --check src\PdfMergeTool\Assets\PdfViewerOfficial\web\editor-adapter.js
node --check src\PdfMergeTool\Assets\PdfViewerOfficial\web\pdf-lib-adapter.js
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-installer.ps1
```

After a tag release, verify the updater-facing release state:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\verify-latest-release.ps1 -ExpectedVersion vX.Y.Z -RetryCount 90 -RetryDelaySeconds 10
```

## Manual QA Still Recommended

Automated checks cover contracts and packaging, but visual interaction should still be sampled before major releases:

1. Open a 30-page PDF and confirm every thumbnail appears.
2. Click thumbnails rapidly across distant pages and verify the main page does not jump backward.
3. Navigate next/previous repeatedly and verify thumbnail selection remains synchronized.
4. Rotate, delete, reorder, save, and reopen a multi-page PDF.
5. Add Korean text with a Windows font, save, reopen, and confirm the text is embedded visually.
6. Add whiteout/redaction/shape/image/signature overlays, save, and reopen.
