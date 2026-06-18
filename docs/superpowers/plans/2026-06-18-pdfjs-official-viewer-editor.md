# PDF.js Official Viewer and Editor Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move the app from its custom PDF viewer to the official PDF.js viewer and establish an editor path based on PDF.js annotation tools plus pdf-lib persistence.

**Architecture:** Keep WPF/WebView2 as the shell. Add an official PDF.js viewer asset folder with a thin adapter script that translates the current app message contract into PDF.js viewer calls. Preserve the old custom viewer as rollback until official viewer parity is verified.

**Tech Stack:** WPF, WebView2, PDF.js official viewer, JavaScript adapter, existing C# PDF pipeline, future pdf-lib integration.

---

### Task 1: Preserve Current Stable Baseline

**Files:**
- Inspect: `src/PdfMergeTool/PdfMergeTool.csproj`
- Inspect: `src/PdfMergeTool/MainWindow.xaml.cs`

- [ ] **Step 1: Verify clean branch state**

Run: `git status --short --branch`
Expected: branch `codex/pdfjs-official-viewer` with only docs changes at this point.

- [ ] **Step 2: Run baseline verification**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\verify-viewer-thumbnails.ps1`
Expected: `viewer thumbnail rendering checks passed.`

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\verify-stability.ps1`
Expected: `stability checks passed.`

### Task 2: Add Official PDF.js Viewer Assets

**Files:**
- Create: `src/PdfMergeTool/Assets/PdfViewerOfficial/`
- Create: `src/PdfMergeTool/Assets/PdfViewerOfficial/app-adapter.js`
- Modify: `scripts/verify-viewer-thumbnails.ps1`

- [ ] **Step 1: Download official PDF.js generic viewer**

Use the npm `pdfjs-dist` package matching the current bundled PDF.js major version where possible. Copy the official `web` folder and required `build` files into `Assets/PdfViewerOfficial`.

- [ ] **Step 2: Add adapter script**

Create `app-adapter.js` that waits for PDF.js application initialization, posts `viewerReady`, receives WebView2 messages, and routes `loadPdf` plus basic navigation commands.

- [ ] **Step 3: Add asset verification**

Extend `scripts\verify-viewer-thumbnails.ps1` to assert:

```powershell
if (-not (Test-Path (Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\viewer.html'))) {
    throw 'official PDF.js viewer.html must be packaged.'
}

if (-not (Test-Path (Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\app-adapter.js'))) {
    throw 'official viewer adapter must be packaged.'
}
```

### Task 3: Switch WebView2 Navigation to Official Viewer

**Files:**
- Modify: `src/PdfMergeTool/MainWindow.xaml.cs`

- [ ] **Step 1: Route WebView2 to official entry point**

Change the viewer folder mapping to include `Assets\PdfViewerOfficial`, and navigate to an official wrapper entry point that loads PDF.js viewer plus `app-adapter.js`.

- [ ] **Step 2: Keep fallback support**

Keep existing fallback messages available. If official viewer cannot open a file, the C# fallback render path remains callable for future work.

- [ ] **Step 3: Verify host navigation**

Add or update a script check that `MainWindow.xaml.cs` references `PdfViewerOfficial` and no longer navigates directly to the old custom `viewer.html`.

### Task 4: Restore Core Viewer Commands

**Files:**
- Modify: `src/PdfMergeTool/Assets/PdfViewerOfficial/app-adapter.js`
- Modify: `src/PdfMergeTool/MainWindow.xaml.cs` only if the message shape needs a small compatibility adjustment.

- [ ] **Step 1: Implement page commands**

Support `nextPage`, `prevPage`, `firstPage`, `lastPage`.

- [ ] **Step 2: Implement zoom commands**

Support `mainZoomIn`, `mainZoomOut`, `mainZoomReset`, and `fitPage`.

- [ ] **Step 3: Post active page changes**

Listen to PDF.js page changing events and post `activePageChanged` and `pageOrderChanged`.

### Task 5: Preserve Page Operations

**Files:**
- Modify: `src/PdfMergeTool/Assets/PdfViewerOfficial/app-adapter.js`
- Modify: `src/PdfMergeTool/MainWindow.xaml.cs`

- [ ] **Step 1: Represent page order in adapter state**

Maintain `pageOrder`, `pageRotations`, and `selectedPages` in the adapter.

- [ ] **Step 2: Route rotate/delete/reorder commands**

Preserve current C# save/export semantics by posting normalized page order and rotation state back to the host.

- [ ] **Step 3: Verify save/reopen flow**

Run a manual smoke flow with a multi-page PDF: rotate, delete, save, reopen.

### Task 6: Add Editor Foundation

**Files:**
- Modify: `src/PdfMergeTool/Assets/PdfViewerOfficial/app-adapter.js`
- Create: `src/PdfMergeTool/Assets/PdfViewerOfficial/editor-adapter.js`

- [ ] **Step 1: Enable PDF.js annotation editor modes**

Expose commands for text, highlight/ink, and stamp/image where supported by the bundled PDF.js viewer.

- [ ] **Step 2: Add editor state export message**

Post editor changes to WPF as a bounded JSON state so the host can mark the document dirty.

### Task 7: Add pdf-lib Persistence Foundation

**Files:**
- Create: `src/PdfMergeTool/Assets/PdfViewerOfficial/pdf-lib-adapter.js`
- Add: `src/PdfMergeTool/Assets/PdfViewerOfficial/vendor/pdf-lib.min.js`

- [ ] **Step 1: Add pdf-lib asset**

Bundle `pdf-lib` locally so the app works offline.

- [ ] **Step 2: Implement Korean text embedding helper**

Load a selected font file supplied by WPF and use pdf-lib to embed it before drawing text.

- [ ] **Step 3: Implement first save path**

Persist app-owned text boxes and simple rectangles into a new PDF byte array.

### Task 8: Verify, Package, Release

**Files:**
- Modify: `src/PdfMergeTool/PdfMergeTool.csproj`

- [ ] **Step 1: Increment version**

Set the next release version after the current latest tag.

- [ ] **Step 2: Run verification**

Run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\verify-viewer-thumbnails.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\verify-stability.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-installer.ps1
```

- [ ] **Step 3: Commit and push**

Commit the official viewer transition and push the branch.

- [ ] **Step 4: Create GitHub release**

Create a new release with `PdfMergeToolSetup.exe` and verify the latest release download URL returns `200 OK`.
