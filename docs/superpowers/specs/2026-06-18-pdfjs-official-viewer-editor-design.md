# PDF.js Official Viewer and Editor Design

## Goal

Replace the custom PDF viewing surface with the official PDF.js viewer so page navigation, thumbnail scrolling, zooming, and search use the maintained viewer implementation. Add the editor foundation with PDF.js annotation tools first, then persist custom text, shapes, images, stamps, and Korean font text through pdf-lib.

## Non-Goals

- Do not attempt Acrobat Pro style in-place editing of existing PDF text objects in the first implementation.
- Do not remove the existing PDF merge, rotate, delete, reorder, print, update, and installer flows.
- Do not replace the WPF application shell.

## Architecture

The WPF app keeps WebView2 as the viewer host. A new `Assets/PdfViewerOfficial` folder contains an unmodified official PDF.js viewer bundle plus a small app adapter. The adapter bridges WebView2 messages to PDF.js viewer APIs and posts normalized events back to `MainWindow.xaml.cs`.

## Current State

The official PDF.js viewer is the shipped viewer. `MainWindow.xaml.cs` maps `Assets\PdfViewerOfficial` to the WebView2 virtual host and navigates to `/web/viewer.html`. The legacy custom viewer assets have been removed and are no longer packaged, so rollback now means reverting through Git rather than switching a runtime asset folder.

Editing is implemented as app-owned overlays on top of the official viewer. The overlay state is collected by WPF before save, then `pdf-lib-adapter.js` flattens text, replacement text, shapes, lines, arrows, ink/highlight, whiteout, redaction, images, stamps, signatures, opacity, rotation, and embedded Korean/Windows fonts into the saved PDF.

## Message Contract

The official viewer adapter must support these incoming message types from WPF:

- `loadPdf`: open a PDF from a virtual-host URL or base64 payload.
- `command`: route existing commands such as `nextPage`, `prevPage`, `firstPage`, `lastPage`, zoom, fit page, rotate selected pages, delete selected pages, undo, redo, and mark clean.
- `nativePageTransferDragOver`, `nativePageTransferDragLeave`, `nativePageTransferDrop`, `nativeFileDragOver`, `nativeFileDrop`: preserve existing drag/drop integration where feasible.

The adapter must post these outgoing message types:

- `viewerReady`
- `viewerFirstPageRendered`
- `pageOrderChanged`
- `activePageChanged`
- `diagnostic`
- existing fallback render messages only if fallback mode remains active for unsupported PDFs.

## Viewer Behavior

The official viewer handles normal page rendering, scroll sync, thumbnails, zooming, and search. The app adapter avoids reimplementing virtual scrolling or thumbnail rendering. Page reorder/delete/rotate operations are represented by app state and saved through the existing C# PDF processing pipeline.

## Editor Behavior

The editor exposes app-owned overlay tools where PDF.js built-ins are not enough:

- Text box with Windows/Korean font selection
- Visual replacement of selected PDF text
- Whiteout and black redaction for selected text or arbitrary regions
- Underline and strikeout markup
- Rectangle, ellipse, line, arrow
- Image/stamp/signature placement
- Freehand pen and highlight
- Move, resize, rotate, opacity, copy, paste, duplicate, delete, and layer order controls

Persisting Phase 2 edits uses pdf-lib. Korean text must embed the selected font file instead of relying on PDF base fonts.

## Testing

Verification must include:

- Existing PowerShell stability scripts.
- Build script.
- Installer build.
- A viewer asset check proving the official viewer bundle and adapter are packaged.
- A smoke check that the WebView2 host navigates to the official viewer entry point.
- A release check proving the latest GitHub release and `PdfMergeToolSetup.exe` download URL are available for the updater.

Manual QA should cover:

- Open a 30-page PDF.
- Click thumbnails rapidly across multiple pages.
- Navigate next/previous repeatedly.
- Zoom and fit page.
- Delete, rotate, reorder pages, save, reopen.
- Add Korean text with a Windows font, save, and reopen.
- Add whiteout/redaction/shape/image/signature overlays, save, and reopen.

## Rollout

The official viewer transition has shipped, and the legacy custom viewer assets have been removed. Release workflow gates now run viewer/stability checks and JavaScript syntax checks before packaging, then verify the published latest release with retries after creating the GitHub release.
