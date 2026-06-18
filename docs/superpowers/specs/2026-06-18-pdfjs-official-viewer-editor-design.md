# PDF.js Official Viewer and Editor Design

## Goal

Replace the custom PDF viewing surface with the official PDF.js viewer so page navigation, thumbnail scrolling, zooming, and search use the maintained viewer implementation. Add the editor foundation with PDF.js annotation tools first, then persist custom text, shapes, images, stamps, and Korean font text through pdf-lib.

## Non-Goals

- Do not attempt Acrobat Pro style in-place editing of existing PDF text objects in the first implementation.
- Do not remove the existing PDF merge, rotate, delete, reorder, print, update, and installer flows.
- Do not replace the WPF application shell.

## Architecture

The WPF app keeps WebView2 as the viewer host. A new `Assets/PdfViewerOfficial` folder contains an unmodified official PDF.js viewer bundle plus a small app adapter. The adapter bridges WebView2 messages to PDF.js viewer APIs and posts normalized events back to `MainWindow.xaml.cs`.

The existing custom `Assets/PdfViewer/viewer.html` remains temporarily as a rollback reference while the official viewer is integrated. Once parity is verified, `MainWindow.xaml.cs` navigates to the official viewer entry point.

Editing is layered in two phases. Phase 1 enables PDF.js annotation editor tools and collects editor state. Phase 2 adds pdf-lib-based save/export helpers that flatten text, simple shapes, images, stamps, signatures, and embedded Korean fonts into the saved PDF.

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

The official viewer handles normal page rendering, scroll sync, thumbnails, zooming, and search. The app adapter must avoid reimplementing virtual scrolling or thumbnail rendering. Page reorder/delete/rotate operations are represented by app state and saved through the existing C# PDF processing pipeline until a more native PDF.js/pdf-lib path replaces it.

## Editor Behavior

Phase 1 exposes stable PDF.js annotation editor modes:

- Text annotation
- Ink/highlight where available
- Stamp/image where available

Phase 2 adds app-owned overlay tools where PDF.js built-ins are not enough:

- Text box with Windows/Korean font selection
- Rectangle, ellipse, line, arrow
- Image/stamp/signature placement
- Move, resize, delete selected overlay object

Persisting Phase 2 edits uses pdf-lib. Korean text must embed the selected font file instead of relying on PDF base fonts.

## Testing

Verification must include:

- Existing PowerShell stability scripts.
- Build script.
- Installer build.
- A viewer asset check proving the official viewer bundle and adapter are packaged.
- A smoke check that the WebView2 host navigates to the official viewer entry point.

Manual QA should cover:

- Open a 30-page PDF.
- Click thumbnails rapidly across multiple pages.
- Navigate next/previous repeatedly.
- Zoom and fit page.
- Delete, rotate, reorder pages, save, reopen.
- Add a Korean text annotation and save when editor persistence lands.

## Rollout

Ship the official viewer transition first without removing the old viewer assets. If the official path has a regression, the app can be pointed back to the old `viewer.html` while keeping the new assets in the repository. After one stable release, remove the old custom viewer code in a separate cleanup.
