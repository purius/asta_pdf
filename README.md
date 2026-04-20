# PDF Merge Tool

Windows PDF viewer and merge utility MVP.

## Current Features

- Open and preview a PDF by default.
- Show page-by-page thumbnails on the left side of the viewer.
- Show the total page count at the top of the thumbnail area.
- Resize the thumbnail sidebar by dragging the splitter.
- Adjust thumbnail zoom and main viewer zoom independently.
- Display thumbnails in a multi-column grid when the sidebar is wide enough.
- Reorder pages in the current viewer by dragging thumbnails.
- Select multiple pages with Ctrl/Shift click and move them together.
- Undo/redo page edits.
- Delete, rotate, extract, split, and reverse pages without editing text.
- Fit one page to the current viewer height.
- Print the current PDF from the File menu or toolbar.
- Open file merge as a separate menu/window.
- Merge page ranges such as `1-3,5,8-z`.
- Interleave two PDFs for odd/even duplex scan workflows.
- Add PDF files manually or by dropping files into the merge window.
- Reorder files with buttons or drag and drop.
- Sort by filename ascending or descending.
- Merge selected PDFs with qpdf.
- Save as `{first-file-name}_통합.pdf` next to the first file by default.
- Standalone installer with app icon, shortcuts, PDF merge context-menu registration, and Windows uninstall entry.

The viewer uses WebView2 and local PDF.js assets. Windows 11 normally includes the WebView2 runtime already.

## Viewer Shortcuts

- `Ctrl+O`: Open PDF
- `Ctrl+M`: Open merge window
- `Ctrl+P`: Print
- `Ctrl+Z` / `Ctrl+Y`: Undo or redo page edits
- `Delete`: Delete selected pages
- `Ctrl+R`: Rotate selected pages clockwise
- `F5`: Reload current PDF
- `Ctrl++` / `Ctrl+-`: Main viewer zoom
- `Ctrl+0`: Main viewer default zoom
- `Ctrl+1`: Fit one page
- `Ctrl+Shift++` / `Ctrl+Shift+-`: Thumbnail zoom
- `Ctrl+Shift+0`: Thumbnail default zoom
- `Ctrl+A`: Select all page thumbnails
- `PageUp`, `ArrowUp`, `ArrowLeft`: Previous page
- `PageDown`, `ArrowDown`, `ArrowRight`: Next page
- `Home` / `End`: First or last page

## Build

```powershell
.\scripts\restore-tools.ps1
.\scripts\build.ps1
```

The scripts install development tools under `.tools` in this workspace. They do not modify the system PATH.

## Publish Executable

```powershell
.\scripts\publish.ps1
```

Output:

```text
dist\PdfMergeTool\PdfMergeTool.exe
```

This is a self-contained Windows x64 build and includes qpdf under `tools\qpdf`.

To create a zip package:

```powershell
.\scripts\package.ps1
```

To create a standalone installer:

```powershell
.\scripts\build-installer.ps1
```

Output:

```text
dist\PdfMergeToolSetup.exe
```

The installer performs a per-user install under `%LOCALAPPDATA%\Programs\PdfMergeTool`, creates Start Menu and Desktop shortcuts, registers the PDF merge context menu, and adds a Windows uninstall entry. It also supports quiet mode:

Release publishing keeps only Korean satellite resources to reduce package size.

```powershell
.\dist\PdfMergeToolSetup.exe --quiet
"$env:LOCALAPPDATA\Programs\PdfMergeTool\PdfMergeToolSetup.exe" --uninstall --quiet
```

## Run

```powershell
.\.tools\dotnet\dotnet.exe .\src\PdfMergeTool\bin\Debug\net8.0-windows\PdfMergeTool.dll
```

Or pass PDFs directly:

```powershell
.\.tools\dotnet\dotnet.exe .\src\PdfMergeTool\bin\Debug\net8.0-windows\PdfMergeTool.dll --merge "C:\Docs\a.pdf" "C:\Docs\b.pdf"
```

## Next Milestones

- Image overlay on selected pages.
