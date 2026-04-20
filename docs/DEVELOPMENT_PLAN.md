# PDF Merge/Edit Tool Development Plan

## Goal

Build a small Windows desktop utility for PDF workflows from File Explorer:

- Show a PDF-only right-click menu.
- Let the user choose merge order before merging.
- Save merged output next to the first input file as `{first-file-name}_통합.pdf`.
- Later add page-level editing and image overlay without full text editing.

## Product Scope

### MVP

- Open selected PDF files from the app or from Explorer command-line arguments.
- Show a reorderable file list.
- Support ascending and descending filename sort.
- Support moving files up/down and removing files.
- Merge files into one PDF.
- Generate a safe output name in the first file's folder.

### Next

- Register/unregister `.pdf` Explorer context menu.
- Add single-instance argument aggregation for multi-select Explorer launches.
- Add PDF preview using WebView2 + PDF.js.
- Add page thumbnails and page reorder/delete/rotate.

### Later

- Add image overlay on selected pages.
- Add stamps/watermarks.
- Add lightweight annotations.

### Out of Scope

- Existing text editing.
- OCR.
- Digital signature validation.
- PDF/A conformance tooling.

## Technical Choices

- UI: C# WPF on .NET 8.
- PDF merge/page operations: `qpdf`.
- PDF image overlay: PDFsharp in a later phase.
- Viewer: PDF.js hosted in WebView2 in a later phase.
- Explorer integration: registry verb first, `IExplorerCommand` shell extension later if needed.

## App Command Contract

```powershell
PdfMergeTool.exe --merge "C:\Docs\a.pdf" "C:\Docs\b.pdf"
```

The app should accept file paths with or without `--merge`. Non-PDF paths are ignored.

## qpdf Resolution

The app resolves qpdf in this order:

1. `tools\qpdf\qpdf.exe` next to the app.
2. `qpdf.exe` from `PATH`.

This keeps development flexible and allows an installer to bundle qpdf later.

## Explorer Context Menu Plan

Initial registry location:

```text
HKCU\Software\Classes\SystemFileAssociations\.pdf\shell\PdfMergeTool
```

Values:

```text
(Default) = PDF 통합...
MultiSelectModel = Player
```

Command:

```text
"C:\Path\To\PdfMergeTool.exe" --merge "%1"
```

Static verbs may still launch one process per file depending on Explorer behavior. The app will therefore add single-instance aggregation in the next phase.

## Risks

- `qpdf` may not be installed. The UI must show a clear setup error.
- Static registry verbs are not first-class Windows 11 context menu entries. A packaged `IExplorerCommand` extension is needed for the modern menu.
- Page/image editing needs coordinate mapping between preview pixels and PDF points.

## Milestones

1. MVP merge window and qpdf service.
2. Build and smoke test.
3. Context menu install/uninstall scripts.
4. Single-instance aggregation.
5. PDF preview and page thumbnails.
6. Page operations.
7. Image overlay.
