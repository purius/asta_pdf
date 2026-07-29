# Page Organizer Stability Plan

**Goal:** Replace competing PDF.js/WPF page-edit state with an app-owned Page Organizer.

**Architecture:** `EditorDocumentState` is the canonical WPF state for page order, selection, rotations, structural history, and dirty state. The embedded viewer is preview-only.

## Work Items

- [x] Capture the failure modes: unreliable checkbox toggle, Shift range, group move, undo/redo, and save divergence.
- [x] Define Page Organizer terminology in `CONTEXT.md`.
- [x] Add pure state tests for retained Shift anchors, checkbox uncheck, multi-page movement, structural undo/redo, rotation, and deletion.
- [x] Move selection, page order, rotations, delete, reverse, and undo/redo into `EditorDocumentState`.
- [x] Add a WPF Page Organizer panel with rendered page previews, checkbox input, Shift/Ctrl selection, and group drag/drop.
- [x] Disable and hide PDF.js split/merge thumbnail editing; retain it only as a preview renderer.
- [x] Route save transforms and PDF navigation from Page Organizer state.
- [x] Route external page/file drops and load-scoped preview events directly through the current Page Organizer document.
- [x] Cancel stale document work, serialize content mutations, and lock page/editor input while a mutation is active.
- [x] Route A4 rendering through the native PDFium service instead of an unimplemented preview bridge.
- [x] Replace obsolete release verification with Page Organizer contract checks.
- [x] Add Windows CI startup smoke coverage for the packaged installer.
- [ ] Run the full manual Windows smoke sequence with a multi-page fixture after installation.

## Windows Smoke Sequence

1. Open a PDF with at least six distinguishable pages.
2. Checkbox-select and unselect a page.
3. Plain-click page 2, Shift-click page 5, then Ctrl-toggle page 3.
4. Drag the selected group before and after another page.
5. Use toolbar and `Ctrl+Z`/`Ctrl+Y` to undo and redo the move, rotation, and deletion.
6. Save to the Working Save Target, reopen the saved file, and compare its order to Page Organizer.
7. Open a second PDF and repeat selection and movement to confirm no state crosses document boundaries.
