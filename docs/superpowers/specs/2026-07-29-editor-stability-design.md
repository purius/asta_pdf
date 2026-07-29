# Page Organizer Stability Design

**Date:** 2026-07-29

## Goal

Make thumbnail selection, multi-page movement, undo/redo, and saving deterministic for the Windows PDF editor.

## Decision

The WPF application owns a single `EditorDocumentState`, presented to users as the **Page Organizer**. PDF.js is a preview surface only: it opens the source PDF, displays the active source page, and reports preview navigation. It does not own page selection, reordering, rotations, deletion, or structural undo/redo.

This removes the race between native PDF.js thumbnail state, the JavaScript adapter, and the WPF save host.

## State Contract

`EditorDocumentState` contains:

- source page numbers in current output order;
- explicit selected page numbers in current order;
- a stable Shift-selection anchor and active page;
- source-page rotation values;
- structural undo and redo snapshots;
- a clean baseline used for dirty-state calculation.

Only structural mutations create Page Organizer history entries. Plain selection, Shift range selection, Ctrl toggle, and preview navigation do not create dirty state.

## Interaction Contract

- Plain page click selects one page and establishes the range anchor.
- Shift-click selects the inclusive current-order range from the retained anchor.
- Ctrl-click and checkbox click toggle one page, including a reliable uncheck path.
- Dragging a selected page moves the selected group in its current relative order.
- Dragging an unselected page moves only that page.
- Delete, rotate, reverse, move, undo, and redo operate through `EditorDocumentState` only.
- External file and page drops are received by WPF and inserted after the current Page Organizer selection; they do not route through the preview adapter.
- The PDF.js split/merge page manager is disabled and hidden from the user.

## Preview And Save

The Page Organizer sends the active source page to the preview with `goToPage`. The preview may update only the active page when a user scrolls it; it never changes the selected group or document order.

Every document load has a monotonically increasing load ID. PDF.js serializes open requests, defers a new document's page navigation until that document is open, and tags preview/editor events with the active ID. WPF ignores messages from older loads so a late event from one PDF cannot mutate the next PDF's UI state.

Save reads ordered page transforms directly from Page Organizer state, assembles a temporary PDF, verifies its page count, publishes it to the Working Save Target, and reloads the published output as a clean document.

`DocumentOperationCoordinator` gives each document generation a cancellation token and serializes content mutations. While save, extract, split, copy/cut, insertion, or A4 conversion is active, the Page Organizer and overlay input are not interactive. Opening another PDF cancels the previous generation, so a late operation cannot publish or overwrite the new document's state.

A4 conversion uses the native PDFium renderer for both normal and fallback viewer modes. It receives the current Page Organizer order and rotations directly instead of waiting for a browser-side export message.

## Deliberate Boundaries

- Overlay object editing keeps its existing editor-local undo/redo when there is no pending Page Organizer history entry.
- Inserting externally generated pages rebuilds a new source PDF and begins a new Page Organizer history baseline.
- Recovery snapshots and a fully unified page-plus-overlay history remain separate future work; neither is required to make page management deterministic.

## Verification

- Unit tests cover Shift anchor retention, Ctrl/checkbox uncheck, grouped movement, rotation/delete restoration, and undo/redo.
- A JavaScript contract test confirms PDF.js structural editing is disabled, external structural drops stay in WPF, and stale asynchronous load events cannot replace the newest preview state.
- Release packaging runs the Page Organizer verifier, overlay geometry/export checks, adapter syntax checks, unit tests, build verification, and a packaged-app startup smoke test.
- A Windows manual smoke test remains required after installation: open, checkbox toggle, Shift range, Ctrl multi-select, drag group, undo/redo, save, reopen, and A4 conversion.
