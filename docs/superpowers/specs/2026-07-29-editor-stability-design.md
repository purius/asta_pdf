# Editor Stability Design

**Date:** 2026-07-29

## Goal

Make page reordering and saving deterministic in the Windows PDF viewer. A user must be able to select one or more pages, move them, undo or redo the move, save the result, and reopen the saved file without the applied page state being lost or changed.

## Problem

The current viewer keeps overlapping state in the PDF.js adapter and `MainWindow.xaml.cs`. The WPF host caches page order, selection, rotations, and dirty state from asynchronous WebView2 messages, then asks the viewer for editor state again during save. PDF.js native thumbnail editing also emits page-order updates independently. These paths can observe different moments of the document and produce incorrect moves, selection-dependent behavior, or a saved file that does not reflect the visible edit.

## Scope

- Page selection, multi-page moves, page order, rotation, delete, insert, undo/redo, and dirty tracking.
- Save target lifecycle, atomic output publication, saved-result verification, close/open prompts, and interrupted-session recovery.
- Automated tests for the supported editing and saving contracts.

## Non-goals

- macOS Finder integration.
- Collaborative editing or concurrent editing of one PDF by multiple app instances.
- Silent overwrite of an original source PDF.
- General PDF content editing beyond preserving the existing overlay editor behavior.

## Canonical State

The WebView viewer owns one `EditorDocumentState` with a monotonically increasing `revision`:

- ordered page transforms: source page number and normalized rotation;
- `selectedPageNumbers` in document order;
- selection anchor for Shift selection;
- active page;
- overlay edits;
- dirty baseline revision and undo/redo history.

WPF sends named commands and receives an acknowledged snapshot after each state-changing command. WPF may display this snapshot, but must not construct a save from separately cached page and editor messages. A save request names the expected revision; the viewer responds with one immutable save snapshot for that revision. A mismatched or unavailable revision fails visibly rather than saving an uncertain state.

## Selection And Move Rules

- Plain thumbnail click replaces the Page Selection with the clicked page and sets the selection anchor.
- Shift-click selects the inclusive range between the selection anchor and clicked page in current document order.
- Ctrl-click toggles only the clicked page and updates the selection anchor to that page.
- Checkbox UI, if retained, is another input for Page Selection only; it never gates a move.
- Dragging an unselected thumbnail moves only that page.
- Dragging a selected thumbnail moves the entire Page Move Group: all selected pages in their current relative order.
- The insertion marker identifies the gap before or after a thumbnail. The implementation removes the move group before resolving its final insertion index, preventing off-by-one errors when moving forward.
- After a move, the moved pages stay selected and the active page is the page that was dragged. No-op drops do not create history entries.

## History

Every completed user operation creates one `Edit History` entry. This includes selection-independent page moves, delete, rotate, insert, and overlay editor changes. A multi-page move is one entry. Undo and redo restore the complete state snapshot and publish the restored revision. A successful save marks the current revision clean but does not discard the history needed for undo while the document remains open.

## Save Lifecycle

The first save establishes a Working Save Target through Save As. Later saves write to that same target. The original source remains unchanged unless a separate, explicit overwrite flow is selected.

For each save:

1. Freeze one acknowledged save snapshot at the current revision and disable concurrent save commands.
2. Assemble transformed pages and overlays into a temporary PDF next to the target when possible.
3. Verify the temporary file exists, can be opened, has the expected page count, and matches the frozen ordered page-transform plan.
4. Replace the Working Save Target atomically, retaining a recoverable previous file until replacement succeeds.
5. Reopen the published target from disk using the frozen revision as the expected clean baseline.
6. Mark the document clean only after reopening and verifying the target. On any failure, leave the editor dirty and retain the last valid target.

The UI reports the exact published file path and any failed verification reason. It never displays a successful save message before publication completes.

## Recovery And Lifecycle

Before file replacement, opening another PDF, and application close, persist a Recovery Snapshot for a dirty document. Opening another file or closing prompts `Save`, `Don't Save`, or `Cancel`. An interrupted session offers the snapshot on next startup; accepting recovery creates a working session and never overwrites either the source or Working Save Target automatically. Successful save, discard, and explicit recovery dismissal remove the snapshot.

## Verification

Tests must cover at least:

1. Plain, Shift, and Ctrl page selection in document order.
2. Single-page moves before and after a target page.
3. Multi-page moves in both directions, preserving relative order.
4. No-op drop behavior and selection preservation after a move.
5. Unified undo/redo across page and overlay operations.
6. Save of reordered and rotated pages, including the reopened result.
7. Save failure leaves the working document dirty and preserves the previous output.
8. First save target, subsequent save target reuse, explicit overwrite, close prompt, and recovery snapshot lifecycle.

Unit tests should cover pure state transitions and save planning. A Windows UI/integration test layer should exercise the WebView command/acknowledgement contract with a multi-page fixture PDF before release packaging.
