# Page Organizer Follow And Keyboard Design

**Date:** 2026-07-29

## Goal

Keep the WPF Page Organizer useful while a user navigates the PDF preview, and make its existing page operations reliably reachable from the keyboard.

## Scope

- Follow the preview's active page in the Page Organizer only when the matching thumbnail leaves the visible viewport.
- Add Page Organizer-focused keyboard navigation and route existing page-operation shortcuts through that focused surface.
- Preserve the responsive multi-column thumbnail layout and the Page Organizer's ownership of selection, order, rotation, and history.

## Decisions

### Active Page Follow

- Preview navigation remains the source of active-page updates. It continues to call the existing Page Organizer activation path without changing Page Selection.
- Rapid preview page changes are coalesced. One dispatcher callback follows only the newest active page after WPF completes layout.
- The callback compares the active item's `ListBoxItem` bounds with the Page Organizer `ScrollViewer` viewport. It changes only the vertical offset required to reveal an item that is above or below the viewport. A visible item does not move.
- Follow is suspended during a Page Move Group drag and while a Document Mutation is active. A stale callback from a previous document generation is ignored.
- Follow never creates edit history, changes page order, modifies rotations, or changes Page Selection.

### Keyboard Focus And Commands

- The Page Organizer `ListBox` accepts keyboard focus. Clicking a thumbnail or its checkbox moves focus to that list.
- Its preview key handler intercepts commands only while that list owns focus. Other app surfaces keep their current keyboard behavior.
- Arrow keys move only the active page by one position in current document order. Direction does not depend on the visual thumbnail grid.
- `PageUp` and `PageDown` move the active page by 10 positions and clamp at the first or last page. `Home` and `End` move to the first or last page.
- Keyboard navigation preserves Page Selection. It sends the resulting active page to the preview but does not add an undo-history entry.
- `Ctrl+C`, `Ctrl+X`, `Ctrl+V`, `Ctrl+Z`, `Ctrl+Y`, and `Delete` route to the same existing Page Organizer operations used by menus and toolbar buttons. Paste continues to prefer transferred page data and otherwise handles a clipboard image.
- This scope does not add Shift-arrow range selection or other new selection modifiers. Existing Shift/Ctrl mouse selection remains unchanged.

## Implementation Boundaries

- Keep follow orchestration in `MainWindow`; do not add structural state ownership to PDF.js.
- Reuse `NavigatePageOrganizer` and `NavigatePageOrganizerBoundary` for keyboard movement rather than duplicating page-index calculations.
- Add a small visual-tree helper only as needed to locate the Page Organizer's generated `ScrollViewer` and active item container.
- Queue follow work with the current document load generation so a delayed callback cannot scroll the newly opened document using an old page number.
- Do not change the existing responsive `WrapPanel`, drag/drop insertion indicators, or clipboard payload format.

## Verification

- Extend the Page Organizer contract test to require the focused keyboard handler, follow queue, generation guard, drag/mutation gate, and viewport-only scroll path.
- Keep the existing state and split tests unchanged; this feature must not alter page selection or history semantics.
- Windows manual smoke test with a multi-column, multi-page PDF:
  1. Scroll the main preview until the active thumbnail leaves and re-enters the visible Page Organizer viewport.
  2. Rapidly wheel through pages and confirm the sidebar settles on the final active page without repeated jumps.
  3. Start a thumbnail group drag and confirm no follow scroll moves the drop target.
  4. Focus Page Organizer and exercise arrows, PageUp/PageDown, Home/End, copy, cut, paste, undo, redo, and delete.
  5. Confirm the selected group is retained through keyboard navigation and each structural command remains undoable.
