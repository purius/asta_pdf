# PDF Merge Tool Context

PDF Merge Tool handles local PDF merge and page-editing workflows on Windows.

## Language

**Split Output Folder**:
The folder created next to a source PDF to hold files produced by a split operation. It never overwrites an existing folder; repeated operations use a numbered sibling folder such as `document_split (2)`.
_Avoid_: split folder, output directory

**Split Batch**:
A consecutive group of up to N source pages written as one file by an every-N-pages split. The final batch is written even when it contains fewer than N pages.
_Avoid_: page range, split unit

**Parity Split**:
A split operation that writes the odd-numbered source pages and even-numbered source pages into separate PDFs. If a parity has no pages, its output PDF is omitted and the completion notice reports zero files for that parity.
_Avoid_: odd-even split, alternating split

**Multi-file Split**:
Applying one split command to several selected source PDFs. Each source file is processed independently; a failure is recorded for that source while later files continue to process.
_Avoid_: bulk split, batch split

**Split Interval**:
The positive integer N used by every-N-pages split. It defaults to 1, accepts values larger than the source page count, and rejects non-integers and values below 1.
_Avoid_: page count, split size

**Page Selection**:
The explicit set of page thumbnails targeted by a page operation. A plain click selects one page, Shift-click selects the inclusive range from the selection anchor, and Ctrl-click toggles one page without changing other selections.
_Avoid_: checkbox state, current page

**Page Move Group**:
The ordered set moved when a selected thumbnail is dragged. Dragging an unselected thumbnail moves only that page; dragging any member of a Page Selection moves every selected page together, preserving their relative order.
_Avoid_: checked pages, drag selection

**Working Save Target**:
The user-chosen PDF file that receives subsequent saves for the currently open editing session. The original source PDF remains unchanged unless the user explicitly chooses to overwrite it.
_Avoid_: current PDF, save path

**Edit History**:
The chronological sequence of user-visible edits for one editing session. The Page Organizer records structural page actions, including a multi-page move, as undoable steps; the existing overlay editor retains its own history when no page action is pending. A successful save establishes a new clean baseline for page actions.
_Avoid_: editor-only undo, page undo

**Recovery Snapshot**:
An internal, non-source copy of unsaved editing state used only to offer recovery after an interrupted session. It never replaces a Working Save Target or the original source PDF without a user decision.
_Avoid_: auto-save to original, backup PDF

**Page Organizer**:
The app-owned page management panel that is the sole authority for page selection, order, rotations, structural edits, and Edit History. The embedded PDF viewer renders the current document but does not own page-editing interactions.
_Avoid_: PDF.js thumbnail editor, viewer page manager

**Active Page Follow**:
When preview navigation changes the active page, the Page Organizer scrolls only when that page's thumbnail is outside its visible viewport. Rapid preview changes are coalesced so only the latest active page is followed after layout. It pauses during a Page Move Group drag or Document Mutation, then resumes on the next preview navigation after that work finishes. It never changes Page Selection, page order, or a currently visible thumbnail's position.
_Avoid_: forced centering, selection sync

**Page Organizer Keyboard Navigation**:
When Page Organizer has keyboard focus, each arrow key moves only the active page by one position in current document order. Page Up and Page Down move it by 10 positions, and Home and End move it to the first and last positions. Navigation never replaces Page Selection; copy, paste, cut, undo, redo, and delete operate on that retained selection. Other app surfaces retain their own keyboard behavior when Page Organizer does not have focus.
_Avoid_: grid-direction navigation, selection replacement

**Document Mutation**:
A long-running operation that produces or replaces PDF content, such as save, extract, split, insert, cut, or A4 conversion. Only one Document Mutation can run for the active PDF. While it runs, Page Organizer and overlay-edit input are temporarily unavailable; opening another PDF cancels the old mutation before it can publish a result.
_Avoid_: background page edit, concurrent save
