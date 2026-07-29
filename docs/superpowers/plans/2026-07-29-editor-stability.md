# Editor Stability Implementation Plan

> **For implementation:** execute in order, keeping each item independently testable. Do not release a build until the scenario tests and saved-PDF verification pass on Windows.

**Goal:** Replace asynchronous split state for page edits and saving with one acknowledged editor-document state, reliable multi-page moves, transactional saving, and recovery.

**Architecture:** The PDF.js adapter owns a revisioned `EditorDocumentState`. WPF acts as a command host and consumes acknowledged immutable snapshots. A dedicated C# save coordinator writes and validates a temporary PDF before publishing the working save target. Pure reducers and save planners receive unit tests; WebView integration covers the message contract.

## 1. Establish a Reproducible Fixture and Test Harness

**Files:**
- Create `test/PdfMergeTool.Tests/EditorDocumentStateTests.cs`
- Create `test/PdfMergeTool.Tests/SavePublicationTests.cs`
- Create `test/fixtures/editor-pages.pdf`
- Modify `test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj`
- Modify `.github/workflows/release.yml`

1. Add a deterministic multi-page fixture whose pages are distinguishable after extraction.
2. Add test helpers that read resulting page order and rotations with the packaged PDF tooling.
3. Add the new test project cases to the release workflow before installer construction.
4. Run the focused test project and confirm it fails before reducer/save implementation.

## 2. Introduce the Revisioned Editor Document State

**Files:**
- Create `src/PdfMergeTool/Services/EditorDocumentState.cs`
- Create `src/PdfMergeTool/Services/EditorDocumentReducer.cs`
- Modify `src/PdfMergeTool/Assets/PdfViewerOfficial/web/app-adapter.js`
- Modify `src/PdfMergeTool/MainWindow.xaml.cs`

1. Model ordered transforms, selection, anchor, active page, overlays, dirty baseline, revision, and history snapshots as one state.
2. Implement pure C# transition tests for selection, move, rotate, delete, insert, undo, redo, and clean-baseline behavior.
3. Mirror the same command and acknowledgement schema in the adapter; messages include command id, expected revision, resulting revision, and complete snapshot.
4. Replace WPF's independently mutated `_pageOrder`, `_selectedPages`, `_pageRotations`, `_activePage`, and `_isDirty` fields with the latest acknowledged snapshot.
5. Reject stale acknowledgements and surface command errors rather than silently accepting them.

## 3. Replace Thumbnail Reorder With Explicit Group Moves

**Files:**
- Modify `src/PdfMergeTool/Assets/PdfViewerOfficial/web/app-adapter.js`
- Modify `src/PdfMergeTool/Assets/PdfViewerOfficial/web/viewer.css`
- Modify `src/PdfMergeTool/MainWindow.xaml.cs`
- Extend `test/PdfMergeTool.Tests/EditorDocumentStateTests.cs`

1. Capture plain, Shift, and Ctrl thumbnail input under the canonical selection rules.
2. Restore custom thumbnail drag handling only as an explicit command producer, not as direct DOM or PDF.js page-manager mutation.
3. Compute the Page Move Group and insertion gap from the canonical state; remove the group before calculating its target index.
4. Render selection and insertion feedback from the snapshot, keeping checkbox state visual only.
5. Add forward/backward multi-page move, no-op, and selection-preservation tests.

## 4. Unify Page And Overlay History

**Files:**
- Modify `src/PdfMergeTool/Assets/PdfViewerOfficial/web/app-adapter.js`
- Modify `src/PdfMergeTool/Assets/PdfViewerOfficial/web/editor-adapter.js`
- Modify `src/PdfMergeTool/MainWindow.xaml.cs`
- Extend `test/PdfMergeTool.Tests/EditorDocumentStateTests.cs`

1. Route overlay changes through the same history transaction API as page commands.
2. Make `Ctrl+Z` and `Ctrl+Y` restore the complete document snapshot and publish one acknowledgement.
3. Disable undo/redo controls only from snapshot history capabilities.
4. Verify page move followed by overlay edit, then undo/redo through both operations.

## 5. Add Transactional Save Publication

**Files:**
- Create `src/PdfMergeTool/Services/PdfSaveCoordinator.cs`
- Create `src/PdfMergeTool/Services/PdfSavePlan.cs`
- Modify `src/PdfMergeTool/Services/PdfMergeService.cs`
- Modify `src/PdfMergeTool/MainWindow.xaml.cs`
- Modify `src/PdfMergeTool/Services/SavePathPromptService.cs`
- Extend `test/PdfMergeTool.Tests/SavePublicationTests.cs`

1. Track a Working Save Target separately from source and currently rendered PDF paths.
2. Freeze one acknowledged snapshot before save and map all overlay edits from its page order.
3. Write to a temporary output, verify it can open and matches expected transforms, then atomically publish it.
4. Reload the published result and mark clean only after validation succeeds.
5. Preserve the prior published target and dirty editor state on cancellation, generation failure, verification failure, or replacement failure.
6. Test first save, subsequent direct save, explicit overwrite, and failure rollback.

## 6. Add Recovery And File-Lifecycle Protection

**Files:**
- Create `src/PdfMergeTool/Services/RecoverySnapshotService.cs`
- Modify `src/PdfMergeTool/App.xaml.cs`
- Modify `src/PdfMergeTool/MainWindow.xaml.cs`
- Modify `src/PdfMergeTool/Services/AppPaths.cs`
- Extend `test/PdfMergeTool.Tests/SavePublicationTests.cs`

1. Serialize a dirty snapshot before close, open-file replacement, and save publication.
2. Add `Save`, `Don't Save`, and `Cancel` decisions for close and file replacement.
3. Offer valid interrupted-session recovery on startup, with explicit accept/discard behavior.
4. Remove snapshots only after successful save, explicit discard, or explicit recovery dismissal.
5. Test all snapshot lifecycle transitions without touching source PDFs.

## 7. Exercise Windows Integration and Release Gates

**Files:**
- Create `scripts/verify-editor-stability.ps1`
- Modify `scripts/verify-stability.ps1`
- Modify `.github/workflows/release.yml`
- Modify `README.md`

1. Add a Windows WebView integration test or deterministic host-contract harness for selection, group move, save, reopen, and recovery.
2. Add an installer/release gate that runs the focused tests and JavaScript syntax checks before packaging.
3. Document selection shortcuts, save-target behavior, recovery behavior, and the exact distinction between save and overwrite.
4. Build the Windows installer, install it in a clean test environment, and run the manual smoke sequence with the fixture PDF.

## Final Verification

1. Run all unit tests.
2. Run the new editor stability verifier and existing viewer/stability verifiers.
3. On Windows, perform: open fixture, Shift-select range, Ctrl-toggle, drag group forward and backward, undo/redo, save-as, modify and save, close/recover, reopen output, inspect resulting page order.
4. Confirm original source bytes remain unchanged unless the explicit overwrite action was chosen.
