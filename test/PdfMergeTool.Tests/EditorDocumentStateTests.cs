using PdfMergeTool.Services;
using Xunit;

namespace PdfMergeTool.Tests;

public sealed class EditorDocumentStateTests
{
    [Fact]
    public async Task DocumentOperationCoordinator_cancels_work_for_a_replaced_document()
    {
        using var coordinator = new DocumentOperationCoordinator();
        coordinator.StartNewDocument();
        var firstDocumentOperation = coordinator.Capture();

        coordinator.StartNewDocument();

        Assert.True(firstDocumentOperation.CancellationToken.IsCancellationRequested);
        Assert.False(coordinator.IsCurrent(firstDocumentOperation));
        await Assert.ThrowsAsync<OperationCanceledException>(async () =>
            await coordinator.EnterMutationAsync(firstDocumentOperation));
    }

    [Fact]
    public async Task DocumentOperationCoordinator_serializes_current_document_mutations()
    {
        using var coordinator = new DocumentOperationCoordinator();
        coordinator.StartNewDocument();
        var operation = coordinator.Capture();

        using var firstLease = await coordinator.EnterMutationAsync(operation);
        var secondLeaseTask = coordinator.EnterMutationAsync(operation);

        Assert.False(secondLeaseTask.IsCompleted);
        firstLease.Dispose();
        using var secondLease = await secondLeaseTask;
    }

    [Fact]
    public void SelectPage_uses_shift_for_inclusive_range_and_ctrl_for_toggle()
    {
        var state = EditorDocumentState.Create(6);

        state = state.SelectPage(2, PageSelectionMode.Replace);
        state = state.SelectPage(5, PageSelectionMode.Range);
        state = state.SelectPage(3, PageSelectionMode.Toggle);

        Assert.Equal(new[] { 2, 4, 5 }, state.SelectedPageNumbers);
        Assert.Equal(3, state.SelectionAnchorPageNumber);
        Assert.Equal(3, state.ActivePageNumber);
    }

    [Fact]
    public void SelectPage_keeps_the_original_shift_anchor_for_subsequent_ranges()
    {
        var state = EditorDocumentState.Create(6)
            .SelectPage(2, PageSelectionMode.Replace)
            .SelectPage(5, PageSelectionMode.Range)
            .SelectPage(4, PageSelectionMode.Range);

        Assert.Equal(new[] { 2, 3, 4 }, state.SelectedPageNumbers);
        Assert.Equal(2, state.SelectionAnchorPageNumber);
        Assert.Equal(4, state.ActivePageNumber);
    }

    [Fact]
    public void MovePageGroup_moves_selected_pages_together_and_preserves_relative_order()
    {
        var state = EditorDocumentState.Create(6)
            .SelectPage(2, PageSelectionMode.Replace)
            .SelectPage(4, PageSelectionMode.Toggle);

        var moved = state.MovePageGroup(2, insertionIndex: 6);

        Assert.Equal(new[] { 1, 3, 5, 6, 2, 4 }, moved.PageNumbers);
        Assert.Equal(new[] { 2, 4 }, moved.SelectedPageNumbers);
        Assert.Equal(2, moved.ActivePageNumber);
        Assert.True(moved.IsDirty);
    }

    [Fact]
    public void MovePageGroup_moves_only_dragged_page_when_it_is_not_selected()
    {
        var state = EditorDocumentState.Create(5)
            .SelectPage(2, PageSelectionMode.Replace)
            .SelectPage(3, PageSelectionMode.Toggle);

        var moved = state.MovePageGroup(5, insertionIndex: 0);

        Assert.Equal(new[] { 5, 1, 2, 3, 4 }, moved.PageNumbers);
        Assert.Equal(new[] { 5 }, moved.SelectedPageNumbers);
        Assert.Equal(5, moved.ActivePageNumber);
    }

    [Fact]
    public void MovePageGroup_does_not_mark_dirty_when_drop_keeps_same_order()
    {
        var state = EditorDocumentState.Create(4).SelectPage(2, PageSelectionMode.Replace);

        var moved = state.MovePageGroup(2, insertionIndex: 2);

        Assert.Equal(state.PageNumbers, moved.PageNumbers);
        Assert.False(moved.IsDirty);
    }

    [Fact]
    public void MovePageGroup_can_be_undone_and_redone()
    {
        var moved = EditorDocumentState.Create(5)
            .SelectPage(2, PageSelectionMode.Replace)
            .SelectPage(4, PageSelectionMode.Toggle)
            .MovePageGroup(2, insertionIndex: 5);

        var undone = moved.Undo();
        var redone = undone.Redo();

        Assert.Equal(new[] { 1, 2, 3, 4, 5 }, undone.PageNumbers);
        Assert.False(undone.IsDirty);
        Assert.Equal(new[] { 1, 3, 5, 2, 4 }, redone.PageNumbers);
        Assert.True(redone.IsDirty);
    }

    [Fact]
    public void SelectPage_toggle_can_clear_a_checked_page()
    {
        var state = EditorDocumentState.Create(4)
            .SelectPage(3, PageSelectionMode.Replace)
            .SelectPage(3, PageSelectionMode.Toggle);

        Assert.Empty(state.SelectedPageNumbers);
        Assert.Equal(3, state.SelectionAnchorPageNumber);
        Assert.Equal(3, state.ActivePageNumber);
    }

    [Fact]
    public void DeleteSelectedPages_undo_restores_the_prior_rotation_and_selection()
    {
        var deleted = EditorDocumentState.Create(4)
            .SelectPage(2, PageSelectionMode.Replace)
            .SelectPage(3, PageSelectionMode.Toggle)
            .RotateSelectedPages(90)
            .DeleteSelectedPages();

        var undoDelete = deleted.Undo();
        var undoRotate = undoDelete.Undo();

        Assert.Equal(new[] { 1, 4 }, deleted.PageNumbers);
        Assert.Equal(new[] { 1, 2, 3, 4 }, undoDelete.PageNumbers);
        Assert.Equal(new[] { 2, 3 }, undoDelete.SelectedPageNumbers);
        Assert.Equal(90, undoDelete.GetRotation(2));
        Assert.Equal(90, undoDelete.GetRotation(3));
        Assert.Equal(new[] { 2, 3 }, undoRotate.SelectedPageNumbers);
        Assert.Equal(0, undoRotate.GetRotation(2));
        Assert.Equal(0, undoRotate.GetRotation(3));
        Assert.False(undoRotate.IsDirty);
    }

    [Fact]
    public void RotateSelectedPages_discards_redo_after_a_new_structural_edit()
    {
        var undoneMove = EditorDocumentState.Create(4)
            .SelectPage(2, PageSelectionMode.Replace)
            .MovePageGroup(2, insertionIndex: 4)
            .Undo();

        var rotated = undoneMove.RotateSelectedPages(90);

        Assert.True(rotated.CanUndo);
        Assert.False(rotated.CanRedo);
        Assert.Equal(90, rotated.GetRotation(2));
    }
}
