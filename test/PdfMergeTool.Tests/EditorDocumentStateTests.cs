using PdfMergeTool.Services;
using Xunit;

namespace PdfMergeTool.Tests;

public sealed class EditorDocumentStateTests
{
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
}
