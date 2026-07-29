using PdfMergeTool.Services;
using Xunit;

namespace PdfMergeTool.Tests;

public sealed class PageOrganizerViewportTests
{
    [Fact]
    public void GetVerticalOffsetToReveal_returns_null_when_item_is_fully_visible()
    {
        var result = PageOrganizerViewport.GetVerticalOffsetToReveal(100, 300, 20, 140, 900);

        Assert.Null(result);
    }

    [Fact]
    public void GetVerticalOffsetToReveal_moves_only_the_amount_needed_below_viewport()
    {
        var result = PageOrganizerViewport.GetVerticalOffsetToReveal(100, 300, 280, 340, 900);

        Assert.Equal(140d, result);
    }

    [Fact]
    public void GetVerticalOffsetToReveal_moves_only_the_amount_needed_above_viewport()
    {
        var result = PageOrganizerViewport.GetVerticalOffsetToReveal(100, 300, -18, 82, 900);

        Assert.Equal(82d, result);
    }

    [Fact]
    public void GetVerticalOffsetToReveal_clamps_to_scrollable_height()
    {
        var result = PageOrganizerViewport.GetVerticalOffsetToReveal(900, 200, 180, 250, 920);

        Assert.Equal(920d, result);
    }
}
