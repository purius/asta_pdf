namespace PdfMergeTool.Services;

public static class PageOrganizerViewport
{
    public static double? GetVerticalOffsetToReveal(
        double currentOffset,
        double viewportHeight,
        double itemTop,
        double itemBottom,
        double scrollableHeight)
    {
        if (viewportHeight <= 0 || itemBottom <= itemTop || scrollableHeight <= 0)
        {
            return null;
        }

        var targetOffset = itemTop < 0
            ? currentOffset + itemTop
            : itemBottom > viewportHeight
                ? currentOffset + itemBottom - viewportHeight
                : currentOffset;
        targetOffset = Math.Clamp(targetOffset, 0, scrollableHeight);

        return Math.Abs(targetOffset - currentOffset) < 0.1 ? null : targetOffset;
    }

    public static double? GetVerticalOffsetToRevealIndexedRow(
        double currentOffset,
        double viewportHeight,
        int rowIndex,
        double estimatedRowHeight,
        double scrollableHeight)
    {
        if (rowIndex < 0 || estimatedRowHeight <= 0)
        {
            return null;
        }

        var itemTop = rowIndex * estimatedRowHeight - currentOffset;
        return GetVerticalOffsetToReveal(
            currentOffset,
            viewportHeight,
            itemTop,
            itemTop + estimatedRowHeight,
            scrollableHeight);
    }
}
