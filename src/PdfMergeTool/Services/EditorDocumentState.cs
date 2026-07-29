namespace PdfMergeTool.Services;

public enum PageSelectionMode
{
    Replace,
    Range,
    Toggle
}

public sealed record EditorDocumentState(
    IReadOnlyList<int> PageNumbers,
    IReadOnlyList<int> SelectedPageNumbers,
    int? SelectionAnchorPageNumber,
    int? ActivePageNumber,
    bool IsDirty)
{
    public static EditorDocumentState Create(int pageCount)
    {
        if (pageCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(pageCount));
        }

        var pages = Enumerable.Range(1, pageCount).ToList();
        return new EditorDocumentState(pages, [pages[0]], pages[0], pages[0], false);
    }

    public EditorDocumentState SelectPage(int pageNumber, PageSelectionMode mode)
    {
        EnsureKnownPage(pageNumber);
        var selected = SelectedPageNumbers.ToHashSet();

        switch (mode)
        {
            case PageSelectionMode.Replace:
                selected = [pageNumber];
                break;
            case PageSelectionMode.Range:
            {
                var anchor = SelectionAnchorPageNumber is { } value && PageNumbers.Contains(value)
                    ? value
                    : pageNumber;
                var start = IndexOfPage(anchor);
                var end = IndexOfPage(pageNumber);
                selected = PageNumbers
                    .Skip(Math.Min(start, end))
                    .Take(Math.Abs(end - start) + 1)
                    .ToHashSet();
                break;
            }
            case PageSelectionMode.Toggle:
                if (!selected.Add(pageNumber))
                {
                    selected.Remove(pageNumber);
                }

                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(mode));
        }

        return this with
        {
            SelectedPageNumbers = PageNumbers.Where(selected.Contains).ToList(),
            SelectionAnchorPageNumber = pageNumber,
            ActivePageNumber = pageNumber
        };
    }

    public EditorDocumentState MovePageGroup(int draggedPageNumber, int insertionIndex)
    {
        EnsureKnownPage(draggedPageNumber);
        if (insertionIndex < 0 || insertionIndex > PageNumbers.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(insertionIndex));
        }

        var selected = SelectedPageNumbers.ToHashSet();
        var group = selected.Contains(draggedPageNumber)
            ? PageNumbers.Where(selected.Contains).ToList()
            : [draggedPageNumber];
        var remaining = PageNumbers.Where(page => !group.Contains(page)).ToList();
        var removedBeforeInsertion = PageNumbers.Take(insertionIndex).Count(group.Contains);
        var targetIndex = Math.Clamp(insertionIndex - removedBeforeInsertion, 0, remaining.Count);
        var nextPages = remaining.ToList();
        nextPages.InsertRange(targetIndex, group);

        if (nextPages.SequenceEqual(PageNumbers))
        {
            return this;
        }

        return this with
        {
            PageNumbers = nextPages,
            SelectedPageNumbers = group,
            SelectionAnchorPageNumber = draggedPageNumber,
            ActivePageNumber = draggedPageNumber,
            IsDirty = true
        };
    }

    private void EnsureKnownPage(int pageNumber)
    {
        if (!PageNumbers.Contains(pageNumber))
        {
            throw new ArgumentOutOfRangeException(nameof(pageNumber));
        }
    }

    private int IndexOfPage(int pageNumber)
    {
        for (var index = 0; index < PageNumbers.Count; index++)
        {
            if (PageNumbers[index] == pageNumber)
            {
                return index;
            }
        }

        return -1;
    }
}
