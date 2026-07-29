namespace PdfMergeTool.Services;

public enum PageSelectionMode
{
    Replace,
    Range,
    Toggle
}

/// <summary>
/// The single source of truth for Page Organizer selection and structural edits.
/// The preview viewer may navigate pages, but it never mutates this state.
/// </summary>
public sealed record EditorDocumentState
{
    private readonly bool _baselineDirty;
    private readonly IReadOnlyList<Snapshot> _undoHistory;
    private readonly IReadOnlyList<Snapshot> _redoHistory;

    private EditorDocumentState(
        IReadOnlyList<int> pageNumbers,
        IReadOnlyList<int> selectedPageNumbers,
        int? selectionAnchorPageNumber,
        int? activePageNumber,
        IReadOnlyDictionary<int, int> pageRotations,
        IReadOnlyList<Snapshot> undoHistory,
        IReadOnlyList<Snapshot> redoHistory,
        bool baselineDirty)
    {
        PageNumbers = pageNumbers;
        SelectedPageNumbers = selectedPageNumbers;
        SelectionAnchorPageNumber = selectionAnchorPageNumber;
        ActivePageNumber = activePageNumber;
        PageRotations = pageRotations;
        _undoHistory = undoHistory;
        _redoHistory = redoHistory;
        _baselineDirty = baselineDirty;
    }

    public IReadOnlyList<int> PageNumbers { get; }

    public IReadOnlyList<int> SelectedPageNumbers { get; }

    public int? SelectionAnchorPageNumber { get; }

    public int? ActivePageNumber { get; }

    public IReadOnlyDictionary<int, int> PageRotations { get; }

    public bool IsDirty => _baselineDirty || _undoHistory.Count > 0;

    public bool CanUndo => _undoHistory.Count > 0;

    public bool CanRedo => _redoHistory.Count > 0;

    public static EditorDocumentState Create(int pageCount, bool isDirty = false)
    {
        if (pageCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(pageCount));
        }

        var pages = Enumerable.Range(1, pageCount).ToArray();
        return CreateState(
            pages,
            [pages[0]],
            pages[0],
            pages[0],
            new Dictionary<int, int>(),
            [],
            [],
            isDirty);
    }

    public EditorDocumentState SelectPage(int pageNumber, PageSelectionMode mode)
    {
        EnsureKnownPage(pageNumber);
        var selected = SelectedPageNumbers.ToHashSet();
        var anchor = SelectionAnchorPageNumber;

        switch (mode)
        {
            case PageSelectionMode.Replace:
                selected = [pageNumber];
                anchor = pageNumber;
                break;
            case PageSelectionMode.Range:
            {
                anchor = anchor is { } value && PageNumbers.Contains(value)
                    ? value
                    : pageNumber;
                var start = IndexOfPage(anchor.Value);
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

                anchor = pageNumber;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(mode));
        }

        return CreateState(
            PageNumbers,
            PageNumbers.Where(selected.Contains),
            anchor,
            pageNumber,
            PageRotations,
            _undoHistory,
            _redoHistory,
            _baselineDirty);
    }

    public EditorDocumentState ActivatePage(int pageNumber)
    {
        EnsureKnownPage(pageNumber);
        if (ActivePageNumber == pageNumber)
        {
            return this;
        }

        return CreateState(
            PageNumbers,
            SelectedPageNumbers,
            SelectionAnchorPageNumber,
            pageNumber,
            PageRotations,
            _undoHistory,
            _redoHistory,
            _baselineDirty);
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

        return RecordEdit(
            nextPages,
            group,
            draggedPageNumber,
            draggedPageNumber,
            PageRotations);
    }

    public EditorDocumentState RotateSelectedPages(int rotationDelta)
    {
        if (SelectedPageNumbers.Count == 0)
        {
            return this;
        }

        var rotations = PageRotations.ToDictionary(pair => pair.Key, pair => pair.Value);
        foreach (var pageNumber in SelectedPageNumbers)
        {
            var current = rotations.TryGetValue(pageNumber, out var value) ? value : 0;
            var next = NormalizeRotation(current + rotationDelta);
            if (next == 0)
            {
                rotations.Remove(pageNumber);
            }
            else
            {
                rotations[pageNumber] = next;
            }
        }

        return RecordEdit(
            PageNumbers,
            SelectedPageNumbers,
            SelectionAnchorPageNumber,
            ActivePageNumber,
            rotations);
    }

    public EditorDocumentState DeleteSelectedPages()
    {
        if (SelectedPageNumbers.Count == 0 || SelectedPageNumbers.Count >= PageNumbers.Count)
        {
            return this;
        }

        var selected = SelectedPageNumbers.ToHashSet();
        var firstRemovedIndex = PageNumbers
            .Select((page, index) => new { page, index })
            .Where(item => selected.Contains(item.page))
            .Select(item => item.index)
            .DefaultIfEmpty(0)
            .Min();
        var remaining = PageNumbers.Where(page => !selected.Contains(page)).ToList();
        var nextActive = remaining[Math.Min(firstRemovedIndex, remaining.Count - 1)];
        var remainingRotations = PageRotations
            .Where(pair => remaining.Contains(pair.Key))
            .ToDictionary(pair => pair.Key, pair => pair.Value);

        return RecordEdit(
            remaining,
            [nextActive],
            nextActive,
            nextActive,
            remainingRotations);
    }

    public EditorDocumentState ReversePageOrder()
    {
        if (PageNumbers.Count <= 1)
        {
            return this;
        }

        return RecordEdit(
            PageNumbers.Reverse(),
            SelectedPageNumbers,
            SelectionAnchorPageNumber,
            ActivePageNumber,
            PageRotations);
    }

    public EditorDocumentState Undo()
    {
        if (_undoHistory.Count == 0)
        {
            return this;
        }

        var previous = _undoHistory[^1];
        return CreateState(
            previous.PageNumbers,
            previous.SelectedPageNumbers,
            previous.SelectionAnchorPageNumber,
            previous.ActivePageNumber,
            previous.PageRotations,
            _undoHistory.Take(_undoHistory.Count - 1),
            _redoHistory.Append(Capture()),
            _baselineDirty);
    }

    public EditorDocumentState Redo()
    {
        if (_redoHistory.Count == 0)
        {
            return this;
        }

        var next = _redoHistory[^1];
        return CreateState(
            next.PageNumbers,
            next.SelectedPageNumbers,
            next.SelectionAnchorPageNumber,
            next.ActivePageNumber,
            next.PageRotations,
            _undoHistory.Append(Capture()),
            _redoHistory.Take(_redoHistory.Count - 1),
            _baselineDirty);
    }

    public EditorDocumentState MarkClean()
    {
        return CreateState(
            PageNumbers,
            SelectedPageNumbers,
            SelectionAnchorPageNumber,
            ActivePageNumber,
            PageRotations,
            [],
            [],
            false);
    }

    public int GetRotation(int pageNumber)
    {
        EnsureKnownPage(pageNumber);
        return PageRotations.TryGetValue(pageNumber, out var rotation) ? rotation : 0;
    }

    private EditorDocumentState RecordEdit(
        IEnumerable<int> pageNumbers,
        IEnumerable<int> selectedPageNumbers,
        int? selectionAnchorPageNumber,
        int? activePageNumber,
        IReadOnlyDictionary<int, int> pageRotations)
    {
        return CreateState(
            pageNumbers,
            selectedPageNumbers,
            selectionAnchorPageNumber,
            activePageNumber,
            pageRotations,
            _undoHistory.Append(Capture()),
            [],
            _baselineDirty);
    }

    private Snapshot Capture()
    {
        return new Snapshot(
            PageNumbers.ToArray(),
            SelectedPageNumbers.ToArray(),
            SelectionAnchorPageNumber,
            ActivePageNumber,
            PageRotations.ToDictionary(pair => pair.Key, pair => pair.Value));
    }

    private static EditorDocumentState CreateState(
        IEnumerable<int> pageNumbers,
        IEnumerable<int> selectedPageNumbers,
        int? selectionAnchorPageNumber,
        int? activePageNumber,
        IReadOnlyDictionary<int, int> pageRotations,
        IEnumerable<Snapshot> undoHistory,
        IEnumerable<Snapshot> redoHistory,
        bool baselineDirty)
    {
        return new EditorDocumentState(
            pageNumbers.ToArray(),
            selectedPageNumbers.ToArray(),
            selectionAnchorPageNumber,
            activePageNumber,
            pageRotations
                .Where(pair => pair.Value != 0)
                .ToDictionary(pair => pair.Key, pair => NormalizeRotation(pair.Value)),
            undoHistory.ToArray(),
            redoHistory.ToArray(),
            baselineDirty);
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

    private static int NormalizeRotation(int rotation)
    {
        var normalized = rotation % 360;
        return normalized < 0 ? normalized + 360 : normalized;
    }

    private sealed record Snapshot(
        IReadOnlyList<int> PageNumbers,
        IReadOnlyList<int> SelectedPageNumbers,
        int? SelectionAnchorPageNumber,
        int? ActivePageNumber,
        IReadOnlyDictionary<int, int> PageRotations);
}
