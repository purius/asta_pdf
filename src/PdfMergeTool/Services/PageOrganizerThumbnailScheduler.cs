namespace PdfMergeTool.Services;

public enum PageOrganizerThumbnailRenderState
{
    Pending,
    Loading,
    Ready,
    Failed,
    Evicted
}

public sealed class PageOrganizerThumbnailScheduler
{
    private const int MinimumCacheRadius = 24;
    private const int MaximumRenderAttempts = 2;
    private readonly int[] _pageNumbers;
    private readonly HashSet<int> _knownPages;
    private readonly Dictionary<int, int> _pageIndexes;
    private readonly HashSet<int> _pendingPages;
    private readonly HashSet<int> _inFlightPages = [];
    private readonly LinkedList<int> _priorityPages = [];
    private readonly HashSet<int> _priorityPageSet = [];
    private readonly Dictionary<int, int> _failureCounts = [];

    public PageOrganizerThumbnailScheduler(IReadOnlyList<int> pageNumbers)
    {
        _pageNumbers = pageNumbers.Distinct().ToArray();
        _knownPages = _pageNumbers.ToHashSet();
        _pageIndexes = _pageNumbers
            .Select((pageNumber, index) => new { pageNumber, index })
            .ToDictionary(entry => entry.pageNumber, entry => entry.index);
        _pendingPages = _pageNumbers.ToHashSet();
    }

    public void Prioritize(IEnumerable<int> pageNumbers)
    {
        var promotedPageNumbers = new List<int>();
        var seenPageNumbers = new HashSet<int>();
        foreach (var pageNumber in pageNumbers)
        {
            if (_pendingPages.Contains(pageNumber) && seenPageNumbers.Add(pageNumber))
            {
                promotedPageNumbers.Add(pageNumber);
            }
        }

        foreach (var pageNumber in promotedPageNumbers)
        {
            if (_priorityPageSet.Remove(pageNumber))
            {
                _priorityPages.Remove(pageNumber);
            }
        }

        for (var index = promotedPageNumbers.Count - 1; index >= 0; index--)
        {
            var pageNumber = promotedPageNumbers[index];
            _priorityPageSet.Add(pageNumber);
            _priorityPages.AddFirst(pageNumber);
        }
    }

    public bool Request(int pageNumber, bool priority)
    {
        if (!_knownPages.Contains(pageNumber) || _inFlightPages.Contains(pageNumber))
        {
            return false;
        }

        var added = _pendingPages.Add(pageNumber);
        if (priority)
        {
            Prioritize([pageNumber]);
        }

        return added;
    }

    public bool TryTakeNext(out int pageNumber)
    {
        while (_priorityPages.First is { } priority)
        {
            _priorityPages.RemoveFirst();
            _priorityPageSet.Remove(priority.Value);
            if (TryStart(priority.Value, out pageNumber))
            {
                return true;
            }
        }

        foreach (var candidate in _pageNumbers)
        {
            if (TryStart(candidate, out pageNumber))
            {
                return true;
            }
        }

        pageNumber = default;
        return false;
    }

    public void Complete(int pageNumber)
    {
        _inFlightPages.Remove(pageNumber);
        _failureCounts.Remove(pageNumber);
    }

    public bool RegisterFailure(int pageNumber)
    {
        _inFlightPages.Remove(pageNumber);
        var failures = _failureCounts.TryGetValue(pageNumber, out var existing)
            ? existing + 1
            : 1;
        _failureCounts[pageNumber] = failures;
        return failures < MaximumRenderAttempts && Request(pageNumber, priority: true);
    }

    public bool RequestManualRetry(int pageNumber)
    {
        _failureCounts.Remove(pageNumber);
        return Request(pageNumber, priority: true);
    }

    public IReadOnlySet<int> GetCacheWindow(IReadOnlyCollection<int> visiblePageNumbers)
    {
        var visibleIndexes = visiblePageNumbers
            .Where(_pageIndexes.ContainsKey)
            .Select(pageNumber => _pageIndexes[pageNumber])
            .Order()
            .ToArray();
        if (visibleIndexes.Length == 0)
        {
            return _pageNumbers.Take(MinimumCacheRadius).ToHashSet();
        }

        var radius = Math.Max(MinimumCacheRadius, visibleIndexes.Length * 2);
        var cacheWindow = new HashSet<int>();
        foreach (var visibleIndex in visibleIndexes)
        {
            var firstIndex = Math.Max(0, visibleIndex - radius);
            var lastIndex = Math.Min(_pageNumbers.Length - 1, visibleIndex + radius);
            for (var index = firstIndex; index <= lastIndex; index++)
            {
                cacheWindow.Add(_pageNumbers[index]);
            }
        }

        return cacheWindow;
    }

    private bool TryStart(int pageNumber, out int nextPageNumber)
    {
        if (_pendingPages.Remove(pageNumber))
        {
            _inFlightPages.Add(pageNumber);
            nextPageNumber = pageNumber;
            return true;
        }

        nextPageNumber = default;
        return false;
    }
}
