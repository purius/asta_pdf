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
    public const int MaximumCacheWindowSize = 96;
    private const int MaximumRenderAttempts = 2;
    private readonly int[] _knownPageNumbers;
    private readonly HashSet<int> _knownPages;
    private readonly HashSet<int> _pendingPages;
    private readonly HashSet<int> _inFlightPages = [];
    private readonly LinkedList<int> _priorityPages = [];
    private readonly HashSet<int> _priorityPageSet = [];
    private readonly Dictionary<int, int> _failureCounts = [];
    private int[] _cacheOrder = [];
    private Dictionary<int, int> _cacheOrderIndexes = [];

    public PageOrganizerThumbnailScheduler(IReadOnlyList<int> pageNumbers)
    {
        _knownPageNumbers = pageNumbers.Distinct().ToArray();
        _knownPages = _knownPageNumbers.ToHashSet();
        _pendingPages = _knownPageNumbers.ToHashSet();
        UpdateCacheOrder(pageNumbers);
    }

    public void UpdateCacheOrder(IReadOnlyList<int> pageNumbers)
    {
        var nextOrder = new List<int>(_knownPageNumbers.Length);
        var includedPages = new HashSet<int>();
        foreach (var pageNumber in pageNumbers)
        {
            if (_knownPages.Contains(pageNumber) && includedPages.Add(pageNumber))
            {
                nextOrder.Add(pageNumber);
            }
        }

        foreach (var pageNumber in _cacheOrder)
        {
            if (includedPages.Add(pageNumber))
            {
                nextOrder.Add(pageNumber);
            }
        }

        foreach (var pageNumber in _knownPageNumbers)
        {
            if (includedPages.Add(pageNumber))
            {
                nextOrder.Add(pageNumber);
            }
        }

        _cacheOrder = nextOrder.ToArray();
        _cacheOrderIndexes = _cacheOrder
            .Select((pageNumber, index) => new { pageNumber, index })
            .ToDictionary(entry => entry.pageNumber, entry => entry.index);
    }

    public void Prioritize(IEnumerable<int> pageNumbers)
    {
        var promotedPageNumbers = pageNumbers
            .Where(_pendingPages.Contains)
            .Distinct()
            .OrderBy(pageNumber => _cacheOrderIndexes[pageNumber])
            .ToArray();

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

        foreach (var candidate in _cacheOrder)
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
            .Where(_cacheOrderIndexes.ContainsKey)
            .Select(pageNumber => _cacheOrderIndexes[pageNumber])
            .Order()
            .ToArray();
        if (visibleIndexes.Length == 0)
        {
            return _cacheOrder
                .Take(Math.Min(MinimumCacheRadius, MaximumCacheWindowSize))
                .ToHashSet();
        }

        var maximumWindowSize = Math.Min(MaximumCacheWindowSize, _cacheOrder.Length);
        var radius = Math.Max(MinimumCacheRadius, visibleIndexes.Length * 2);
        var firstVisibleIndex = visibleIndexes[0];
        var lastVisibleIndex = visibleIndexes[^1];
        var firstCacheIndex = Math.Max(0, firstVisibleIndex - radius);
        var lastCacheIndex = Math.Min(_cacheOrder.Length - 1, lastVisibleIndex + radius);

        if (lastCacheIndex - firstCacheIndex + 1 > maximumWindowSize)
        {
            var visibleSpan = lastVisibleIndex - firstVisibleIndex + 1;
            if (visibleSpan >= maximumWindowSize)
            {
                firstCacheIndex = Math.Clamp(
                    firstVisibleIndex,
                    0,
                    _cacheOrder.Length - maximumWindowSize);
            }
            else
            {
                var leadingCachePages = (maximumWindowSize - visibleSpan) / 2;
                firstCacheIndex = Math.Clamp(
                    firstVisibleIndex - leadingCachePages,
                    0,
                    _cacheOrder.Length - maximumWindowSize);
            }

            lastCacheIndex = firstCacheIndex + maximumWindowSize - 1;
        }

        return _cacheOrder
            .Skip(firstCacheIndex)
            .Take(lastCacheIndex - firstCacheIndex + 1)
            .ToHashSet();
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
