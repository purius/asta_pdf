using PdfMergeTool.Services;
using Xunit;

namespace PdfMergeTool.Tests;

public sealed class PageOrganizerThumbnailSchedulerTests
{
    [Fact]
    public void TryTakeNext_prioritizes_visible_high_pages_before_background_pages()
    {
        var scheduler = new PageOrganizerThumbnailScheduler(Enumerable.Range(1, 166).ToArray());
        scheduler.Prioritize([97, 98, 99]);

        Assert.True(scheduler.TryTakeNext(out var first));
        scheduler.Complete(first);
        Assert.True(scheduler.TryTakeNext(out var second));
        scheduler.Complete(second);
        Assert.True(scheduler.TryTakeNext(out var third));

        Assert.Equal(new[] { 97, 98, 99 }, new[] { first, second, third });
    }

    [Fact]
    public void Prioritize_promotes_later_visible_pages_ahead_of_stale_cache_priority()
    {
        var scheduler = new PageOrganizerThumbnailScheduler(Enumerable.Range(1, 200).ToArray());
        scheduler.Prioritize(Enumerable.Range(1, 24));
        scheduler.Prioritize(Enumerable.Range(73, 52));
        scheduler.Prioritize([97, 98, 99]);

        Assert.True(scheduler.TryTakeNext(out var first));
        scheduler.Complete(first);
        Assert.True(scheduler.TryTakeNext(out var second));
        scheduler.Complete(second);
        Assert.True(scheduler.TryTakeNext(out var third));

        Assert.Equal(new[] { 97, 98, 99 }, new[] { first, second, third });
    }

    [Fact]
    public void UpdateCacheOrder_uses_the_current_display_order_for_background_priority()
    {
        var scheduler = new PageOrganizerThumbnailScheduler(Enumerable.Range(1, 200).ToArray());
        var currentDisplayOrder = new[] { 200 }.Concat(Enumerable.Range(1, 199)).ToArray();

        scheduler.UpdateCacheOrder(currentDisplayOrder);

        Assert.True(scheduler.TryTakeNext(out var first));

        Assert.Equal(200, first);
    }

    [Fact]
    public void Prioritize_uses_the_current_display_order_for_visible_pages()
    {
        var scheduler = new PageOrganizerThumbnailScheduler(Enumerable.Range(1, 200).ToArray());
        var currentDisplayOrder = new[] { 200 }.Concat(Enumerable.Range(1, 199)).ToArray();

        scheduler.UpdateCacheOrder(currentDisplayOrder);
        scheduler.Prioritize([1, 200]);

        Assert.True(scheduler.TryTakeNext(out var first));
        scheduler.Complete(first);
        Assert.True(scheduler.TryTakeNext(out var second));

        Assert.Equal(new[] { 200, 1 }, new[] { first, second });
    }

    [Fact]
    public void RegisterFailure_retries_once_then_leaves_the_failed_page_out_of_the_queue()
    {
        var scheduler = new PageOrganizerThumbnailScheduler([1, 2, 3]);

        Assert.True(scheduler.TryTakeNext(out var firstAttempt));
        Assert.Equal(1, firstAttempt);
        Assert.True(scheduler.RegisterFailure(firstAttempt));
        Assert.True(scheduler.TryTakeNext(out var retryAttempt));
        Assert.Equal(1, retryAttempt);
        Assert.False(scheduler.RegisterFailure(retryAttempt));
        Assert.True(scheduler.TryTakeNext(out var nextPage));

        Assert.Equal(2, nextPage);
    }

    [Fact]
    public void RequestManualRetry_restores_a_page_after_its_automatic_retry_is_exhausted()
    {
        var scheduler = new PageOrganizerThumbnailScheduler([1, 2]);

        Assert.True(scheduler.TryTakeNext(out var page));
        Assert.True(scheduler.RegisterFailure(page));
        Assert.True(scheduler.TryTakeNext(out page));
        Assert.False(scheduler.RegisterFailure(page));
        Assert.True(scheduler.RequestManualRetry(page));
        Assert.True(scheduler.TryTakeNext(out var retriedPage));

        Assert.Equal(1, retriedPage);
    }

    [Fact]
    public void GetCacheWindow_keeps_a_bounded_range_around_visible_pages()
    {
        var scheduler = new PageOrganizerThumbnailScheduler(Enumerable.Range(1, 200).ToArray());

        var window = scheduler.GetCacheWindow([97, 98, 99, 100]);

        Assert.Equal(52, window.Count);
        Assert.Contains(73, window);
        Assert.Contains(124, window);
        Assert.DoesNotContain(72, window);
        Assert.DoesNotContain(125, window);
    }

    [Fact]
    public void GetCacheWindow_uses_a_single_current_display_order_range_after_reorder()
    {
        var scheduler = new PageOrganizerThumbnailScheduler(Enumerable.Range(1, 200).ToArray());
        var currentDisplayOrder = Enumerable.Range(1, 199).ToList();
        currentDisplayOrder.Insert(96, 200);

        scheduler.UpdateCacheOrder(currentDisplayOrder);

        var window = scheduler.GetCacheWindow([200, 97, 98, 99, 100]);

        Assert.Equal(53, window.Count);
        Assert.Contains(73, window);
        Assert.Contains(200, window);
        Assert.Contains(124, window);
        Assert.DoesNotContain(72, window);
        Assert.DoesNotContain(125, window);
        Assert.DoesNotContain(176, window);
    }

    [Fact]
    public void GetCacheWindow_hard_limits_a_contiguous_current_display_order_range()
    {
        var scheduler = new PageOrganizerThumbnailScheduler(Enumerable.Range(1, 200).ToArray());
        var currentDisplayOrder = new[] { 200 }.Concat(Enumerable.Range(1, 199)).ToArray();
        var visiblePageNumbers = currentDisplayOrder.Skip(65).Take(70).ToArray();

        scheduler.UpdateCacheOrder(currentDisplayOrder);

        var window = scheduler.GetCacheWindow(visiblePageNumbers);
        var cacheIndexes = window
            .Select(pageNumber => Array.IndexOf(currentDisplayOrder, pageNumber))
            .Order()
            .ToArray();

        Assert.Equal(96, window.Count);
        Assert.All(visiblePageNumbers, pageNumber => Assert.Contains(pageNumber, window));
        Assert.Equal(Enumerable.Range(cacheIndexes[0], cacheIndexes.Length), cacheIndexes);
    }
}
