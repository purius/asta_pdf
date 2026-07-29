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
    public void GetCacheWindow_unions_bounded_ranges_for_distant_visible_pages_after_reorder()
    {
        var scheduler = new PageOrganizerThumbnailScheduler(Enumerable.Range(1, 200).ToArray());

        var window = scheduler.GetCacheWindow([200, 1]);

        Assert.Equal(50, window.Count);
        Assert.All(window, page => Assert.True(
            (page >= 1 && page <= 25) || (page >= 176 && page <= 200)));
        Assert.Contains(1, window);
        Assert.Contains(25, window);
        Assert.Contains(176, window);
        Assert.Contains(200, window);
        Assert.DoesNotContain(26, window);
        Assert.DoesNotContain(175, window);
    }
}
