using PdfMergeTool.Services;
using Xunit;

namespace PdfMergeTool.Tests;

public sealed class PdfSplitPlannerTests
{
    [Fact]
    public void Parse_recognizes_split_commands()
    {
        Assert.Equal(PdfExplorerCommand.PageSplit, PdfExplorerCommandParser.Parse(new[] { "--split-pages", "source.pdf" }));
        Assert.Equal(PdfExplorerCommand.IntervalSplit, PdfExplorerCommandParser.Parse(new[] { "--split-interval", "source.pdf" }));
        Assert.Equal(PdfExplorerCommand.ParitySplit, PdfExplorerCommandParser.Parse(new[] { "--split-parity", "source.pdf" }));
    }

    [Fact]
    public void CreateIntervalPlan_keeps_final_partial_batch()
    {
        var plan = PdfSplitPlanner.CreateIntervalPlan(Path.Combine(Path.GetTempPath(), "source.pdf"), 10, 3);

        Assert.Equal(
            new[] { "1-3", "4-6", "7-9", "10" },
            plan.Outputs.Select(output => output.PageRange));
    }

    [Fact]
    public void CreateIntervalPlan_rejects_zero_interval()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => PdfSplitPlanner.CreateIntervalPlan(Path.Combine(Path.GetTempPath(), "source.pdf"), 10, 0));
    }

    [Fact]
    public void CreateParityPlan_omits_empty_even_output()
    {
        var plan = PdfSplitPlanner.CreateParityPlan(Path.Combine(Path.GetTempPath(), "source.pdf"), 1);

        var output = Assert.Single(plan.Outputs);
        Assert.Equal(PdfSplitKind.Odd, output.Kind);
        Assert.Equal("1", output.PageRange);
        Assert.EndsWith("source_홀수.pdf", output.OutputPath);
    }

    [Fact]
    public void CreateIntervalPlan_uses_numbered_folder_when_base_folder_exists()
    {
        using var directory = new TemporaryDirectory();
        var sourcePath = Path.Combine(directory.Path, "source.pdf");
        Directory.CreateDirectory(Path.Combine(directory.Path, "source_분할"));

        var plan = PdfSplitPlanner.CreateIntervalPlan(sourcePath, 2, 1);

        Assert.EndsWith("source_분할 (2)", plan.OutputFolder);
    }

    [Fact]
    public void CreateIntervalPlan_preserves_current_page_order_and_rotation()
    {
        var pages = new[]
        {
            new PdfPageTransform(3, 90),
            new PdfPageTransform(1, 0),
            new PdfPageTransform(2, 180)
        };

        var plan = PdfSplitPlanner.CreateIntervalPlan(
            Path.Combine(Path.GetTempPath(), "source.pdf"),
            pages,
            2);

        Assert.Equal("3,1", plan.Outputs[0].PageRange);
        Assert.Equal(new[] { 3, 1 }, plan.Outputs[0].Pages.Select(page => page.PageNumber));
        Assert.Equal(new[] { 90, 0 }, plan.Outputs[0].Pages.Select(page => page.Rotation));
        Assert.Equal("2", plan.Outputs[1].PageRange);
        Assert.Equal(180, plan.Outputs[1].Pages.Single().Rotation);
    }

    [Fact]
    public void CreateParityPlan_preserves_current_page_order_and_rotation()
    {
        var pages = new[]
        {
            new PdfPageTransform(4, 0),
            new PdfPageTransform(3, 90),
            new PdfPageTransform(2, 180),
            new PdfPageTransform(1, 270)
        };

        var plan = PdfSplitPlanner.CreateParityPlan(Path.Combine(Path.GetTempPath(), "source.pdf"), pages);

        var odd = plan.Outputs.Single(output => output.Kind == PdfSplitKind.Odd);
        var even = plan.Outputs.Single(output => output.Kind == PdfSplitKind.Even);
        Assert.Equal(new[] { 3, 1 }, odd.Pages.Select(page => page.PageNumber));
        Assert.Equal(new[] { 90, 270 }, odd.Pages.Select(page => page.Rotation));
        Assert.Equal(new[] { 4, 2 }, even.Pages.Select(page => page.PageNumber));
        Assert.Equal(new[] { 0, 180 }, even.Pages.Select(page => page.Rotation));
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"PdfMergeTool.Tests-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            Directory.Delete(Path, recursive: true);
        }
    }
}
