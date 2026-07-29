using PdfMergeTool.Services;
using Xunit;

namespace PdfMergeTool.Tests;

public sealed class PdfSavePublisherTests
{
    [Fact]
    public void Publish_replaces_existing_target_only_after_temp_file_is_ready()
    {
        using var directory = new TemporaryDirectory();
        var target = Path.Combine(directory.Path, "working.pdf");
        var temporary = Path.Combine(directory.Path, "working.pending.pdf");
        File.WriteAllText(target, "previous");
        File.WriteAllText(temporary, "published");

        PdfSavePublisher.Publish(temporary, target);

        Assert.Equal("published", File.ReadAllText(target));
        Assert.False(File.Exists(temporary));
    }

    [Fact]
    public void Publish_keeps_existing_target_when_temp_file_is_missing()
    {
        using var directory = new TemporaryDirectory();
        var target = Path.Combine(directory.Path, "working.pdf");
        File.WriteAllText(target, "previous");

        Assert.Throws<FileNotFoundException>(() => PdfSavePublisher.Publish(Path.Combine(directory.Path, "missing.pdf"), target));

        Assert.Equal("previous", File.ReadAllText(target));
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"PdfMergeTool.Tests-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose() => Directory.Delete(Path, recursive: true);
    }
}
