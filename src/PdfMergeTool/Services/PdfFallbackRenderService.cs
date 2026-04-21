using System.Collections.Concurrent;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using PdfiumViewer;

namespace PdfMergeTool.Services;

public sealed class PdfFallbackRenderService : IDisposable
{
    private const int DefaultLargePageThreshold = 180;
    private const long DefaultLargeFileThresholdBytes = 80L * 1024 * 1024;
    private readonly ConcurrentDictionary<string, PdfFallbackSession> _sessions = new(StringComparer.OrdinalIgnoreCase);

    public PdfFallbackSessionInfo OpenDocument(string path)
    {
        AppPaths.EnsureDirectories();
        CleanupStaleCaches();

        var sessionId = Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture);
        var cacheDirectory = Path.Combine(AppPaths.FallbackCacheDirectory, sessionId);
        Directory.CreateDirectory(cacheDirectory);

        FileStream stream = new(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);

        try
        {
            var document = PdfDocument.Load(stream);
            var pageInfos = document.PageSizes
                .Select((size, index) => new PdfFallbackPageInfo(index + 1, size.Width, size.Height))
                .ToList();
            var largeDocumentMode = pageInfos.Count >= DefaultLargePageThreshold ||
                                    new FileInfo(path).Length >= DefaultLargeFileThresholdBytes;

            var session = new PdfFallbackSession(
                sessionId,
                path,
                stream,
                document,
                cacheDirectory,
                pageInfos,
                largeDocumentMode);

            if (!_sessions.TryAdd(sessionId, session))
            {
                session.Dispose();
                throw new InvalidOperationException("PDF fallback session creation failed.");
            }

            return new PdfFallbackSessionInfo(sessionId, pageInfos, largeDocumentMode);
        }
        catch
        {
            stream.Dispose();
            TryDeleteDirectory(cacheDirectory);
            throw;
        }
    }

    public async Task<PdfFallbackRenderedPage> RenderPageAsync(
        string sessionId,
        int pageNumber,
        int targetWidth,
        int rotationDegrees,
        bool thumbnail,
        CancellationToken cancellationToken)
    {
        if (!_sessions.TryGetValue(sessionId, out var session))
        {
            throw new InvalidOperationException("PDF fallback session not found.");
        }

        var pageIndex = pageNumber - 1;
        if (pageIndex < 0 || pageIndex >= session.PageInfos.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(pageNumber));
        }

        var pageInfo = session.PageInfos[pageIndex];
        var normalizedRotation = NormalizeRotation(rotationDegrees);
        var width = Math.Clamp(targetWidth, thumbnail ? 48 : 240, thumbnail ? 240 : 2200);
        var cacheKey = $"{pageNumber}-{(thumbnail ? "thumb" : "main")}-{width}-{normalizedRotation}";
        if (session.RenderCache.TryGetValue(cacheKey, out var cachedPath) && File.Exists(cachedPath))
        {
            return new PdfFallbackRenderedPage(pageNumber, cachedPath, pageInfo.Width, pageInfo.Height);
        }

        await session.RenderGate.WaitAsync(cancellationToken);
        try
        {
            if (session.RenderCache.TryGetValue(cacheKey, out cachedPath) && File.Exists(cachedPath))
            {
                return new PdfFallbackRenderedPage(pageNumber, cachedPath, pageInfo.Width, pageInfo.Height);
            }

            var renderSize = CalculateRenderSize(pageInfo, width, normalizedRotation);
            var renderFlags = PdfRenderFlags.Annotations |
                              PdfRenderFlags.LcdText |
                              PdfRenderFlags.CorrectFromDpi;
            using var image = session.Document.Render(
                pageIndex,
                renderSize.Width,
                renderSize.Height,
                96,
                96,
                ToPdfRotation(normalizedRotation),
                renderFlags);
            using var bitmap = new Bitmap(image);
            var outputPath = Path.Combine(session.CacheDirectory, $"{cacheKey}.png");
            bitmap.Save(outputPath, ImageFormat.Png);
            session.RenderCache[cacheKey] = outputPath;
            return new PdfFallbackRenderedPage(pageNumber, outputPath, pageInfo.Width, pageInfo.Height);
        }
        finally
        {
            session.RenderGate.Release();
        }
    }

    public async Task<IReadOnlyList<PdfImagePage>> ExportPagesAsImagesAsync(
        string sessionId,
        IReadOnlyList<int> pageOrder,
        IReadOnlyDictionary<int, int> rotations,
        int maxWidth,
        int maxHeight,
        long jpegQuality,
        CancellationToken cancellationToken)
    {
        if (!_sessions.TryGetValue(sessionId, out var session))
        {
            throw new InvalidOperationException("PDF fallback session not found.");
        }

        var pages = new List<PdfImagePage>(pageOrder.Count);
        foreach (var pageNumber in pageOrder)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var pageIndex = pageNumber - 1;
            if (pageIndex < 0 || pageIndex >= session.PageInfos.Count)
            {
                continue;
            }

            var rotation = rotations.TryGetValue(pageNumber, out var customRotation)
                ? NormalizeRotation(customRotation)
                : 0;
            var pageInfo = session.PageInfos[pageIndex];
            var renderSize = CalculateBoundedRenderSize(pageInfo, maxWidth, maxHeight, rotation);
            var renderFlags = PdfRenderFlags.Annotations |
                              PdfRenderFlags.LcdText |
                              PdfRenderFlags.CorrectFromDpi |
                              PdfRenderFlags.ForPrinting;

            var jpegBytes = await Task.Run(() =>
            {
                using var image = session.Document.Render(
                    pageIndex,
                    renderSize.Width,
                    renderSize.Height,
                    96,
                    96,
                    ToPdfRotation(rotation),
                    renderFlags);
                using var bitmap = new Bitmap(image);
                using var stream = new MemoryStream();
                SaveJpeg(bitmap, stream, jpegQuality);
                return stream.ToArray();
            }, cancellationToken);

            pages.Add(new PdfImagePage(jpegBytes, renderSize.Width, renderSize.Height));
        }

        return pages;
    }

    public void CloseSession(string? sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return;
        }

        if (_sessions.TryRemove(sessionId, out var session))
        {
            session.Dispose();
        }
    }

    public void Dispose()
    {
        foreach (var sessionId in _sessions.Keys.ToList())
        {
            CloseSession(sessionId);
        }
    }

    private static Size CalculateRenderSize(PdfFallbackPageInfo pageInfo, int width, int rotationDegrees)
    {
        var sourceWidth = UsesSwappedAxes(rotationDegrees) ? pageInfo.Height : pageInfo.Width;
        var sourceHeight = UsesSwappedAxes(rotationDegrees) ? pageInfo.Width : pageInfo.Height;
        var scale = width / Math.Max(sourceWidth, 1d);
        var height = Math.Max(1, (int)Math.Round(sourceHeight * scale, MidpointRounding.AwayFromZero));
        return new Size(width, height);
    }

    private static Size CalculateBoundedRenderSize(PdfFallbackPageInfo pageInfo, int maxWidth, int maxHeight, int rotationDegrees)
    {
        var sourceWidth = UsesSwappedAxes(rotationDegrees) ? pageInfo.Height : pageInfo.Width;
        var sourceHeight = UsesSwappedAxes(rotationDegrees) ? pageInfo.Width : pageInfo.Height;
        var scale = Math.Min(maxWidth / Math.Max(sourceWidth, 1d), maxHeight / Math.Max(sourceHeight, 1d));
        scale = Math.Max(scale, 0.05);
        var width = Math.Max(1, (int)Math.Round(sourceWidth * scale, MidpointRounding.AwayFromZero));
        var height = Math.Max(1, (int)Math.Round(sourceHeight * scale, MidpointRounding.AwayFromZero));
        return new Size(width, height);
    }

    private static int NormalizeRotation(int rotationDegrees)
    {
        var normalized = rotationDegrees % 360;
        if (normalized < 0)
        {
            normalized += 360;
        }

        return normalized;
    }

    private static bool UsesSwappedAxes(int rotationDegrees)
    {
        return rotationDegrees is 90 or 270;
    }

    private static PdfRotation ToPdfRotation(int rotationDegrees)
    {
        return NormalizeRotation(rotationDegrees) switch
        {
            90 => PdfRotation.Rotate90,
            180 => PdfRotation.Rotate180,
            270 => PdfRotation.Rotate270,
            _ => PdfRotation.Rotate0
        };
    }

    private static void SaveJpeg(Bitmap bitmap, Stream stream, long jpegQuality)
    {
        var encoder = ImageCodecInfo.GetImageDecoders()
            .FirstOrDefault(codec => string.Equals(codec.FormatID.ToString(), ImageFormat.Jpeg.Guid.ToString(), StringComparison.OrdinalIgnoreCase));
        if (encoder is null)
        {
            bitmap.Save(stream, ImageFormat.Jpeg);
            return;
        }

        using var parameters = new EncoderParameters(1);
        parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, jpegQuality);
        bitmap.Save(stream, encoder, parameters);
    }

    private static void CleanupStaleCaches()
    {
        foreach (var directory in Directory.EnumerateDirectories(AppPaths.FallbackCacheDirectory))
        {
            try
            {
                if (Directory.GetCreationTimeUtc(directory) < DateTime.UtcNow.AddDays(-1))
                {
                    Directory.Delete(directory, recursive: true);
                }
            }
            catch
            {
                // Best effort cleanup only.
            }
        }
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch
        {
            // Best effort cleanup only.
        }
    }

    private sealed class PdfFallbackSession(
        string sessionId,
        string sourcePath,
        FileStream stream,
        PdfDocument document,
        string cacheDirectory,
        IReadOnlyList<PdfFallbackPageInfo> pageInfos,
        bool largeDocumentMode) : IDisposable
    {
        public string SessionId { get; } = sessionId;
        public string SourcePath { get; } = sourcePath;
        public FileStream Stream { get; } = stream;
        public PdfDocument Document { get; } = document;
        public string CacheDirectory { get; } = cacheDirectory;
        public IReadOnlyList<PdfFallbackPageInfo> PageInfos { get; } = pageInfos;
        public bool LargeDocumentMode { get; } = largeDocumentMode;
        public SemaphoreSlim RenderGate { get; } = new(1, 1);
        public ConcurrentDictionary<string, string> RenderCache { get; } = new(StringComparer.OrdinalIgnoreCase);

        public void Dispose()
        {
            Document.Dispose();
            Stream.Dispose();
            RenderGate.Dispose();
            TryDeleteDirectory(CacheDirectory);
        }
    }
}

public sealed record PdfFallbackSessionInfo(
    string SessionId,
    IReadOnlyList<PdfFallbackPageInfo> Pages,
    bool LargeDocumentMode);

public sealed record PdfFallbackPageInfo(int Number, double Width, double Height);

public sealed record PdfFallbackRenderedPage(int PageNumber, string ImagePath, double SourceWidth, double SourceHeight);
