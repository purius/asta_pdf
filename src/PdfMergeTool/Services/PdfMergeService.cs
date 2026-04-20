using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using PdfMergeTool.Models;

namespace PdfMergeTool.Services;

public sealed class PdfMergeService
{
    private const double A4WidthPoints = 595.28;
    private const double A4HeightPoints = 841.89;
    private const double ImagePageMarginPoints = 36;

    public async Task<PdfMergeResult> MergeAsync(IReadOnlyList<PdfInputFile> files, string outputPath, CancellationToken cancellationToken)
    {
        if (files.Count < 2)
        {
            throw new InvalidOperationException("두 개 이상의 PDF를 선택해야 합니다.");
        }

        var args = new List<string>
        {
            "--warning-exit-0",
            "--empty",
            "--pages"
        };

        foreach (var file in files)
        {
            args.Add(file.Path);
            if (!string.IsNullOrWhiteSpace(file.PageRange))
            {
                args.Add(file.PageRange.Trim());
            }
        }

        args.Add("--");
        args.Add(outputPath);

        return await RunQpdfAsync(args, "qpdf 병합 실패", outputPath, cancellationToken);
    }

    public async Task<PdfMergeResult> InterleaveAsync(
        PdfInputFile oddPagesFile,
        PdfInputFile evenPagesFile,
        string outputPath,
        CancellationToken cancellationToken)
    {
        var ranges = new List<(string Path, string? Range)>
        {
            (oddPagesFile.Path, string.IsNullOrWhiteSpace(oddPagesFile.PageRange) ? null : oddPagesFile.PageRange.Trim()),
            (evenPagesFile.Path, string.IsNullOrWhiteSpace(evenPagesFile.PageRange) ? "z-1" : evenPagesFile.PageRange.Trim())
        };

        return await RunPageAssemblyAsync(ranges, outputPath, "--collate=1", "PDF 섞기 실패", cancellationToken);
    }

    public async Task<PdfMergeResult> ReorderPagesAsync(
        string inputPath,
        IReadOnlyList<int> pageOrder,
        string outputPath,
        CancellationToken cancellationToken)
    {
        if (pageOrder.Count == 0)
        {
            throw new InvalidOperationException("저장할 페이지 순서가 없습니다.");
        }

        var pages = pageOrder.Select(page => new PdfPageTransform(page, 0)).ToList();
        return await SaveTransformedPagesAsync(inputPath, pages, outputPath, cancellationToken);
    }

    public async Task<PdfMergeResult> SaveTransformedPagesAsync(
        string inputPath,
        IReadOnlyList<PdfPageTransform> pages,
        string outputPath,
        CancellationToken cancellationToken)
    {
        if (pages.Count == 0)
        {
            throw new InvalidOperationException("저장할 페이지가 없습니다.");
        }

        var outputFolder = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        var normalizedRotations = pages
            .Select((page, index) => new { Position = index + 1, Rotation = NormalizeRotation(page.Rotation) })
            .Where(page => page.Rotation != 0)
            .GroupBy(page => page.Rotation)
            .ToDictionary(group => group.Key, group => string.Join(",", group.Select(item => item.Position)));

        var pageRange = string.Join(",", pages.Select(page => page.PageNumber));
        if (normalizedRotations.Count == 0)
        {
            return await RunPageAssemblyAsync([(inputPath, pageRange)], outputPath, null, "qpdf 페이지 저장 실패", cancellationToken);
        }

        var tempPath = Path.Combine(
            Path.GetTempPath(),
            $"PdfMergeTool-{Guid.NewGuid():N}.pdf");

        try
        {
            await RunPageAssemblyAsync([(inputPath, pageRange)], tempPath, null, "qpdf 페이지 저장 실패", cancellationToken);
            var rotateArgs = new List<string>();
            foreach (var (rotation, range) in normalizedRotations)
            {
                rotateArgs.Add($"--rotate=+{rotation}:{range}");
            }

            return await RunQpdfAsync(
                [.. rotateArgs.Prepend(tempPath).Append(outputPath)],
                "qpdf 페이지 회전 저장 실패",
                outputPath,
                cancellationToken);
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    public async Task<IReadOnlyList<PdfMergeResult>> SplitPagesAsync(
        string inputPath,
        IReadOnlyList<PdfPageTransform> pages,
        string outputFolder,
        string baseName,
        CancellationToken cancellationToken)
    {
        if (pages.Count == 0)
        {
            throw new InvalidOperationException("분할할 페이지가 없습니다.");
        }

        Directory.CreateDirectory(outputFolder);
        var results = new List<PdfMergeResult>();
        for (var index = 0; index < pages.Count; index++)
        {
            var outputPath = Path.Combine(outputFolder, $"{baseName}_{index + 1:000}.pdf");
            results.Add(await SaveTransformedPagesAsync(inputPath, [pages[index]], outputPath, cancellationToken));
        }

        return results;
    }

    public async Task<PdfMergeResult> InsertPdfPagesAsync(
        string inputPath,
        IReadOnlyList<PdfPageTransform> pages,
        string insertPath,
        int insertionIndex,
        string outputPath,
        CancellationToken cancellationToken)
    {
        if (pages.Count == 0)
        {
            throw new InvalidOperationException("삽입할 기준 PDF 페이지가 없습니다.");
        }

        if (!File.Exists(insertPath))
        {
            throw new FileNotFoundException("삽입할 PDF 파일을 찾을 수 없습니다.", insertPath);
        }

        if (insertionIndex < 0 || insertionIndex > pages.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(insertionIndex));
        }

        var basePath = Path.Combine(
            Path.GetTempPath(),
            $"PdfMergeTool-base-{Guid.NewGuid():N}.pdf");

        try
        {
            await SaveTransformedPagesAsync(inputPath, pages, basePath, cancellationToken);

            var ranges = new List<(string Path, string? Range)>();
            if (insertionIndex > 0)
            {
                ranges.Add((basePath, $"1-{insertionIndex}"));
            }

            ranges.Add((insertPath, "1-z"));

            if (insertionIndex < pages.Count)
            {
                ranges.Add((basePath, $"{insertionIndex + 1}-z"));
            }

            return await RunPageAssemblyAsync(ranges, outputPath, null, "PDF 페이지 삽입 실패", cancellationToken);
        }
        finally
        {
            TryDelete(basePath);
        }
    }

    public void CreateBlankA4Pdf(string outputPath)
    {
        WriteSingleA4PagePdf(outputPath, null);
    }

    public void CreateImageA4Pdf(string outputPath, byte[] jpegBytes, int imageWidth, int imageHeight)
    {
        if (jpegBytes.Length == 0)
        {
            throw new InvalidOperationException("클립보드 이미지 데이터가 비어 있습니다.");
        }

        if (imageWidth <= 0 || imageHeight <= 0)
        {
            throw new InvalidOperationException("클립보드 이미지 크기를 확인할 수 없습니다.");
        }

        WriteSingleA4PagePdf(outputPath, new PdfImage(jpegBytes, imageWidth, imageHeight));
    }

    public async Task<PdfMergeResult> CreateA4ImagePagesPdfAsync(
        IReadOnlyList<PdfImagePage> images,
        string outputPath,
        CancellationToken cancellationToken)
    {
        if (images.Count == 0)
        {
            throw new InvalidOperationException("A4로 변환할 페이지 이미지가 없습니다.");
        }

        var tempPaths = new List<string>();
        try
        {
            foreach (var image in images)
            {
                var tempPath = Path.Combine(
                    Path.GetTempPath(),
                    $"PdfMergeTool-a4-page-{Guid.NewGuid():N}.pdf");
                CreateImageA4Pdf(tempPath, image.JpegBytes, image.Width, image.Height);
                tempPaths.Add(tempPath);
            }

            return await RunPageAssemblyAsync(
                tempPaths.Select(path => (path, (string?)"1")).ToList(),
                outputPath,
                null,
                "A4 PDF 저장 실패",
                cancellationToken);
        }
        finally
        {
            foreach (var tempPath in tempPaths)
            {
                TryDelete(tempPath);
            }
        }
    }

    private async Task<PdfMergeResult> RunPageAssemblyAsync(
        IReadOnlyList<(string Path, string? Range)> inputs,
        string outputPath,
        string? pagesOption,
        string failurePrefix,
        CancellationToken cancellationToken)
    {
        var args = new List<string>
        {
            "--warning-exit-0"
        };

        if (!string.IsNullOrWhiteSpace(pagesOption))
        {
            args.Add(pagesOption);
        }

        args.Add("--empty");
        args.Add("--pages");

        foreach (var input in inputs)
        {
            args.Add(input.Path);
            if (!string.IsNullOrWhiteSpace(input.Range))
            {
                args.Add(input.Range);
            }
        }

        args.Add("--");
        args.Add(outputPath);

        return await RunQpdfAsync(args, failurePrefix, outputPath, cancellationToken);
    }

    private static async Task<PdfMergeResult> RunQpdfAsync(
        IReadOnlyList<string> args,
        string failurePrefix,
        string outputPath,
        CancellationToken cancellationToken)
    {
        var outputFolder = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        var effectiveOutputPath = CreateEffectiveOutputPath(outputPath);
        var qpdfArgs = args.ToList();
        if (qpdfArgs.Count == 0)
        {
            throw new InvalidOperationException("qpdf 인자가 비어 있습니다.");
        }

        qpdfArgs[^1] = effectiveOutputPath;

        var startInfo = new ProcessStartInfo
        {
            FileName = ResolveQpdfPath(),
            UseShellExecute = false,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            CreateNoWindow = true
        };

        foreach (var arg in qpdfArgs)
        {
            startInfo.ArgumentList.Add(arg);
        }

        try
        {
            using var process = Process.Start(startInfo)
                ?? throw new InvalidOperationException("qpdf 프로세스를 시작하지 못했습니다.");

            var stdOutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stdErrTask = process.StandardError.ReadToEndAsync(cancellationToken);
            await process.WaitForExitAsync(cancellationToken);

            var stdOut = await stdOutTask;
            var stdErr = await stdErrTask;

            if (process.ExitCode is not 0 and not 3)
            {
                var message = string.IsNullOrWhiteSpace(stdErr) ? stdOut : stdErr;
                throw new InvalidOperationException($"{failurePrefix}: {message.Trim()}");
            }

            ReplaceOutputIfNeeded(effectiveOutputPath, outputPath);

            var warningMessage = process.ExitCode == 3 || ContainsWarning(stdOut) || ContainsWarning(stdErr)
                ? "PDF 내부 구조 경고가 있었지만 결과 파일은 생성되었습니다."
                : null;

            return new PdfMergeResult(outputPath, warningMessage);
        }
        finally
        {
            if (!string.Equals(effectiveOutputPath, outputPath, StringComparison.OrdinalIgnoreCase))
            {
                TryDelete(effectiveOutputPath);
            }
        }
    }

    private static string CreateEffectiveOutputPath(string outputPath)
    {
        if (!File.Exists(outputPath))
        {
            return outputPath;
        }

        var outputFolder = Path.GetDirectoryName(outputPath);
        var extension = Path.GetExtension(outputPath);
        return Path.Combine(
            string.IsNullOrWhiteSpace(outputFolder) ? Path.GetTempPath() : outputFolder,
            $"PdfMergeTool-overwrite-{Guid.NewGuid():N}{extension}");
    }

    private static void ReplaceOutputIfNeeded(string effectiveOutputPath, string outputPath)
    {
        if (string.Equals(effectiveOutputPath, outputPath, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        File.Copy(effectiveOutputPath, outputPath, overwrite: true);
    }

    private static bool ContainsWarning(string text)
    {
        return text.Contains("WARNING", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("warnings", StringComparison.OrdinalIgnoreCase);
    }

    private static string ResolveQpdfPath()
    {
        var appFolder = AppContext.BaseDirectory;
        var bundled = Path.Combine(appFolder, "tools", "qpdf", "qpdf.exe");
        if (File.Exists(bundled))
        {
            return bundled;
        }

        var current = new DirectoryInfo(appFolder);
        while (current is not null)
        {
            var devTool = Path.Combine(current.FullName, ".tools", "qpdf", "qpdf-12.3.2-msvc64", "bin", "qpdf.exe");
            if (File.Exists(devTool))
            {
                return devTool;
            }

            current = current.Parent;
        }

        return "qpdf.exe";
    }

    private static int NormalizeRotation(int rotation)
    {
        var normalized = rotation % 360;
        if (normalized < 0)
        {
            normalized += 360;
        }

        return normalized;
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Temporary files are best-effort cleanup.
        }
    }

    private static void WriteSingleA4PagePdf(string outputPath, PdfImage? image)
    {
        var outputFolder = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        var content = image is null ? string.Empty : BuildImageContent(image);
        var contentBytes = Encoding.ASCII.GetBytes(content);
        using var stream = File.Create(outputPath);
        var offsets = new List<long> { 0 };

        WriteAscii(stream, "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");
        WriteObject(stream, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(stream, offsets, 2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>");

        var resources = image is null
            ? "<< >>"
            : "<< /XObject << /Im1 5 0 R >> >>";
        WriteObject(
            stream,
            offsets,
            3,
            string.Create(
                CultureInfo.InvariantCulture,
                $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {A4WidthPoints:0.##} {A4HeightPoints:0.##}] /Resources {resources} /Contents 4 0 R >>"));

        offsets.Add(stream.Position);
        WriteAscii(stream, $"4 0 obj\n<< /Length {contentBytes.Length} >>\nstream\n");
        stream.Write(contentBytes);
        WriteAscii(stream, "\nendstream\nendobj\n");

        if (image is not null)
        {
            offsets.Add(stream.Position);
            WriteAscii(
                stream,
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"5 0 obj\n<< /Type /XObject /Subtype /Image /Width {image.Width} /Height {image.Height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {image.JpegBytes.Length} >>\nstream\n"));
            stream.Write(image.JpegBytes);
            WriteAscii(stream, "\nendstream\nendobj\n");
        }

        var startXref = stream.Position;
        WriteAscii(stream, $"xref\n0 {offsets.Count}\n");
        WriteAscii(stream, "0000000000 65535 f \n");
        for (var index = 1; index < offsets.Count; index++)
        {
            WriteAscii(stream, $"{offsets[index]:0000000000} 00000 n \n");
        }

        WriteAscii(stream, $"trailer\n<< /Size {offsets.Count} /Root 1 0 R >>\nstartxref\n{startXref}\n%%EOF\n");
    }

    private static string BuildImageContent(PdfImage image)
    {
        var fitWidth = A4WidthPoints - (ImagePageMarginPoints * 2);
        var fitHeight = A4HeightPoints - (ImagePageMarginPoints * 2);
        var scale = Math.Min(fitWidth / image.Width, fitHeight / image.Height);
        var drawWidth = image.Width * scale;
        var drawHeight = image.Height * scale;
        var x = (A4WidthPoints - drawWidth) / 2;
        var y = (A4HeightPoints - drawHeight) / 2;

        return string.Create(
            CultureInfo.InvariantCulture,
            $"q\n{drawWidth:0.###} 0 0 {drawHeight:0.###} {x:0.###} {y:0.###} cm\n/Im1 Do\nQ\n");
    }

    private static void WriteObject(Stream stream, List<long> offsets, int number, string body)
    {
        offsets.Add(stream.Position);
        WriteAscii(stream, $"{number} 0 obj\n{body}\nendobj\n");
    }

    private static void WriteAscii(Stream stream, string text)
    {
        var bytes = Encoding.ASCII.GetBytes(text);
        stream.Write(bytes);
    }

    private sealed record PdfImage(byte[] JpegBytes, int Width, int Height);
}

public sealed record PdfMergeResult(string OutputPath, string? WarningMessage);

public sealed record PdfPageTransform(int PageNumber, int Rotation);

public sealed record PdfImagePage(byte[] JpegBytes, int Width, int Height);
