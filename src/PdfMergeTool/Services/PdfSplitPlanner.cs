using System.IO;

namespace PdfMergeTool.Services;

public static class PdfSplitPlanner
{
    public static PdfSplitPlan CreateIntervalPlan(string inputPath, int pageCount, int interval)
    {
        ValidatePageCount(pageCount);
        return CreateIntervalPlan(inputPath, CreatePageTransforms(1, pageCount), interval);
    }

    public static PdfSplitPlan CreateIntervalPlan(
        string inputPath,
        IReadOnlyList<PdfPageTransform> pages,
        int interval,
        string? outputPathReference = null)
    {
        ValidatePages(pages);
        if (interval < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(interval), "분할 페이지 수는 1 이상이어야 합니다.");
        }

        var outputReference = outputPathReference ?? inputPath;
        var outputFolder = AllocateOutputFolder(outputReference);
        var baseName = GetBaseName(outputReference);
        var outputs = new List<PdfSplitOutput>();

        for (var startIndex = 0; startIndex < pages.Count; startIndex += interval)
        {
            var outputPages = pages.Skip(startIndex).Take(interval).ToList();
            var outputIndex = outputs.Count + 1;
            outputs.Add(new PdfSplitOutput(
                PdfSplitKind.Interval,
                Path.Combine(outputFolder, $"{baseName}_{outputIndex:000}.pdf"),
                FormatPageRange(outputPages),
                outputPages));
        }

        return new PdfSplitPlan(inputPath, outputFolder, outputs);
    }

    public static PdfSplitPlan CreateParityPlan(string inputPath, int pageCount)
    {
        ValidatePageCount(pageCount);
        return CreateParityPlan(inputPath, CreatePageTransforms(1, pageCount));
    }

    public static PdfSplitPlan CreateParityPlan(
        string inputPath,
        IReadOnlyList<PdfPageTransform> pages,
        string? outputPathReference = null)
    {
        ValidatePages(pages);

        var outputReference = outputPathReference ?? inputPath;
        var outputFolder = AllocateOutputFolder(outputReference);
        var baseName = GetBaseName(outputReference);
        var outputs = new List<PdfSplitOutput>();

        AddParityOutput(outputs, outputFolder, baseName, PdfSplitKind.Odd, pages, 1, "홀수");
        AddParityOutput(outputs, outputFolder, baseName, PdfSplitKind.Even, pages, 0, "짝수");

        return new PdfSplitPlan(inputPath, outputFolder, outputs);
    }

    private static void AddParityOutput(
        List<PdfSplitOutput> outputs,
        string outputFolder,
        string baseName,
        PdfSplitKind kind,
        IReadOnlyList<PdfPageTransform> sourcePages,
        int parity,
        string label)
    {
        var pages = sourcePages
            .Where(page => page.PageNumber % 2 == parity)
            .ToList();

        if (pages.Count == 0)
        {
            return;
        }

        outputs.Add(new PdfSplitOutput(
            kind,
            Path.Combine(outputFolder, $"{baseName}_{label}.pdf"),
            string.Join(",", pages.Select(page => page.PageNumber)),
            pages));
    }

    private static string AllocateOutputFolder(string inputPath)
    {
        var sourceFolder = Path.GetDirectoryName(inputPath);
        if (string.IsNullOrWhiteSpace(sourceFolder))
        {
            sourceFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        var baseFolder = Path.Combine(sourceFolder, $"{GetBaseName(inputPath)}_분할");
        if (!Directory.Exists(baseFolder))
        {
            return baseFolder;
        }

        for (var suffix = 2; ; suffix++)
        {
            var candidate = $"{baseFolder} ({suffix})";
            if (!Directory.Exists(candidate))
            {
                return candidate;
            }
        }
    }

    private static IReadOnlyList<PdfPageTransform> CreatePageTransforms(int startPage, int endPage)
    {
        return Enumerable.Range(startPage, endPage - startPage + 1)
            .Select(page => new PdfPageTransform(page, 0))
            .ToList();
    }

    private static string FormatPageRange(IReadOnlyList<PdfPageTransform> pages)
    {
        if (pages.Count == 1)
        {
            return pages[0].PageNumber.ToString();
        }

        var firstPage = pages[0].PageNumber;
        var isContiguousRange = pages
            .Select((page, index) => page.PageNumber == firstPage + index)
            .All(isContiguous => isContiguous);
        if (isContiguousRange)
        {
            return $"{firstPage}-{pages[^1].PageNumber}";
        }

        return string.Join(",", pages.Select(page => page.PageNumber));
    }

    private static string GetBaseName(string inputPath)
    {
        return Path.GetFileNameWithoutExtension(inputPath) ?? "분할";
    }

    private static void ValidatePageCount(int pageCount)
    {
        if (pageCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(pageCount), "PDF 페이지 수는 1 이상이어야 합니다.");
        }
    }

    private static void ValidatePages(IReadOnlyList<PdfPageTransform> pages)
    {
        ArgumentNullException.ThrowIfNull(pages);
        if (pages.Count < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(pages), "분할할 페이지가 없습니다.");
        }
    }
}

public enum PdfSplitMode
{
    PageByPage,
    Interval,
    Parity
}

public enum PdfSplitKind
{
    Interval,
    Odd,
    Even
}

public sealed record PdfSplitPlan(
    string InputPath,
    string OutputFolder,
    IReadOnlyList<PdfSplitOutput> Outputs);

public sealed record PdfSplitOutput(
    PdfSplitKind Kind,
    string OutputPath,
    string PageRange,
    IReadOnlyList<PdfPageTransform> Pages);

public sealed record PdfPageTransform(int PageNumber, int Rotation);
