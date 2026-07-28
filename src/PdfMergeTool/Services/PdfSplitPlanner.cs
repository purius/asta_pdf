using System.IO;

namespace PdfMergeTool.Services;

public static class PdfSplitPlanner
{
    public static PdfSplitPlan CreateIntervalPlan(string inputPath, int pageCount, int interval)
    {
        ValidatePageCount(pageCount);
        if (interval < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(interval), "분할 페이지 수는 1 이상이어야 합니다.");
        }

        var outputFolder = AllocateOutputFolder(inputPath);
        var baseName = GetBaseName(inputPath);
        var outputs = new List<PdfSplitOutput>();

        for (var startPage = 1; startPage <= pageCount; startPage += interval)
        {
            var endPage = Math.Min(startPage + interval - 1, pageCount);
            var outputIndex = outputs.Count + 1;
            outputs.Add(new PdfSplitOutput(
                PdfSplitKind.Interval,
                Path.Combine(outputFolder, $"{baseName}_{outputIndex:000}.pdf"),
                FormatPageRange(startPage, endPage),
                CreatePageTransforms(startPage, endPage)));
        }

        return new PdfSplitPlan(inputPath, outputFolder, outputs);
    }

    public static PdfSplitPlan CreateParityPlan(string inputPath, int pageCount)
    {
        ValidatePageCount(pageCount);

        var outputFolder = AllocateOutputFolder(inputPath);
        var baseName = GetBaseName(inputPath);
        var outputs = new List<PdfSplitOutput>();

        AddParityOutput(outputs, outputFolder, baseName, PdfSplitKind.Odd, pageCount, 1, "홀수");
        AddParityOutput(outputs, outputFolder, baseName, PdfSplitKind.Even, pageCount, 2, "짝수");

        return new PdfSplitPlan(inputPath, outputFolder, outputs);
    }

    private static void AddParityOutput(
        List<PdfSplitOutput> outputs,
        string outputFolder,
        string baseName,
        PdfSplitKind kind,
        int pageCount,
        int firstPage,
        string label)
    {
        var pages = Enumerable.Range(firstPage, pageCount - firstPage + 1)
            .Where(page => page <= pageCount && page % 2 == firstPage % 2)
            .Select(page => new PdfPageTransform(page, 0))
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

    private static string FormatPageRange(int startPage, int endPage)
    {
        return startPage == endPage ? startPage.ToString() : $"{startPage}-{endPage}";
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
