using System.IO;
using PdfMergeTool.Models;

namespace PdfMergeTool.Services;

public static class OutputPathService
{
    public static string CreateDefaultOutputPath(IReadOnlyList<PdfInputFile> files, string suffix = "_통합")
    {
        if (files.Count == 0)
        {
            throw new InvalidOperationException("At least one input file is required.");
        }

        var first = files[0];
        var baseName = System.IO.Path.GetFileNameWithoutExtension(first.FileName);
        var folder = first.Folder;
        suffix = string.IsNullOrWhiteSpace(suffix) ? "_통합" : suffix.Trim();
        suffix = suffix.StartsWith('_') ? suffix : $"_{suffix}";
        var candidate = System.IO.Path.Combine(folder, $"{baseName}{suffix}.pdf");

        if (!File.Exists(candidate))
        {
            return candidate;
        }

        for (var index = 2; index < 1000; index++)
        {
            candidate = System.IO.Path.Combine(folder, $"{baseName}{suffix} ({index}).pdf");
            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new IOException("Could not create a unique output filename.");
    }
}
