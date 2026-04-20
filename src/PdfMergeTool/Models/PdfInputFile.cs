using System.IO;
using System.Windows.Media;
using PdfMergeTool.Services;

namespace PdfMergeTool.Models;

public sealed class PdfInputFile
{
    public PdfInputFile(string path)
    {
        Path = System.IO.Path.GetFullPath(path);
        ThumbnailSource = ShellThumbnailService.TryLoadThumbnail(Path, 56, 74);
    }

    public string Path { get; }

    public string FileName => System.IO.Path.GetFileName(Path);

    public string Folder => System.IO.Path.GetDirectoryName(Path) ?? string.Empty;

    public string PageRange { get; set; } = string.Empty;

    public ImageSource? ThumbnailSource { get; }
}
