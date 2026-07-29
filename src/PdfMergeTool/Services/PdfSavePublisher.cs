using System.IO;

namespace PdfMergeTool.Services;

public static class PdfSavePublisher
{
    public static void Publish(string temporaryPath, string targetPath)
    {
        if (!File.Exists(temporaryPath))
        {
            throw new FileNotFoundException("저장할 임시 PDF 파일을 찾을 수 없습니다.", temporaryPath);
        }

        var targetFolder = Path.GetDirectoryName(targetPath);
        if (!string.IsNullOrWhiteSpace(targetFolder))
        {
            Directory.CreateDirectory(targetFolder);
        }

        File.Move(temporaryPath, targetPath, overwrite: true);
    }
}
