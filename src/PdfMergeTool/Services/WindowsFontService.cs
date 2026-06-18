using System.IO;
using Microsoft.Win32;

namespace PdfMergeTool.Services;

internal sealed record WindowsFontInfo(string Name, string Path);

internal static class WindowsFontService
{
    private const string FontsRegistryPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";
    private static readonly string FontsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);

    public static IReadOnlyList<WindowsFontInfo> GetInstalledFonts()
    {
        var fonts = new Dictionary<string, WindowsFontInfo>(StringComparer.OrdinalIgnoreCase);
        AddRegistryFonts(fonts, Registry.LocalMachine);
        AddRegistryFonts(fonts, Registry.CurrentUser);

        return fonts.Values
            .OrderBy(font => font.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    public static IReadOnlyDictionary<string, string> ReadFontBase64(IEnumerable<string> fontNames)
    {
        var installedFonts = GetInstalledFonts()
            .ToDictionary(font => font.Name, StringComparer.OrdinalIgnoreCase);
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var fontName in fontNames.Where(name => !string.IsNullOrWhiteSpace(name)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var match = FindFont(fontName, installedFonts);
            if (match is null)
            {
                continue;
            }

            try
            {
                result[fontName] = Convert.ToBase64String(File.ReadAllBytes(match.Path));
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, $"Font file could not be read: {match.Path}");
            }
        }

        return result;
    }

    private static WindowsFontInfo? FindFont(string fontName, IReadOnlyDictionary<string, WindowsFontInfo> installedFonts)
    {
        if (installedFonts.TryGetValue(fontName, out var exact))
        {
            return exact;
        }

        return installedFonts.Values.FirstOrDefault(font =>
            font.Name.Contains(fontName, StringComparison.OrdinalIgnoreCase) ||
            fontName.Contains(font.Name, StringComparison.OrdinalIgnoreCase));
    }

    private static void AddRegistryFonts(Dictionary<string, WindowsFontInfo> fonts, RegistryKey root)
    {
        using var key = root.OpenSubKey(FontsRegistryPath);
        if (key is null)
        {
            return;
        }

        foreach (var valueName in key.GetValueNames())
        {
            var value = key.GetValue(valueName)?.ToString();
            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            var path = Path.IsPathRooted(value)
                ? value
                : Path.Combine(FontsDirectory, value);
            if (!File.Exists(path) || !IsEmbeddableFontFile(path))
            {
                continue;
            }

            foreach (var name in NormalizeFontNames(valueName))
            {
                if (!fonts.ContainsKey(name))
                {
                    fonts[name] = new WindowsFontInfo(name, path);
                }
            }
        }
    }

    private static bool IsEmbeddableFontFile(string path)
    {
        return Path.GetExtension(path).ToLowerInvariant() is ".ttf" or ".otf" or ".ttc";
    }

    private static IReadOnlyList<string> NormalizeFontNames(string registryName)
    {
        var parenIndex = registryName.IndexOf(" (", StringComparison.Ordinal);
        var normalized = (parenIndex > 0 ? registryName[..parenIndex] : registryName).Trim();
        return normalized
            .Split(" & ", StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }
}
