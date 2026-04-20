using System.Diagnostics;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;

namespace PdfMergeTool.Services;

internal static class UpdateService
{
    private const string LatestReleaseApiUrl = "https://api.github.com/repos/purius/asta_pdf/releases/latest";
    private const string InstallerAssetName = "PdfMergeToolSetup.exe";

    public static string CurrentVersionText
    {
        get
        {
            var informationalVersion = Assembly.GetExecutingAssembly()
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                ?.InformationalVersion;

            return NormalizeVersionText(informationalVersion) ??
                   Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ??
                   "0.0.0";
        }
    }

    public static async Task<UpdateCheckResult> CheckForUpdatesAsync(CancellationToken cancellationToken = default)
    {
        using var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("PdfMergeTool");
        httpClient.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");

        using var response = await httpClient.GetAsync(LatestReleaseApiUrl, cancellationToken);
        response.EnsureSuccessStatusCode();

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        var root = document.RootElement;

        var latestTag = root.GetProperty("tag_name").GetString() ?? string.Empty;
        var latestVersionText = NormalizeVersionText(latestTag) ?? latestTag;
        var releaseUrl = GetUri(root, "html_url") ?? new Uri("https://github.com/purius/asta_pdf/releases/latest");
        var installerUrl = FindInstallerUrl(root) ?? releaseUrl;

        var currentVersion = ParseVersion(CurrentVersionText);
        var latestVersion = ParseVersion(latestVersionText);

        return new UpdateCheckResult(
            CurrentVersionText,
            latestVersionText,
            latestVersion > currentVersion,
            releaseUrl,
            installerUrl);
    }

    public static void OpenUpdatePage(UpdateCheckResult result)
    {
        Process.Start(new ProcessStartInfo(result.ReleaseUrl.AbsoluteUri)
        {
            UseShellExecute = true
        });
    }

    private static Uri? FindInstallerUrl(JsonElement root)
    {
        if (!root.TryGetProperty("assets", out var assets) || assets.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        foreach (var asset in assets.EnumerateArray())
        {
            var name = asset.TryGetProperty("name", out var nameElement)
                ? nameElement.GetString()
                : null;

            if (!string.Equals(name, InstallerAssetName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            return GetUri(asset, "browser_download_url");
        }

        return null;
    }

    private static Uri? GetUri(JsonElement root, string propertyName)
    {
        if (!root.TryGetProperty(propertyName, out var value) ||
            value.ValueKind != JsonValueKind.String ||
            !Uri.TryCreate(value.GetString(), UriKind.Absolute, out var uri))
        {
            return null;
        }

        return uri;
    }

    private static Version ParseVersion(string value)
    {
        var normalized = NormalizeVersionText(value) ?? "0.0.0";
        return Version.TryParse(normalized, out var version) ? version : new Version(0, 0, 0);
    }

    private static string? NormalizeVersionText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var normalized = value.Trim().TrimStart('v', 'V');
        var suffixStart = normalized.IndexOfAny(['-', '+']);
        if (suffixStart >= 0)
        {
            normalized = normalized[..suffixStart];
        }

        return normalized.Trim();
    }
}

internal sealed record UpdateCheckResult(
    string CurrentVersionText,
    string LatestVersionText,
    bool IsUpdateAvailable,
    Uri ReleaseUrl,
    Uri InstallerUrl);
