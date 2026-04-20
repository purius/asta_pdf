using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Application = System.Windows.Forms.Application;
using MessageBox = System.Windows.Forms.MessageBox;

namespace PdfMergeTool.Installer;

internal static class Program
{
    private const uint ShcneAssocChanged = 0x08000000;
    private const uint ShcnfIdList = 0x0000;
    private const string ProductId = "PdfMergeTool";
    private const string AppName = "PDF 뷰어";
    private const string Publisher = "PdfMergeTool";
    private const string AppVersion = "1.0.4";
    private const string AppExeName = "PdfMergeTool.exe";
    private const string SetupExeName = "PdfMergeToolSetup.exe";
    private const string PayloadResourceName = "PdfMergeToolPayload.zip";
    private const string PdfProgId = "PdfMergeTool.Pdf";
    private static readonly string[] PdfContextMenuPaths =
    [
        @"Software\Classes\SystemFileAssociations\.pdf\shell\PdfMergeTool",
        @"Software\Classes\.pdf\shell\PdfMergeTool",
        @$"Software\Classes\{PdfProgId}\shell\PdfMergeTool"
    ];

    private static readonly string InstallDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Programs",
        ProductId);

    [STAThread]
    private static int Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        var quiet = args.Any(arg => string.Equals(arg, "--quiet", StringComparison.OrdinalIgnoreCase));
        var uninstall = args.Any(arg => string.Equals(arg, "--uninstall", StringComparison.OrdinalIgnoreCase));

        try
        {
            if (uninstall)
            {
                return RunUninstall(quiet);
            }

            return RunInstall(quiet);
        }
        catch (Exception ex)
        {
            if (!quiet)
            {
                MessageBox.Show(ex.Message, $"{AppName} 설치 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return 1;
        }
    }

    private static int RunInstall(bool quiet)
    {
        if (!quiet)
        {
            var result = MessageBox.Show(
                $"{AppName}를 설치합니다.\n\n설치 위치:\n{InstallDirectory}",
                $"{AppName} 설치",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information);

            if (result != DialogResult.OK)
            {
                return 2;
            }
        }

        StopInstalledApp();
        ResetInstallDirectory();
        Directory.CreateDirectory(InstallDirectory);
        ExtractPayload(InstallDirectory);
        CopyInstallerToInstallDirectory();

        var appExe = Path.Combine(InstallDirectory, AppExeName);
        var setupExe = Path.Combine(InstallDirectory, SetupExeName);

        RegisterPdfContextMenu(appExe);
        RegisterPdfFileAssociation(appExe);
        RegisterUninstaller(appExe, setupExe);
        CreateShortcuts(appExe, setupExe);
        RefreshShellAssociations();

        if (!quiet)
        {
            var result = MessageBox.Show(
                "설치가 완료되었습니다.\n\n지금 PDF 뷰어를 실행할까요?",
                $"{AppName} 설치",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
            {
                Process.Start(new ProcessStartInfo(appExe) { UseShellExecute = true });
            }
        }

        return 0;
    }

    private static int RunUninstall(bool quiet)
    {
        if (!quiet)
        {
            var result = MessageBox.Show(
                $"{AppName}를 제거합니다.",
                $"{AppName} 제거",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning);

            if (result != DialogResult.OK)
            {
                return 2;
            }
        }

        StopInstalledApp();
        RemoveShortcuts();
        RemovePdfContextMenu();
        RemovePdfFileAssociation();
        RemoveUninstallerRegistration();
        RefreshShellAssociations();

        if (!quiet)
        {
            MessageBox.Show("제거가 완료되었습니다.", $"{AppName} 제거", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        DeleteInstallDirectoryAfterExit();
        return 0;
    }

    private static void ExtractPayload(string destination)
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(PayloadResourceName)
            ?? throw new InvalidOperationException("설치 파일에 앱 payload가 포함되어 있지 않습니다.");
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var destinationFullPath = Path.GetFullPath(destination);
        foreach (var entry in archive.Entries)
        {
            var targetPath = Path.GetFullPath(Path.Combine(destinationFullPath, entry.FullName));
            if (!targetPath.StartsWith(destinationFullPath, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("잘못된 설치 파일 항목이 감지되었습니다.");
            }

            if (string.IsNullOrEmpty(entry.Name))
            {
                Directory.CreateDirectory(targetPath);
                continue;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(targetPath)!);
            entry.ExtractToFile(targetPath, overwrite: true);
        }
    }

    private static void CopyInstallerToInstallDirectory()
    {
        var currentExe = Environment.ProcessPath
            ?? throw new InvalidOperationException("현재 설치 프로그램 경로를 확인할 수 없습니다.");
        var target = Path.Combine(InstallDirectory, SetupExeName);

        if (!string.Equals(Path.GetFullPath(currentExe), Path.GetFullPath(target), StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(currentExe, target, overwrite: true);
        }
    }

    private static void ResetInstallDirectory()
    {
        if (!Directory.Exists(InstallDirectory))
        {
            return;
        }

        var fullInstallDirectory = Path.GetFullPath(InstallDirectory);
        var allowedRoot = Path.GetFullPath(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Programs"));

        if (!fullInstallDirectory.StartsWith(allowedRoot, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(fullInstallDirectory, allowedRoot, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("설치 폴더 경로가 안전하지 않아 초기화하지 않았습니다.");
        }

        DeleteDirectoryWithRetry(fullInstallDirectory);
    }

    private static void RegisterPdfContextMenu(string appExe)
    {
        foreach (var contextMenuPath in PdfContextMenuPaths)
        {
            using var verbKey = Registry.CurrentUser.CreateSubKey(contextMenuPath);
            verbKey.SetValue(null, "PDF 통합...");
            verbKey.SetValue("Icon", appExe);
            verbKey.SetValue("MultiSelectModel", "Player");

            using var commandKey = verbKey.CreateSubKey("command");
            commandKey.SetValue(null, $"\"{appExe}\" --merge \"%1\"");
        }
    }

    private static void RemovePdfContextMenu()
    {
        foreach (var contextMenuPath in PdfContextMenuPaths)
        {
            Registry.CurrentUser.DeleteSubKeyTree(contextMenuPath, throwOnMissingSubKey: false);
        }
    }

    private static void RegisterPdfFileAssociation(string appExe)
    {
        using (var progIdKey = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{PdfProgId}"))
        {
            progIdKey.SetValue(null, "PDF 문서");
            progIdKey.SetValue("FriendlyTypeName", "PDF 문서");
            progIdKey.SetValue("EditFlags", 0, RegistryValueKind.DWord);
        }

        using (var iconKey = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{PdfProgId}\DefaultIcon"))
        {
            iconKey.SetValue(null, $"\"{appExe}\",0");
        }

        using (var openCommandKey = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{PdfProgId}\shell\open\command"))
        {
            openCommandKey.SetValue(null, $"\"{appExe}\" \"%1\"");
        }

        using (var appKey = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Applications\PdfMergeTool.exe"))
        {
            appKey.SetValue("FriendlyAppName", AppName);
        }

        using (var appIconKey = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Applications\PdfMergeTool.exe\DefaultIcon"))
        {
            appIconKey.SetValue(null, $"\"{appExe}\",0");
        }

        using (var appOpenCommandKey = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Applications\PdfMergeTool.exe\shell\open\command"))
        {
            appOpenCommandKey.SetValue(null, $"\"{appExe}\" \"%1\"");
        }

        using (var extensionKey = Registry.CurrentUser.CreateSubKey(@"Software\Classes\.pdf"))
        {
            extensionKey.SetValue(null, PdfProgId);
            extensionKey.SetValue("Content Type", "application/pdf");
            extensionKey.SetValue("PerceivedType", "document");
        }

        using (var openWithProgIdsKey = Registry.CurrentUser.CreateSubKey(@"Software\Classes\.pdf\OpenWithProgids"))
        {
            openWithProgIdsKey.SetValue(PdfProgId, Array.Empty<byte>(), RegistryValueKind.Binary);
        }

        using (var explorerOpenWithProgIdsKey = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\OpenWithProgids"))
        {
            explorerOpenWithProgIdsKey.SetValue(PdfProgId, Array.Empty<byte>(), RegistryValueKind.Binary);
        }

        using (var capabilitiesKey = Registry.CurrentUser.CreateSubKey(@"Software\PdfMergeTool\Capabilities"))
        {
            capabilitiesKey.SetValue("ApplicationName", AppName);
            capabilitiesKey.SetValue("ApplicationDescription", "PDF 보기, 페이지 정리, 통합 도구");
        }

        using (var fileAssociationsKey = Registry.CurrentUser.CreateSubKey(@"Software\PdfMergeTool\Capabilities\FileAssociations"))
        {
            fileAssociationsKey.SetValue(".pdf", PdfProgId);
        }

        using (var registeredApplicationsKey = Registry.CurrentUser.CreateSubKey(@"Software\RegisteredApplications"))
        {
            registeredApplicationsKey.SetValue(ProductId, @"Software\PdfMergeTool\Capabilities");
        }
    }

    private static void RemovePdfFileAssociation()
    {
        RemovePdfProgIdFromOpenWith(@"Software\Classes\.pdf\OpenWithProgids");
        RemovePdfProgIdFromOpenWith(@"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\OpenWithProgids");
        RemoveDefaultPdfProgIdIfOwned();

        Registry.CurrentUser.DeleteSubKeyTree($@"Software\Classes\{PdfProgId}", throwOnMissingSubKey: false);
        Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\Applications\PdfMergeTool.exe", throwOnMissingSubKey: false);
        Registry.CurrentUser.DeleteSubKeyTree(@"Software\PdfMergeTool\Capabilities", throwOnMissingSubKey: false);
        Registry.CurrentUser.DeleteSubKeyTree(@"Software\PdfMergeTool", throwOnMissingSubKey: false);

        using var registeredApplicationsKey = Registry.CurrentUser.OpenSubKey(@"Software\RegisteredApplications", writable: true);
        registeredApplicationsKey?.DeleteValue(ProductId, throwOnMissingValue: false);
    }

    private static void RemovePdfProgIdFromOpenWith(string subKeyPath)
    {
        using var key = Registry.CurrentUser.OpenSubKey(subKeyPath, writable: true);
        key?.DeleteValue(PdfProgId, throwOnMissingValue: false);
    }

    private static void RemoveDefaultPdfProgIdIfOwned()
    {
        using var extensionKey = Registry.CurrentUser.OpenSubKey(@"Software\Classes\.pdf", writable: true);
        if (extensionKey is not null &&
            string.Equals(extensionKey.GetValue(null) as string, PdfProgId, StringComparison.OrdinalIgnoreCase))
        {
            extensionKey.DeleteValue(string.Empty, throwOnMissingValue: false);
        }
    }

    private static void RegisterUninstaller(string appExe, string setupExe)
    {
        using var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\PdfMergeTool");
        key.SetValue("DisplayName", AppName);
        key.SetValue("DisplayVersion", AppVersion);
        key.SetValue("Publisher", Publisher);
        key.SetValue("InstallLocation", InstallDirectory);
        key.SetValue("DisplayIcon", appExe);
        key.SetValue("UninstallString", $"\"{setupExe}\" --uninstall");
        key.SetValue("QuietUninstallString", $"\"{setupExe}\" --uninstall --quiet");
        key.SetValue("NoModify", 1, RegistryValueKind.DWord);
        key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
        key.SetValue("EstimatedSize", GetDirectorySizeInKb(InstallDirectory), RegistryValueKind.DWord);
    }

    private static void RemoveUninstallerRegistration()
    {
        Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\PdfMergeTool", throwOnMissingSubKey: false);
    }

    private static void CreateShortcuts(string appExe, string setupExe)
    {
        var startMenuDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
            "Programs",
            AppName);
        Directory.CreateDirectory(startMenuDir);

        CreateShortcut(Path.Combine(startMenuDir, $"{AppName}.lnk"), appExe, "", InstallDirectory, appExe);
        CreateShortcut(Path.Combine(startMenuDir, $"{AppName} 제거.lnk"), setupExe, "--uninstall", InstallDirectory, appExe);

        var desktopShortcut = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            $"{AppName}.lnk");
        CreateShortcut(desktopShortcut, appExe, "", InstallDirectory, appExe);
    }

    private static void RemoveShortcuts()
    {
        var startMenuDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
            "Programs",
            AppName);
        if (Directory.Exists(startMenuDir))
        {
            Directory.Delete(startMenuDir, recursive: true);
        }

        var desktopShortcut = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            $"{AppName}.lnk");
        if (File.Exists(desktopShortcut))
        {
            File.Delete(desktopShortcut);
        }
    }

    private static void CreateShortcut(string shortcutPath, string targetPath, string arguments, string workingDirectory, string iconPath)
    {
        var shellType = Type.GetTypeFromProgID("WScript.Shell")
            ?? throw new InvalidOperationException("바로가기를 만들 수 없습니다.");
        var shell = Activator.CreateInstance(shellType)
            ?? throw new InvalidOperationException("바로가기를 만들 수 없습니다.");

        try
        {
            dynamic shortcut = shellType.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, shell, [shortcutPath])!;
            shortcut.TargetPath = targetPath;
            shortcut.Arguments = arguments;
            shortcut.WorkingDirectory = workingDirectory;
            shortcut.IconLocation = iconPath;
            shortcut.Save();
        }
        finally
        {
            if (Marshal.IsComObject(shell))
            {
                Marshal.FinalReleaseComObject(shell);
            }
        }
    }

    private static void StopInstalledApp()
    {
        foreach (var process in Process.GetProcessesByName("PdfMergeTool"))
        {
            try
            {
                process.Kill(entireProcessTree: true);
                process.WaitForExit(10000);
            }
            catch (InvalidOperationException)
            {
                // The process has already exited.
            }
            catch
            {
                // If Windows denies process access, the later directory delete will fail with a clear error.
            }
        }
    }

    private static void DeleteDirectoryWithRetry(string path)
    {
        for (var attempt = 1; attempt <= 8; attempt++)
        {
            try
            {
                Directory.Delete(path, recursive: true);
                return;
            }
            catch (Exception) when (attempt < 8)
            {
                StopInstalledApp();
                Thread.Sleep(500 * attempt);
            }
        }

        Directory.Delete(path, recursive: true);
    }

    private static void DeleteInstallDirectoryAfterExit()
    {
        if (!Directory.Exists(InstallDirectory))
        {
            return;
        }

        var fullInstallDirectory = Path.GetFullPath(InstallDirectory);
        var allowedRoot = Path.GetFullPath(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Programs"));

        if (!fullInstallDirectory.StartsWith(allowedRoot, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(fullInstallDirectory, allowedRoot, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("설치 폴더 경로가 안전하지 않아 삭제하지 않았습니다.");
        }

        var arguments = $"/c timeout /t 2 /nobreak > nul & rmdir /s /q \"{fullInstallDirectory}\"";
        Process.Start(new ProcessStartInfo("cmd.exe", arguments)
        {
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        });
    }

    private static int GetDirectorySizeInKb(string path)
    {
        if (!Directory.Exists(path))
        {
            return 0;
        }

        var size = Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories)
            .Sum(file => new FileInfo(file).Length);
        return (int)Math.Max(1, size / 1024);
    }

    private static void RefreshShellAssociations()
    {
        SHChangeNotify(ShcneAssocChanged, ShcnfIdList, IntPtr.Zero, IntPtr.Zero);
    }

    [DllImport("shell32.dll")]
    private static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
