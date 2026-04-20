using System.Diagnostics;
using System.Drawing;
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
    private const string AppVersion = "1.0.7";
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

        using var progress = quiet ? null : new InstallerProgressForm($"{AppName} 설치");
        progress?.Show();
        progress?.SetStep("설치 준비 중...", 0);

        try
        {
            progress?.SetStep("실행 중인 프로그램을 종료하는 중...", 5);
            StopInstalledApp();

            progress?.SetStep("기존 설치 파일을 정리하는 중...", 12);
            ResetInstallDirectory(progress);

            progress?.SetStep("설치 폴더를 준비하는 중...", 18);
            Directory.CreateDirectory(InstallDirectory);

            ExtractPayload(InstallDirectory, progress, 20, 58);
            CopyInstallerToInstallDirectory(progress);
        }
        catch (Exception ex)
        {
            progress?.Fail(ex.Message);
            throw;
        }

        var appExe = Path.Combine(InstallDirectory, AppExeName);
        var setupExe = Path.Combine(InstallDirectory, SetupExeName);

        progress?.SetStep("우클릭 메뉴를 등록하는 중...", 82);
        RegisterPdfContextMenu(appExe);
        progress?.SetStep("PDF 기본 프로그램 정보를 등록하는 중...", 87);
        RegisterPdfFileAssociation(appExe);
        progress?.SetStep("제거 정보를 등록하는 중...", 92);
        RegisterUninstaller(appExe, setupExe);
        progress?.SetStep("바로가기를 만드는 중...", 95);
        CreateShortcuts(appExe, setupExe);
        progress?.SetStep("Windows 파일 연결을 새로고침하는 중...", 98);
        RefreshShellAssociations();
        progress?.SetStep("설치가 완료되었습니다.", 100);
        progress?.Detail($"설치 위치: {InstallDirectory}");
        progress?.Close();

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

    private static void ExtractPayload(string destination, InstallerProgressForm? progress, int basePercent, int spanPercent)
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(PayloadResourceName)
            ?? throw new InvalidOperationException("설치 파일에 앱 payload가 포함되어 있지 않습니다.");
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var destinationFullPath = Path.GetFullPath(destination);
        var fileCount = archive.Entries.Count(entry => !string.IsNullOrEmpty(entry.Name));
        var copied = 0;
        progress?.SetStep("프로그램 파일을 복사하는 중...", basePercent);
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
            copied++;
            var percent = basePercent + (int)Math.Round(spanPercent * (double)copied / Math.Max(1, fileCount));
            progress?.SetStep($"프로그램 파일 복사 중... ({copied}/{fileCount})", percent);
            progress?.Detail(entry.FullName);
        }
    }

    private static void CopyInstallerToInstallDirectory(InstallerProgressForm? progress)
    {
        progress?.SetStep("설치 프로그램을 복사하는 중...", 80);
        var currentExe = Environment.ProcessPath
            ?? throw new InvalidOperationException("현재 설치 프로그램 경로를 확인할 수 없습니다.");
        var target = Path.Combine(InstallDirectory, SetupExeName);

        if (!string.Equals(Path.GetFullPath(currentExe), Path.GetFullPath(target), StringComparison.OrdinalIgnoreCase))
        {
            progress?.Detail(Path.GetFileName(target));
            File.Copy(currentExe, target, overwrite: true);
        }
    }

    private static void ResetInstallDirectory(InstallerProgressForm? progress = null)
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

        progress?.Detail(fullInstallDirectory);
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

    private sealed class InstallerProgressForm : System.Windows.Forms.Form
    {
        private readonly System.Windows.Forms.Label _stepLabel = new();
        private readonly System.Windows.Forms.ProgressBar _progressBar = new();
        private readonly System.Windows.Forms.TextBox _detailsBox = new();

        public InstallerProgressForm(string title)
        {
            Text = title;
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(560, 330);
            TopMost = true;

            var headerLabel = new System.Windows.Forms.Label
            {
                Text = "설치를 진행하고 있습니다.",
                Font = new Font(Font.FontFamily, 12, FontStyle.Bold),
                AutoSize = false,
                Location = new Point(18, 16),
                Size = new Size(524, 26)
            };

            _stepLabel.AutoSize = false;
            _stepLabel.Location = new Point(18, 52);
            _stepLabel.Size = new Size(524, 24);
            _stepLabel.Text = "준비 중...";

            _progressBar.Location = new Point(18, 84);
            _progressBar.Size = new Size(524, 22);
            _progressBar.Minimum = 0;
            _progressBar.Maximum = 100;

            _detailsBox.Location = new Point(18, 120);
            _detailsBox.Size = new Size(524, 178);
            _detailsBox.Multiline = true;
            _detailsBox.ReadOnly = true;
            _detailsBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            _detailsBox.BackColor = Color.White;

            Controls.Add(headerLabel);
            Controls.Add(_stepLabel);
            Controls.Add(_progressBar);
            Controls.Add(_detailsBox);
        }

        public void SetStep(string message, int percent)
        {
            if (IsDisposed)
            {
                return;
            }

            _stepLabel.Text = message;
            _progressBar.Value = Math.Clamp(percent, _progressBar.Minimum, _progressBar.Maximum);
            Detail(message);
        }

        public void Detail(string message)
        {
            if (IsDisposed || string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            _detailsBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
            _detailsBox.SelectionStart = _detailsBox.TextLength;
            _detailsBox.ScrollToCaret();
            Application.DoEvents();
        }

        public void Fail(string message)
        {
            if (IsDisposed)
            {
                return;
            }

            _stepLabel.Text = "설치 중 오류가 발생했습니다.";
            Detail($"오류: {message}");
            Application.DoEvents();
        }
    }

    [DllImport("shell32.dll")]
    private static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
