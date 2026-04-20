using System.Windows;
using PdfMergeTool.Services;

namespace PdfMergeTool;

public partial class SettingsWindow : Window
{
    private readonly AppSettings _settings;

    public SettingsWindow(AppSettings settings)
    {
        InitializeComponent();
        _settings = settings;
        LoadSettings();
        RefreshDefaultAppStatus();
    }

    public bool SettingsSaved { get; private set; }

    private void LoadSettings()
    {
        ContextMenuCheckBox.IsChecked = WindowsIntegrationService.IsPdfContextMenuRegistered();
        RecentFileLimitTextBox.Text = _settings.RecentFileLimit.ToString();
        ReorderedSuffixTextBox.Text = _settings.ReorderedSuffix;
        MergedSuffixTextBox.Text = _settings.MergedSuffix;
        A4FitSuffixTextBox.Text = _settings.A4FitSuffix;
        A4OptimizedSuffixTextBox.Text = _settings.A4OptimizedSuffix;
        AutoCleanTempFilesCheckBox.IsChecked = _settings.AutoCleanTempFiles;
    }

    private void RefreshDefaultAppStatus()
    {
        var current = WindowsIntegrationService.GetCurrentPdfDefaultApp();
        DefaultPdfAppText.Text = WindowsIntegrationService.IsPdfDefaultAppRegistered()
            ? $"현재 PDF 기본 앱: PDF 뷰어 ({current})"
            : $"현재 PDF 기본 앱: {current}";
    }

    private void OnOpenDefaultAppsClick(object sender, RoutedEventArgs e)
    {
        try
        {
            WindowsIntegrationService.OpenDefaultAppsSettings();
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, "Windows 기본 앱 설정을 열지 못했습니다.");
            MessageBox.Show(this, ex.Message, "기본 앱 설정", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnRefreshDefaultAppClick(object sender, RoutedEventArgs e)
    {
        RefreshDefaultAppStatus();
    }

    private void OnCleanTempFilesClick(object sender, RoutedEventArgs e)
    {
        var count = WindowsIntegrationService.CleanTempFiles();
        MessageBox.Show(this, $"{count}개 항목을 정리했습니다.", "임시파일 정리", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void OnOpenLogsClick(object sender, RoutedEventArgs e)
    {
        try
        {
            WindowsIntegrationService.OpenFolder(AppPaths.LogsDirectory);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, "로그 폴더를 열지 못했습니다.");
            MessageBox.Show(this, ex.Message, "로그 폴더", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnSaveClick(object sender, RoutedEventArgs e)
    {
        if (!int.TryParse(RecentFileLimitTextBox.Text, out var recentLimit))
        {
            MessageBox.Show(this, "최근 파일 개수는 숫자로 입력해주세요.", "설정", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        _settings.RecentFileLimit = Math.Clamp(recentLimit, 1, 30);
        _settings.ReorderedSuffix = NormalizeSuffix(ReorderedSuffixTextBox.Text, "_재정렬");
        _settings.MergedSuffix = NormalizeSuffix(MergedSuffixTextBox.Text, "_통합");
        _settings.A4FitSuffix = NormalizeSuffix(A4FitSuffixTextBox.Text, "_A4맞춤");
        _settings.A4OptimizedSuffix = NormalizeSuffix(A4OptimizedSuffixTextBox.Text, "_A4최적화");
        _settings.AutoCleanTempFiles = AutoCleanTempFilesCheckBox.IsChecked == true;
        _settings.RemoveMissingRecentFiles();
        _settings.Save();

        try
        {
            if (ContextMenuCheckBox.IsChecked == true)
            {
                WindowsIntegrationService.RegisterPdfContextMenu();
            }
            else
            {
                WindowsIntegrationService.RemovePdfContextMenu();
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, "우클릭 메뉴 설정을 저장하지 못했습니다.");
            MessageBox.Show(this, ex.Message, "우클릭 메뉴 설정", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        SettingsSaved = true;
        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private static string NormalizeSuffix(string? value, string fallback)
    {
        value = value?.Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            return fallback;
        }

        return value.StartsWith('_') ? value : $"_{value}";
    }
}
