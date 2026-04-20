using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using PdfMergeTool.Models;
using PdfMergeTool.Services;

namespace PdfMergeTool;

public partial class MergeWindow : Window
{
    private readonly PdfMergeService _mergeService = new();
    private readonly AppSettings _settings = AppSettings.Load();
    private Point _dragStartPoint;
    private string? _manualOutputPath;

    public MergeWindow(IEnumerable<string> initialFiles)
    {
        InitializeComponent();
        Files = new ObservableCollection<PdfInputFile>();
        DataContext = this;
        AddFiles(initialFiles);
        RefreshOutputPath();
    }

    public ObservableCollection<PdfInputFile> Files { get; }

    public void AddFiles(IEnumerable<string> paths)
    {
        var existing = Files.Select(file => file.Path).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var path in paths)
        {
            if (!File.Exists(path) || !string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var fullPath = Path.GetFullPath(path);
            if (existing.Add(fullPath))
            {
                Files.Add(new PdfInputFile(fullPath));
            }
        }

        RefreshOutputPath();
    }

    private void RefreshOutputPath()
    {
        if (!string.IsNullOrWhiteSpace(_manualOutputPath))
        {
            OutputPathText.Text = _manualOutputPath;
            return;
        }

        OutputPathText.Text = Files.Count == 0
            ? "PDF 파일을 추가하세요."
            : OutputPathService.CreateDefaultOutputPath(Files, _settings.MergedSuffix);
    }

    private void OnAddClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "PDF 파일 (*.pdf)|*.pdf",
            Multiselect = true,
            Title = "통합할 PDF 선택"
        };

        if (dialog.ShowDialog(this) == true)
        {
            AddFiles(dialog.FileNames);
        }
    }

    private void OnRemoveClick(object sender, RoutedEventArgs e)
    {
        var selected = FilesList.SelectedItems.Cast<PdfInputFile>().ToList();
        foreach (var item in selected)
        {
            Files.Remove(item);
        }

        _manualOutputPath = null;
        RefreshOutputPath();
    }

    private void OnMoveUpClick(object sender, RoutedEventArgs e)
    {
        var index = FilesList.SelectedIndex;
        if (index <= 0)
        {
            return;
        }

        Files.Move(index, index - 1);
        FilesList.SelectedIndex = index - 1;
        _manualOutputPath = null;
        RefreshOutputPath();
    }

    private void OnMoveDownClick(object sender, RoutedEventArgs e)
    {
        var index = FilesList.SelectedIndex;
        if (index < 0 || index >= Files.Count - 1)
        {
            return;
        }

        Files.Move(index, index + 1);
        FilesList.SelectedIndex = index + 1;
        _manualOutputPath = null;
        RefreshOutputPath();
    }

    private void OnSortAscendingClick(object sender, RoutedEventArgs e)
    {
        ReplaceFiles(Files.OrderBy(file => file.FileName, StringComparer.CurrentCultureIgnoreCase));
    }

    private void OnSortDescendingClick(object sender, RoutedEventArgs e)
    {
        ReplaceFiles(Files.OrderByDescending(file => file.FileName, StringComparer.CurrentCultureIgnoreCase));
    }

    private void ReplaceFiles(IEnumerable<PdfInputFile> orderedFiles)
    {
        var items = orderedFiles.ToList();
        Files.Clear();
        foreach (var item in items)
        {
            Files.Add(item);
        }

        _manualOutputPath = null;
        RefreshOutputPath();
    }

    private void OnChooseOutputClick(object sender, RoutedEventArgs e)
    {
        if (Files.Count == 0)
        {
            MessageBox.Show(this, "먼저 PDF 파일을 추가하세요.", "PDF 통합", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = SavePathPromptService.CreateSaveDialog(
            OutputPathService.CreateDefaultOutputPath(Files, _settings.MergedSuffix),
            "통합 PDF 저장 위치");

        if (dialog.ShowDialog(this) == true)
        {
            _manualOutputPath = dialog.FileName;
            RefreshOutputPath();
        }
    }

    private async void OnMergeClick(object sender, RoutedEventArgs e)
    {
        if (Files.Count < 2)
        {
            MessageBox.Show(this, "두 개 이상의 PDF를 추가하세요.", "PDF 통합", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var outputPath = _manualOutputPath ?? OutputPathService.CreateDefaultOutputPath(Files, _settings.MergedSuffix);
        outputPath = SavePathPromptService.ResolveOutputPath(this, outputPath, "PDF 통합") ?? string.Empty;
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            return;
        }

        _manualOutputPath = outputPath;
        RefreshOutputPath();
        MergeButton.IsEnabled = false;
        MergeButton.Content = "통합 중...";

        try
        {
            var result = await _mergeService.MergeAsync(Files.ToList(), outputPath, CancellationToken.None);
            var message = string.IsNullOrWhiteSpace(result.WarningMessage)
                ? $"저장 완료:\n{result.OutputPath}"
                : $"저장 완료:\n{result.OutputPath}\n\n참고: {result.WarningMessage}";

            MessageBox.Show(this, message, "PDF 통합", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "PDF 통합 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            MergeButton.Content = "통합";
            MergeButton.IsEnabled = true;
        }
    }

    private async void OnInterleaveClick(object sender, RoutedEventArgs e)
    {
        if (Files.Count != 2)
        {
            MessageBox.Show(this, "양면 스캔 섞기는 PDF 두 개가 필요합니다.\n첫 번째 파일은 앞면, 두 번째 파일은 뒷면 역순으로 처리합니다.", "PDF 통합", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var outputPath = _manualOutputPath ?? OutputPathService.CreateDefaultOutputPath(Files, _settings.MergedSuffix);
        outputPath = SavePathPromptService.ResolveOutputPath(this, outputPath, "PDF 섞기") ?? string.Empty;
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            return;
        }

        _manualOutputPath = outputPath;
        RefreshOutputPath();
        MergeButton.IsEnabled = false;
        MergeButton.Content = "섞는 중...";

        try
        {
            var result = await _mergeService.InterleaveAsync(Files[0], Files[1], outputPath, CancellationToken.None);
            var message = string.IsNullOrWhiteSpace(result.WarningMessage)
                ? $"저장 완료:\n{result.OutputPath}"
                : $"저장 완료:\n{result.OutputPath}\n\n참고: {result.WarningMessage}";

            MessageBox.Show(this, message, "PDF 통합", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "PDF 통합 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            MergeButton.Content = "통합";
            MergeButton.IsEnabled = true;
        }
    }

    private void OnWindowDragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.Move;
        e.Handled = true;
    }

    private void OnWindowDrop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
        {
            AddFiles(paths);
        }
    }

    private void OnFilesListMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _dragStartPoint = e.GetPosition(null);
    }

    private void OnFilesListMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || FilesList.SelectedItem is not PdfInputFile selected)
        {
            return;
        }

        var current = e.GetPosition(null);
        if (Math.Abs(current.X - _dragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(current.Y - _dragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
        {
            return;
        }

        DragDrop.DoDragDrop(FilesList, selected, DragDropEffects.Move);
    }

    private void OnFilesListDrop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
        {
            AddFiles(paths);
            return;
        }

        if (e.Data.GetData(typeof(PdfInputFile)) is not PdfInputFile source)
        {
            return;
        }

        var target = GetTargetItem(e);
        if (target is null || ReferenceEquals(source, target))
        {
            return;
        }

        var oldIndex = Files.IndexOf(source);
        var newIndex = Files.IndexOf(target);
        if (oldIndex >= 0 && newIndex >= 0)
        {
            Files.Move(oldIndex, newIndex);
            FilesList.SelectedItem = source;
            _manualOutputPath = null;
            RefreshOutputPath();
        }
    }

    private PdfInputFile? GetTargetItem(DragEventArgs e)
    {
        var element = FilesList.InputHitTest(e.GetPosition(FilesList)) as DependencyObject;
        while (element is not null)
        {
            if (element is FrameworkElement { DataContext: PdfInputFile file })
            {
                return file;
            }

            element = System.Windows.Media.VisualTreeHelper.GetParent(element);
        }

        return null;
    }
}
