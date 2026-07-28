using System.Windows;

namespace PdfMergeTool;

public partial class SplitIntervalWindow : Window
{
    public SplitIntervalWindow()
    {
        InitializeComponent();
        Loaded += (_, _) =>
        {
            IntervalTextBox.Focus();
            IntervalTextBox.SelectAll();
        };
    }

    public int Interval { get; private set; } = 1;

    private void OnSplitClick(object sender, RoutedEventArgs e)
    {
        if (!int.TryParse(IntervalTextBox.Text.Trim(), out var interval) || interval < 1)
        {
            MessageBox.Show(this, "1 이상의 정수를 입력하세요.", "N페이지마다 분리", MessageBoxButton.OK, MessageBoxImage.Warning);
            IntervalTextBox.Focus();
            IntervalTextBox.SelectAll();
            return;
        }

        Interval = interval;
        DialogResult = true;
    }
}
