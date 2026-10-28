using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using CompareTool.Models;
using CompareTool.Services;
using Forms = System.Windows.Forms;
using WpfTextBox = System.Windows.Controls.TextBox;
using WpfMessageBox = System.Windows.MessageBox;

namespace CompareTool;

public partial class MainWindow : Window
{
    private readonly FileScanner _scanner = new();
    private readonly ObservableCollection<FileComparisonResult> _visibleResults = [];
    private List<FileComparisonResult> _allResults = [];
    private string? _currentBatchFolder;

    public MainWindow()
    {
        InitializeComponent();
        ResultsList.ItemsSource = _visibleResults;
    }

    private void BrowseOld_Click(object sender, RoutedEventArgs e) => ChooseFolder(OldFolderBox);
    private void BrowseNew_Click(object sender, RoutedEventArgs e) => ChooseFolder(NewFolderBox);
    private void BrowseOutput_Click(object sender, RoutedEventArgs e) => ChooseFolder(OutputFolderBox);

    private static void ChooseFolder(WpfTextBox target)
    {
        using var dialog = new Forms.FolderBrowserDialog { Description = "Chọn folder" };
        if (dialog.ShowDialog() == Forms.DialogResult.OK) target.Text = dialog.SelectedPath;
    }

    private async void Compare_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateFolders()) return;
        try
        {
            SetBusy(true, "Đang scan và tính hash...");
            _allResults = (await _scanner.ScanAsync(OldFolderBox.Text.Trim(), NewFolderBox.Text.Trim())).ToList();
            ApplyFilter();
            StatusText.Text = $"Đã scan {_allResults.Count:N0} file.";
        }
        catch (Exception ex) { WpfMessageBox.Show(this, ex.Message, "Compare lỗi", MessageBoxButton.OK, MessageBoxImage.Error); }
        finally { SetBusy(false); }
    }

    private void ResultsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ResultsList.SelectedItem is not FileComparisonResult item || item.Status != ComparisonStatus.Different) return;
        if (item.OldPath is null || item.NewPath is null) return;
        var pairRoot = _currentBatchFolder ?? OutputFolderBox.Text.Trim();
        var pairFolder = string.IsNullOrWhiteSpace(pairRoot)
            ? null
            : Path.Combine(pairRoot, "Different", GetPairFolderName(item.FileName));
        var review = new ReviewWindow(item.FileName, item.OldPath, item.NewPath, pairFolder) { Owner = this };
        review.Show();
    }

    private void FilterBox_SelectionChanged(object sender, SelectionChangedEventArgs e) => ApplyFilter();

    private void ApplyFilter()
    {
        // SelectionChanged can fire while InitializeComponent is still building the XAML tree.
        if (_visibleResults is null || SummaryText is null || FilterBox is null) return;
        var tag = (FilterBox.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "All";
        var filtered = tag == "All" ? _allResults : _allResults.Where(x => x.Status.ToString() == tag);
        _visibleResults.Clear();
        foreach (var item in filtered) _visibleResults.Add(item);
        SummaryText.Text = $"{_allResults.Count:N0} file / {_allResults.Count(x => x.Status == ComparisonStatus.Different):N0} file khác";
    }

    private void Move_Click(object sender, RoutedEventArgs e)
    {
        if (_allResults.Count == 0) { WpfMessageBox.Show(this, "Hãy Compare trước."); return; }
        if (!Directory.Exists(OutputFolderBox.Text)) Directory.CreateDirectory(OutputFolderBox.Text);
        var answer = WpfMessageBox.Show(this, "Move các file theo nhóm vào Output? File đích đã tồn tại sẽ được bỏ qua.", "Xác nhận", MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (answer != MessageBoxResult.Yes) return;
        _currentBatchFolder = CreateNextBatchFolder(OutputFolderBox.Text.Trim());
        var moved = 0;
        foreach (var item in _allResults)
        {
            if (item.Status == ComparisonStatus.Matched && item.OldPath is not null) moved += MoveOne(item.OldPath, Path.Combine(_currentBatchFolder, "Matched"));
            if (item.Status == ComparisonStatus.Different)
            {
                var pairFolder = Path.Combine(_currentBatchFolder, "Different", GetPairFolderName(item.FileName));
                if (item.OldPath is not null) moved += MovePairFile(item.OldPath, pairFolder, "Old");
                if (item.NewPath is not null) moved += MovePairFile(item.NewPath, pairFolder, "New");
            }
            if (item.Status == ComparisonStatus.OnlyInOld && item.OldPath is not null) moved += MoveOne(item.OldPath, Path.Combine(_currentBatchFolder, "OnlyInOld"));
            if (item.Status == ComparisonStatus.OnlyInNew && item.NewPath is not null) moved += MoveOne(item.NewPath, Path.Combine(_currentBatchFolder, "OnlyInNew"));
        }
        StatusText.Text = $"Đã move {moved:N0} file vào {Path.GetFileName(_currentBatchFolder)}. File trùng đích được bỏ qua.";
    }

    private static int MoveOne(string source, string folder)
    {
        Directory.CreateDirectory(folder);
        var destination = Path.Combine(folder, Path.GetFileName(source));
        if (File.Exists(destination)) return 0;
        File.Move(source, destination);
        return 1;
    }

    private static int MovePairFile(string source, string pairFolder, string versionName)
    {
        Directory.CreateDirectory(pairFolder);
        var destination = Path.Combine(pairFolder, versionName + Path.GetExtension(source));
        if (File.Exists(destination)) return 0;
        File.Move(source, destination);
        return 1;
    }

    private static string GetPairFolderName(string fileName)
    {
        var baseName = Path.GetFileNameWithoutExtension(fileName);
        foreach (var invalid in Path.GetInvalidFileNameChars()) baseName = baseName.Replace(invalid, '_');
        return string.IsNullOrWhiteSpace(baseName) ? "Unnamed" : baseName;
    }

    private static string CreateNextBatchFolder(string outputFolder)
    {
        var number = 1;
        string folder;
        do folder = Path.Combine(outputFolder, $"Batch{number++}");
        while (Directory.Exists(folder));
        Directory.CreateDirectory(folder);
        return folder;
    }

    private bool ValidateFolders()
    {
        if (!Directory.Exists(OldFolderBox.Text) || !Directory.Exists(NewFolderBox.Text))
        { WpfMessageBox.Show(this, "Old folder và New folder phải tồn tại."); return false; }
        if (string.IsNullOrWhiteSpace(OutputFolderBox.Text))
        { WpfMessageBox.Show(this, "Hãy chọn Output folder."); return false; }
        return true;
    }

    private void SetBusy(bool busy, string? message = null)
    {
        System.Windows.Input.Mouse.OverrideCursor = busy ? System.Windows.Input.Cursors.Wait : null;
        StatusText.Text = message ?? StatusText.Text;
    }
}
