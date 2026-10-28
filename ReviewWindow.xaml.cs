using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using CompareTool.Services;
using WpfColor = System.Windows.Media.Color;
using WpfMessageBox = System.Windows.MessageBox;
using WpfRichTextBox = System.Windows.Controls.RichTextBox;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfFontFamily = System.Windows.Media.FontFamily;
using WpfColorConverter = System.Windows.Media.ColorConverter;

namespace CompareTool;

public partial class ReviewWindow : Window
{
    private readonly List<SummaryItem> _summaryItems = [];
    private string _oldPath = "";
    private string _newPath = "";
    private readonly string? _comparisonFolder;

    public ReviewWindow(string fileName, string oldPath, string newPath, string? comparisonFolder = null)
    {
        InitializeComponent();
        _oldPath = oldPath;
        _newPath = newPath;
        _comparisonFolder = comparisonFolder;
        HeaderText.Text = $"{fileName} — preview có cấu trúc Paragraph/Table";
        try
        {
            RenderDiff(DocxTextReader.ReadBlocks(oldPath), DocxTextReader.ReadBlocks(newPath));
        }
        catch (Exception ex)
        {
            WpfMessageBox.Show(this, ex.Message, "Không đọc được document", MessageBoxButton.OK, MessageBoxImage.Error);
            Close();
        }
    }

    private void RenderDiff(List<DocxBlock> oldBlocks, List<DocxBlock> newBlocks)
    {
        OldText.Document.Blocks.Clear();
        NewText.Document.Blocks.Clear();
        SummaryList.Items.Clear();
        _summaryItems.Clear();

        foreach (var pair in AlignBlocks(oldBlocks, newBlocks))
        {
            var changed = pair.OldBlock?.Key != pair.NewBlock?.Key;
            var oldAnchor = RenderBlock(OldText, pair.OldBlock, changed, WpfColor.FromRgb(255, 224, 224));
            var newAnchor = RenderBlock(NewText, pair.NewBlock, changed, WpfColor.FromRgb(223, 246, 221));
            if (!changed) continue;

            var item = new SummaryItem
            {
                Label = BuildSummaryLabel(_summaryItems.Count + 1, pair.OldBlock, pair.NewBlock),
                OldNavigate = CreateNavigator(OldText, oldAnchor),
                NewNavigate = CreateNavigator(NewText, newAnchor)
            };
            _summaryItems.Add(item);
            SummaryList.Items.Add(item);
        }

        if (_summaryItems.Count == 0)
            SummaryList.Items.Add(new SummaryItem { Label = "Không phát hiện khác biệt" });
    }

    private static string BuildSummaryLabel(int number, DocxBlock? oldBlock, DocxBlock? newBlock)
    {
        var block = oldBlock ?? newBlock;
        var kind = block?.Kind == DocxBlockKind.Table ? "TABLE" : "TEXT";
        var oldText = DescribeBlock(oldBlock);
        var newText = DescribeBlock(newBlock);
        var detail = oldBlock is not null && newBlock is not null &&
                     oldBlock.Kind == DocxBlockKind.Paragraph && newBlock.Kind == DocxBlockKind.Paragraph &&
                     oldBlock.Text == newBlock.Text && oldBlock.Key != newBlock.Key
            ? $"Spacing/formatting thay đổi | OLD: {oldText} | NEW: {newText}"
            : oldBlock is null ? $"NEW: {newText}" :
              newBlock is null ? $"OLD: {oldText}" : $"OLD: {oldText}  →  NEW: {newText}";
        return $"{number}. [{kind}] {Shorten(detail)}";
    }

    private static string Shorten(string value) => value.Length <= 140 ? value : value[..140] + "…";

    private static string DescribeBlock(DocxBlock? block)
    {
        if (block is null) return "(không có)";
        if (block.Kind == DocxBlockKind.Table)
        {
            var columns = block.Rows.DefaultIfEmpty([]).Max(row => row.Count);
            return $"table {block.Rows.Count} dòng × {columns} cột";
        }
        var text = block.Text.Replace(" ", "·").Replace("\t", "⇥").Replace("\r", "").Replace("\n", "↵");
        return string.IsNullOrWhiteSpace(text) ? "(paragraph trống)" : text.Trim();
    }

    private static IReadOnlyList<(DocxBlock? OldBlock, DocxBlock? NewBlock)> AlignBlocks(List<DocxBlock> oldBlocks, List<DocxBlock> newBlocks)
    {
        var lcs = new int[oldBlocks.Count + 1, newBlocks.Count + 1];
        for (var i = oldBlocks.Count - 1; i >= 0; i--)
        for (var j = newBlocks.Count - 1; j >= 0; j--)
            lcs[i, j] = oldBlocks[i].Key == newBlocks[j].Key ? lcs[i + 1, j + 1] + 1 : Math.Max(lcs[i + 1, j], lcs[i, j + 1]);

        var result = new List<(DocxBlock?, DocxBlock?)>();
        var oldIndex = 0;
        var newIndex = 0;
        while (oldIndex < oldBlocks.Count && newIndex < newBlocks.Count)
        {
            if (oldBlocks[oldIndex].Key == newBlocks[newIndex].Key ||
                (oldBlocks[oldIndex].Kind == newBlocks[newIndex].Kind && oldBlocks[oldIndex].Text == newBlocks[newIndex].Text))
                result.Add((oldBlocks[oldIndex++], newBlocks[newIndex++]));
            else if (lcs[oldIndex + 1, newIndex] >= lcs[oldIndex, newIndex + 1])
                result.Add((oldBlocks[oldIndex++], null));
            else
                result.Add((null, newBlocks[newIndex++]));
        }
        while (oldIndex < oldBlocks.Count) result.Add((oldBlocks[oldIndex++], null));
        while (newIndex < newBlocks.Count) result.Add((null, newBlocks[newIndex++]));
        return result;
    }

    private static TextPointer? RenderBlock(WpfRichTextBox box, DocxBlock? block, bool changed, WpfColor color)
    {
        if (block is null) return null;
        if (block.Kind == DocxBlockKind.Paragraph)
        {
            var paragraph = new Paragraph
            {
                Margin = new Thickness(block.LeftIndentPoints, block.BeforeSpacingPoints, block.RightIndentPoints, block.AfterSpacingPoints + 8),
                TextIndent = block.FirstLineIndentPoints
            };
            if (block.Runs.Count == 0)
            {
                var run = new Run(block.Text + Environment.NewLine);
                if (changed) run.Background = new SolidColorBrush(color);
                paragraph.Inlines.Add(run);
            }
            else
            {
                foreach (var docxRun in block.Runs)
                {
                    var run = new Run(docxRun.Text)
                    {
                        FontWeight = docxRun.Bold ? FontWeights.Bold : FontWeights.Normal,
                        FontStyle = docxRun.Italic ? FontStyles.Italic : FontStyles.Normal,
                        TextDecorations = BuildDecorations(docxRun),
                        FontSize = docxRun.FontSizePoints ?? 12,
                        FontFamily = string.IsNullOrWhiteSpace(docxRun.FontFamily) ? new WpfFontFamily("Segoe UI") : new WpfFontFamily(docxRun.FontFamily)
                    };
                    if (!string.IsNullOrWhiteSpace(docxRun.Foreground) && docxRun.Foreground != "auto")
                        run.Foreground = new SolidColorBrush(ParseColor(docxRun.Foreground, Colors.Black));
                    if (!string.IsNullOrWhiteSpace(docxRun.Highlight) && docxRun.Highlight != "none")
                        run.Background = new SolidColorBrush(ParseHighlight(docxRun.Highlight));
                    if (changed) run.Background = new SolidColorBrush(color);
                    paragraph.Inlines.Add(run);
                }
                paragraph.Inlines.Add(new Run(Environment.NewLine));
            }
            box.Document.Blocks.Add(paragraph);
            return paragraph.ContentStart;
        }

        var table = new Table { CellSpacing = 0, Margin = new Thickness(0, 4, 0, 12) };
        var columns = block.Rows.DefaultIfEmpty([]).Max(row => row.Count);
        for (var column = 0; column < columns; column++) table.Columns.Add(new TableColumn());
        var group = new TableRowGroup();
        foreach (var row in block.Rows)
        {
            var tableRow = new TableRow();
            for (var column = 0; column < columns; column++)
            {
                var cell = new TableCell(new Paragraph(new Run(column < row.Count ? row[column] : "")))
                {
                    BorderBrush = WpfBrushes.Gray,
                    BorderThickness = new Thickness(0.5),
                    Padding = new Thickness(5),
                    Background = changed ? new SolidColorBrush(color) : WpfBrushes.Transparent
                };
                tableRow.Cells.Add(cell);
            }
            group.Rows.Add(tableRow);
        }
        table.RowGroups.Add(group);
        box.Document.Blocks.Add(table);
        return table.ContentStart;
    }

    private void SummaryList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (SummaryList.SelectedItem is not SummaryItem item) return;
        item.OldNavigate?.Invoke();
        item.NewNavigate?.Invoke();
    }

    private async void OpenInWord_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var service = Services.ComparisonServiceFactory.Create();
            var outputFolder = _comparisonFolder ?? Path.Combine(Path.GetTempPath(), "CompareTool");
            var output = await service.CreateComparisonAsync(_oldPath, _newPath, outputFolder);
            if (output is not null && service is Services.AsposeComparisonService)
                Services.ExternalFileLauncher.Open(output);
            if (output is not null) HeaderText.Text = $"Đã tạo Word Compare: {output}";
        }
        catch (Exception ex) { WpfMessageBox.Show(this, ex.Message, "Word Compare", MessageBoxButton.OK, MessageBoxImage.Error); }
    }

    private sealed class SummaryItem
    {
        public string Label { get; init; } = "";
        public Action? OldNavigate { get; init; }
        public Action? NewNavigate { get; init; }
    }

    private static Action? CreateNavigator(WpfRichTextBox box, TextPointer? pointer) => pointer is null
        ? null
        : () => NavigateTo(box, pointer);

    private static void NavigateTo(WpfRichTextBox box, TextPointer pointer)
    {
        box.Focus();
        box.Selection.Select(pointer, pointer);
        box.CaretPosition = pointer;
        box.UpdateLayout();
        var rect = pointer.GetCharacterRect(LogicalDirection.Forward);
        var scrollViewer = FindVisualChild<ScrollViewer>(box);
        if (scrollViewer is not null)
            scrollViewer.ScrollToVerticalOffset(Math.Max(0, scrollViewer.VerticalOffset + rect.Top - 50));
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (var index = 0; index < VisualTreeHelper.GetChildrenCount(parent); index++)
        {
            var child = VisualTreeHelper.GetChild(parent, index);
            if (child is T match) return match;
            var nested = FindVisualChild<T>(child);
            if (nested is not null) return nested;
        }
        return null;
    }

    private static TextDecorationCollection? BuildDecorations(DocxRun run)
    {
        var decorations = new TextDecorationCollection();
        if (run.Underline) decorations.Add(TextDecorations.Underline[0]);
        if (run.Strike) decorations.Add(TextDecorations.Strikethrough[0]);
        return decorations.Count == 0 ? null : decorations;
    }

    private static WpfColor ParseColor(string value, WpfColor fallback)
    {
        try { return (WpfColor)WpfColorConverter.ConvertFromString("#" + value)!; }
        catch { return fallback; }
    }

    private static WpfColor ParseHighlight(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "yellow" => Colors.Yellow,
            "green" => Colors.LightGreen,
            "cyan" => Colors.Cyan,
            "magenta" => Colors.Magenta,
            "red" => Colors.LightCoral,
            "blue" => Colors.LightBlue,
            "darkblue" => Colors.DarkBlue,
            "gray" or "darkgray" => Colors.LightGray,
            _ => ParseColor(value, Colors.Yellow)
        };
    }
}
