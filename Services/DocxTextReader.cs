using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace CompareTool.Services;

public enum DocxBlockKind { Paragraph, Table }

public sealed class DocxRun
{
    public string Text { get; init; } = "";
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public bool Strike { get; init; }
    public double? FontSizePoints { get; init; }
    public string? FontFamily { get; init; }
    public string? Foreground { get; init; }
    public string? Highlight { get; init; }
}

public sealed class DocxBlock
{
    public required DocxBlockKind Kind { get; init; }
    public string Text { get; init; } = "";
    public List<DocxRun> Runs { get; init; } = [];
    public List<List<string>> Rows { get; init; } = [];
    public string FormattingKey { get; init; } = "";
    public double BeforeSpacingPoints { get; init; }
    public double AfterSpacingPoints { get; init; }
    public double LeftIndentPoints { get; init; }
    public double RightIndentPoints { get; init; }
    public double FirstLineIndentPoints { get; init; }
    public string Key => Kind == DocxBlockKind.Paragraph ? $"P|{Normalize(Text)}|F|{FormattingKey}" : $"T|{string.Join("|", Rows.Select(row => string.Join("\u001f", row.Select(Normalize))))}";
    private static string Normalize(string value) => value.Replace("\r\n", "\n").Replace("\r", "\n");
}

public static class DocxTextReader
{
    public static List<DocxBlock> ReadBlocks(string path)
    {
        if (!path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) && !path.EndsWith(".docm", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("Preview trong app hiện hỗ trợ .docx và .docm. Hãy dùng Open in Word Compare cho định dạng khác.");
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry("word/document.xml") ?? throw new InvalidDataException("Không tìm thấy word/document.xml trong file Word.");
        using var stream = entry.Open();
        var xml = XDocument.Load(stream);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var body = xml.Root?.Element(w + "body") ?? throw new InvalidDataException("Document không có body.");
        var blocks = new List<DocxBlock>();
        foreach (var element in body.Elements())
        {
            if (element.Name == w + "p")
            {
                blocks.Add(new DocxBlock
                {
                    Kind = DocxBlockKind.Paragraph,
                    Text = ReadElementText(element, w),
                    Runs = ReadRuns(element, w),
                    FormattingKey = ReadFormattingKey(element, w),
                    BeforeSpacingPoints = ReadSpacingPoints(element, w, "before"),
                    AfterSpacingPoints = ReadSpacingPoints(element, w, "after"),
                    LeftIndentPoints = ReadIndentPoints(element, w, "left", "start"),
                    RightIndentPoints = ReadIndentPoints(element, w, "right", "end"),
                    FirstLineIndentPoints = ReadIndentPoints(element, w, "firstLine", null) - ReadIndentPoints(element, w, "hanging", null)
                });
            }
            else if (element.Name == w + "tbl")
            {
                var rows = element.Elements(w + "tr").Select(row => row.Elements(w + "tc").Select(cell => ReadElementText(cell, w)).ToList()).ToList();
                blocks.Add(new DocxBlock { Kind = DocxBlockKind.Table, Rows = rows });
            }
        }
        return blocks;
    }

    public static string Read(string path)
    {
        var builder = new StringBuilder();
        foreach (var block in ReadBlocks(path))
        {
            if (block.Kind == DocxBlockKind.Paragraph) builder.AppendLine(block.Text);
            else foreach (var row in block.Rows) builder.AppendLine(string.Join(" | ", row));
        }
        return builder.ToString();
    }

    private static string ReadElementText(XElement element, XNamespace w)
    {
        var builder = new StringBuilder();
        foreach (var node in element.Descendants())
        {
            if (node.Name == w + "t") builder.Append(node.Value);
            else if (node.Name == w + "tab") builder.Append('\t');
            else if (node.Name == w + "br") builder.AppendLine();
        }
        return builder.ToString();
    }

    private static List<DocxRun> ReadRuns(XElement paragraph, XNamespace w)
    {
        var runs = new List<DocxRun>();
        foreach (var run in paragraph.Descendants(w + "r"))
        {
            var text = new StringBuilder();
            foreach (var node in run.Descendants())
            {
                if (node.Name == w + "t") text.Append(node.Value);
                else if (node.Name == w + "tab") text.Append('\t');
                else if (node.Name == w + "br") text.AppendLine();
            }
            if (text.Length == 0) continue;
            var properties = run.Element(w + "rPr");
            var size = properties?.Element(w + "sz")?.Attribute(w + "val")?.Value;
            var font = properties?.Element(w + "rFonts")?.Attribute(w + "ascii")?.Value
                       ?? properties?.Element(w + "rFonts")?.Attribute(w + "hAnsi")?.Value;
            runs.Add(new DocxRun
            {
                Text = text.ToString(),
                Bold = IsOn(properties?.Element(w + "b")),
                Italic = IsOn(properties?.Element(w + "i")),
                Underline = properties?.Element(w + "u")?.Attribute(w + "val")?.Value is not "none",
                Strike = IsOn(properties?.Element(w + "strike")),
                FontSizePoints = int.TryParse(size, out var halfPoints) ? halfPoints / 2.0 : null,
                FontFamily = font,
                Foreground = properties?.Element(w + "color")?.Attribute(w + "val")?.Value,
                Highlight = properties?.Element(w + "highlight")?.Attribute(w + "val")?.Value
            });
        }
        return runs;
    }

    private static bool IsOn(XElement? element)
    {
        if (element is null) return false;
        var value = element.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "val")?.Value;
        return value is null || (value != "0" && value != "false" && value != "off");
    }

    private static string ReadFormattingKey(XElement paragraph, XNamespace w)
    {
        var paragraphProperties = CanonicalizeFormatting(paragraph.Element(w + "pPr"));
        var runProperties = string.Join("|", paragraph.Descendants(w + "rPr").Select(CanonicalizeFormatting));
        return paragraphProperties + "|R|" + runProperties;
    }

    private static string CanonicalizeFormatting(XElement? element)
    {
        if (element is null) return "";
        var copy = new XElement(element);
        foreach (var node in copy.DescendantsAndSelf().Where(node => node.Name.LocalName is "pStyle" or "numPr").ToList()) node.Remove();
        foreach (var attribute in copy.DescendantsAndSelf().Attributes().Where(attribute => attribute.Name.LocalName.StartsWith("rsid", StringComparison.OrdinalIgnoreCase)).ToList()) attribute.Remove();
        foreach (var node in copy.DescendantsAndSelf())
        {
            foreach (var attribute in node.Attributes().Where(attribute => attribute.Name.LocalName == "space" && attribute.Value == "0").ToList()) attribute.Remove();
            node.ReplaceAttributes(node.Attributes().OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal).ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal).ThenBy(attribute => attribute.Value, StringComparer.Ordinal));
        }
        return copy.ToString(SaveOptions.DisableFormatting);
    }

    private static double ReadSpacingPoints(XElement paragraph, XNamespace w, string attributeName)
    {
        var spacing = paragraph.Element(w + "pPr")?.Element(w + "spacing");
        var attribute = spacing?.Attribute(XName.Get(attributeName, w.NamespaceName))?.Value;
        return int.TryParse(attribute, out var twips) ? twips / 20.0 : 0;
    }

    private static double ReadIndentPoints(XElement paragraph, XNamespace w, string primaryAttribute, string? fallbackAttribute)
    {
        var indentation = paragraph.Element(w + "pPr")?.Element(w + "ind");
        var value = indentation?.Attribute(XName.Get(primaryAttribute, w.NamespaceName))?.Value;
        if (value is null && fallbackAttribute is not null)
            value = indentation?.Attribute(XName.Get(fallbackAttribute, w.NamespaceName))?.Value;
        return int.TryParse(value, out var twips) ? twips / 20.0 : 0;
    }
}
