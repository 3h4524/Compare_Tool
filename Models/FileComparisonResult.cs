namespace CompareTool.Models;

public enum ComparisonStatus
{
    Matched,
    Different,
    OnlyInOld,
    OnlyInNew,
    Unsupported,
    Error
}

public sealed class FileComparisonResult
{
    public required string FileName { get; init; }
    public string? OldPath { get; init; }
    public string? NewPath { get; init; }
    public ComparisonStatus Status { get; set; }
    public long? OldSize { get; init; }
    public long? NewSize { get; init; }
    public string StatusText => Status switch
    {
        ComparisonStatus.Matched => "Trùng",
        ComparisonStatus.Different => "Khác",
        ComparisonStatus.OnlyInOld => "Chỉ có Old",
        ComparisonStatus.OnlyInNew => "Chỉ có New",
        ComparisonStatus.Unsupported => "Không hỗ trợ",
        _ => "Lỗi"
    };
    public string SizeText => $"{FormatSize(OldSize)} / {FormatSize(NewSize)}";

    private static string FormatSize(long? value) => value is null ? "-" : $"{value.Value:N0} B";
}
