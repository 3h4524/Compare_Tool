using System.Security.Cryptography;
using System.IO;
using CompareTool.Models;

namespace CompareTool.Services;

public sealed class FileScanner
{
    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".docx", ".docm", ".doc", ".dotx", ".dotm", ".rtf"
    };

    public async Task<IReadOnlyList<FileComparisonResult>> ScanAsync(
        string oldFolder, string newFolder, CancellationToken cancellationToken = default)
    {
        var oldFiles = FindFiles(oldFolder).GroupBy(x => Path.GetFileName(x), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.OrderBy(p => p).First(), StringComparer.OrdinalIgnoreCase);
        var newFiles = FindFiles(newFolder).GroupBy(x => Path.GetFileName(x), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.OrderBy(p => p).First(), StringComparer.OrdinalIgnoreCase);

        var names = oldFiles.Keys.Union(newFiles.Keys, StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
        var result = new List<FileComparisonResult>(names.Count);

        foreach (var name in names)
        {
            cancellationToken.ThrowIfCancellationRequested();
            oldFiles.TryGetValue(name, out var oldPath);
            newFiles.TryGetValue(name, out var newPath);
            var status = oldPath is null ? ComparisonStatus.OnlyInNew :
                         newPath is null ? ComparisonStatus.OnlyInOld :
                         await SameContentAsync(oldPath, newPath, cancellationToken) ? ComparisonStatus.Matched : ComparisonStatus.Different;

            result.Add(new FileComparisonResult
            {
                FileName = name,
                OldPath = oldPath,
                NewPath = newPath,
                OldSize = oldPath is null ? null : new FileInfo(oldPath).Length,
                NewSize = newPath is null ? null : new FileInfo(newPath).Length,
                Status = status
            });
        }
        return result;
    }

    private static IEnumerable<string> FindFiles(string folder) =>
        Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
            .Where(path => SupportedExtensions.Contains(Path.GetExtension(path)));

    private static async Task<bool> SameContentAsync(string oldPath, string newPath, CancellationToken token)
    {
        var oldInfo = new FileInfo(oldPath);
        var newInfo = new FileInfo(newPath);
        if (oldInfo.Length != newInfo.Length) return false;
        await using var oldStream = File.OpenRead(oldPath);
        await using var newStream = File.OpenRead(newPath);
        var oldHash = await SHA256.HashDataAsync(oldStream, token);
        var newHash = await SHA256.HashDataAsync(newStream, token);
        return oldHash.AsSpan().SequenceEqual(newHash);
    }
}
