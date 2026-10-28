using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace CompareTool.Services;

public interface IWordComparisonService
{
    Task<string?> CreateComparisonAsync(string oldPath, string newPath, string outputFolder, CancellationToken token = default);
}

// Provider placeholder: add Aspose.Words.dll on the target machine, then implement this provider
// with Document.Compare(...), CompareOptions and Document.Revisions. The UI does not depend on it.
public sealed class AsposeComparisonService : IWordComparisonService
{
    private readonly Assembly _assembly;
    private readonly string _licensePath;

    private AsposeComparisonService(Assembly assembly, string licensePath)
    {
        _assembly = assembly;
        _licensePath = licensePath;
    }

    public static AsposeComparisonService? TryCreate()
    {
        var licensePath = Environment.GetEnvironmentVariable("ASPOSE_WORDS_LICENSE");
        if (string.IsNullOrWhiteSpace(licensePath))
        {
            var appFolder = AppContext.BaseDirectory;
            var candidates = new[]
            {
                Path.Combine(appFolder, "Aspose.Total.NET.lic"),
                Path.Combine(appFolder, "Aspose.Words.lic"),
                Path.Combine(appFolder, "AsposeLicense.lic")
            };
            licensePath = candidates.FirstOrDefault(File.Exists);
        }
        if (string.IsNullOrWhiteSpace(licensePath) || !File.Exists(licensePath)) return null;

        try
        {
            var assembly = Assembly.Load("Aspose.Words");
            var licenseType = assembly.GetType("Aspose.Words.License", throwOnError: true)!;
            var license = Activator.CreateInstance(licenseType)!;
            licenseType.GetMethod("SetLicense", new[] { typeof(string) })!.Invoke(license, new object[] { licensePath });
            return new AsposeComparisonService(assembly, licensePath);
        }
        catch { return null; }
    }

    public Task<string?> CreateComparisonAsync(string oldPath, string newPath, string outputFolder, CancellationToken token = default)
    {
        token.ThrowIfCancellationRequested();
        var documentType = _assembly.GetType("Aspose.Words.Document", throwOnError: true)!;
        var optionsType = _assembly.GetType("Aspose.Words.Comparing.CompareOptions", throwOnError: true)!;
        var granularityType = _assembly.GetType("Aspose.Words.Comparing.Granularity", throwOnError: true)!;
        var oldDocument = Activator.CreateInstance(documentType, oldPath)!;
        var newDocument = Activator.CreateInstance(documentType, newPath)!;
        var options = Activator.CreateInstance(optionsType)!;
        SetProperty(optionsType, options, "IgnoreFormatting", false);
        SetProperty(optionsType, options, "IgnoreHeadersAndFooters", false);
        SetProperty(optionsType, options, "IgnoreTables", false);
        SetProperty(optionsType, options, "IgnoreComments", false);
        SetProperty(optionsType, options, "IgnoreTextboxes", false);
        SetProperty(optionsType, options, "IgnoreFields", false);
        SetProperty(optionsType, options, "IgnoreFootnotes", false);
        SetProperty(optionsType, options, "Granularity", Enum.Parse(granularityType, "Word"));

        var compare = documentType.GetMethod("Compare", new[] { documentType, typeof(string), typeof(DateTime), optionsType })
                     ?? throw new MissingMethodException("Không tìm thấy Aspose Document.Compare.");
        compare.Invoke(oldDocument, new[] { newDocument, "CompareTool", DateTime.Now, options });
        Directory.CreateDirectory(outputFolder);
        var resultPath = Path.Combine(outputFolder, "CompareResult.docx");
        documentType.GetMethod("Save", new[] { typeof(string) })!.Invoke(oldDocument, new object[] { resultPath });
        return Task.FromResult<string?>(resultPath);
    }

    private static void SetProperty(Type type, object instance, string name, object value) => type.GetProperty(name)?.SetValue(instance, value);
}

public static class ComparisonServiceFactory
{
    public static IWordComparisonService Create() => (IWordComparisonService?)AsposeComparisonService.TryCreate() ?? new WordInteropComparisonService();
}

public sealed class WordInteropComparisonService : IWordComparisonService
{
    public Task<string?> CreateComparisonAsync(string oldPath, string newPath, string outputFolder, CancellationToken token = default)
    {
        token.ThrowIfCancellationRequested();
        var wordType = Type.GetTypeFromProgID("Word.Application")
                       ?? throw new InvalidOperationException("Không tìm thấy Microsoft Word trên máy này.");
        dynamic? app = null;
        dynamic? oldDocument = null;
        dynamic? newDocument = null;
        try
        {
            app = Activator.CreateInstance(wordType) ?? throw new InvalidOperationException("Không khởi động được Word.");
            app.Visible = true;
            oldDocument = app.Documents.Open(oldPath, ReadOnly: true, AddToRecentFiles: false);
            newDocument = app.Documents.Open(newPath, ReadOnly: true, AddToRecentFiles: false);
            // Explicitly enable every Word comparison category. Do not rely on Word's optional defaults.
            dynamic compared = app.CompareDocuments(
                oldDocument,
                newDocument,
                Destination: 2,          // wdCompareDestinationNew
                Granularity: 1,          // wdGranularityWordLevel
                CompareFormatting: true,
                CompareCaseChanges: true,
                CompareWhitespace: true,
                CompareTables: true,
                CompareHeaders: true,
                CompareFootnotes: true,
                CompareTextboxes: true,
                CompareFields: true,
                CompareComments: true,
                CompareMoves: true,
                RevisedAuthor: "CompareTool",
                IgnoreAllComparisonWarnings: false);
            Directory.CreateDirectory(outputFolder);
            var resultPath = Path.Combine(outputFolder, "CompareResult.docx");
            compared.SaveAs2(resultPath);
            compared.Activate();
            return Task.FromResult<string?>(resultPath);
        }
        finally
        {
            if (oldDocument is not null) { try { oldDocument.Close(false); } catch { } Marshal.FinalReleaseComObject(oldDocument); }
            if (newDocument is not null) { try { newDocument.Close(false); } catch { } Marshal.FinalReleaseComObject(newDocument); }
            if (app is not null) Marshal.FinalReleaseComObject(app);
        }
    }
}

public static class ExternalFileLauncher
{
    public static void Open(string path) => Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
}
