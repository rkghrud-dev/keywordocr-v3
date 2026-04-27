using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace KeywordOcr.App.Services;

public sealed class MarketPlusCategoryMapAutoUploader
{
    private const string HelperBaseUrl = "http://127.0.0.1:5555";
    private static readonly Regex GsCodeRegex = new(@"(GS\d{7}[A-Z0-9]*)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(8) };

    private readonly string _v3Root;
    private readonly IProgress<string>? _progress;

    public MarketPlusCategoryMapAutoUploader(string v3Root, IProgress<string>? progress = null)
    {
        _v3Root = v3Root;
        _progress = progress;
    }

    public async Task<string?> UploadLatestAsync(
        IEnumerable<string?> candidateRoots,
        IEnumerable<string?> priorityFiles,
        CancellationToken cancellationToken)
    {
        var mapPath = FindLatestCategoryMap(candidateRoots, priorityFiles);
        if (mapPath is null)
        {
            _progress?.Report("[카테고리맵] 자동 업로드 스킵: category_match 파일 없음");
            return null;
        }

        await EnsureHelperServerAsync(cancellationToken);

        var aliases = CollectProductAliases(priorityFiles);
        var payload = JsonSerializer.Serialize(new
        {
            filename = Path.GetFileName(mapPath),
            contentBase64 = Convert.ToBase64String(await File.ReadAllBytesAsync(mapPath, cancellationToken)),
            aliases
        });

        using var content = new StringContent(payload, Encoding.UTF8, "application/json");
        using var response = await Http.PostAsync($"{HelperBaseUrl}/api/upload-map", content, cancellationToken);
        var body = await response.Content.ReadAsStringAsync(cancellationToken);

        if (!response.IsSuccessStatusCode)
            throw new InvalidOperationException($"helper upload failed: {(int)response.StatusCode} {body}");

        using var doc = JsonDocument.Parse(body);
        var root = doc.RootElement;
        var ok = root.TryGetProperty("ok", out var okElement) && okElement.ValueKind == JsonValueKind.True;
        if (!ok)
        {
            var error = root.TryGetProperty("error", out var errorElement) ? errorElement.GetString() : body;
            throw new InvalidOperationException(error ?? "helper upload failed");
        }

        var recordCount = root.TryGetProperty("recordCount", out var recordElement) ? recordElement.GetInt32() : 0;
        var productCount = root.TryGetProperty("productCount", out var productElement) ? productElement.GetInt32() : 0;
        _progress?.Report($"[카테고리맵] 자동 업로드 완료: {Path.GetFileName(mapPath)} (상품 {productCount}, 매칭 {recordCount}, 별칭 {aliases.Count})");
        return mapPath;
    }

    private async Task EnsureHelperServerAsync(CancellationToken cancellationToken)
    {
        if (await IsHelperReadyAsync(cancellationToken))
            return;

        StartHelperServer();

        for (var i = 0; i < 20; i++)
        {
            await Task.Delay(300, cancellationToken);
            if (await IsHelperReadyAsync(cancellationToken))
                return;
        }

        throw new InvalidOperationException("MarketPlus helper server did not become ready on localhost:5555");
    }

    private static async Task<bool> IsHelperReadyAsync(CancellationToken cancellationToken)
    {
        try
        {
            using var response = await Http.GetAsync($"{HelperBaseUrl}/api/map/status", cancellationToken);
            if (!response.IsSuccessStatusCode)
                return false;

            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.TryGetProperty("supportsAliases", out var aliasesElement)
                && aliasesElement.ValueKind == JsonValueKind.True;
        }
        catch
        {
            return false;
        }
    }

    private void StartHelperServer()
    {
        var launcherPath = Path.Combine(_v3Root, "tools", "marketplus-category-helper", "marketplus_category_launcher.ps1");
        if (!File.Exists(launcherPath))
            throw new FileNotFoundException("MarketPlus helper launcher not found.", launcherPath);

        _progress?.Report("[카테고리맵] MarketPlus helper 서버 시작 중...");

        var startInfo = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{launcherPath}\" -Force",
            WorkingDirectory = Path.GetDirectoryName(launcherPath) ?? _v3Root,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        Process.Start(startInfo);
    }

    private static string? FindLatestCategoryMap(IEnumerable<string?> candidateRoots, IEnumerable<string?> priorityFiles)
    {
        var candidates = new List<string>();

        foreach (var file in priorityFiles.Where(p => !string.IsNullOrWhiteSpace(p)))
        {
            if (File.Exists(file) &&
                Path.GetFileName(file).Contains("category_match", StringComparison.OrdinalIgnoreCase) &&
                !Path.GetFileName(file).StartsWith("~$", StringComparison.Ordinal))
            {
                candidates.Add(file);
            }
        }

        foreach (var root in candidateRoots.Where(p => !string.IsNullOrWhiteSpace(p)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (!Directory.Exists(root))
                continue;

            try
            {
                candidates.AddRange(Directory.EnumerateFiles(root, "*category_match*.xlsx", SearchOption.TopDirectoryOnly)
                    .Where(IsUsableWorkbook));
            }
            catch
            {
                // Ignore roots that disappear while the workflow is running.
            }
        }

        foreach (var root in candidateRoots.Where(p => !string.IsNullOrWhiteSpace(p)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (!Directory.Exists(root))
                continue;

            try
            {
                candidates.AddRange(Directory.EnumerateFiles(root, "*category_match*.xlsx", SearchOption.AllDirectories)
                    .Where(IsUsableWorkbook));
            }
            catch
            {
                // Some user folders may be locked by Excel or cloud sync; skip and keep searching.
            }
        }

        return candidates
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(File.GetLastWriteTime)
            .FirstOrDefault();
    }

    private static bool IsUsableWorkbook(string path)
        => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal);

    private static List<CategoryMapAlias> CollectProductAliases(IEnumerable<string?> files)
    {
        var aliases = new List<CategoryMapAlias>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var file in files.Where(p => !string.IsNullOrWhiteSpace(p)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (!File.Exists(file) || !IsUsableWorkbook(file) || !string.Equals(Path.GetExtension(file), ".xlsx", StringComparison.OrdinalIgnoreCase))
                continue;

            try
            {
                using var workbook = new XLWorkbook(file);
                foreach (var worksheet in workbook.Worksheets)
                    CollectAliasesFromWorksheet(worksheet, aliases, seen);
            }
            catch
            {
                // Alias extraction is best-effort. The category map upload itself should still run.
            }
        }

        return aliases;
    }

    private static void CollectAliasesFromWorksheet(IXLWorksheet worksheet, List<CategoryMapAlias> aliases, HashSet<string> seen)
    {
        var range = worksheet.RangeUsed();
        if (range is null)
            return;

        var headerRow = range.FirstRowUsed();
        if (headerRow is null)
            return;

        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var cell in headerRow.CellsUsed())
        {
            var key = NormalizeHeader(cell.GetString());
            if (!string.IsNullOrWhiteSpace(key) && !headers.ContainsKey(key))
                headers[key] = cell.Address.ColumnNumber;
        }

        var nameColumn = FindColumn(headers, "상품명", "제품명", "productname", "product_name", "name", "최종키워드2차");
        if (nameColumn <= 0)
            return;

        var codeColumn = FindColumn(headers, "자체상품코드", "자체 상품코드", "상품코드", "gs코드", "gscode", "gs_code", "productcode", "product_code");

        foreach (var row in worksheet.RowsUsed().Where(r => r.RowNumber() > headerRow.RowNumber()))
        {
            var productName = row.Cell(nameColumn).GetString().Trim();
            if (string.IsNullOrWhiteSpace(productName))
                continue;

            var codeSource = codeColumn > 0 ? row.Cell(codeColumn).GetString() : "";
            var gsCode = ExtractGsCode(codeSource);
            if (string.IsNullOrWhiteSpace(gsCode))
                gsCode = ExtractGsCode(productName);

            if (string.IsNullOrWhiteSpace(gsCode))
                continue;

            var seenKey = $"{gsCode}|{productName}";
            if (!seen.Add(seenKey))
                continue;

            aliases.Add(new CategoryMapAlias(gsCode, gsCode, productName));
        }
    }

    private static int FindColumn(IReadOnlyDictionary<string, int> headers, params string[] candidates)
    {
        foreach (var candidate in candidates)
            if (headers.TryGetValue(NormalizeHeader(candidate), out var column))
                return column;
        return -1;
    }

    private static string NormalizeHeader(string value)
        => Regex.Replace(value ?? "", @"\s+", "").Trim().ToLowerInvariant();

    private static string ExtractGsCode(string value)
    {
        var match = GsCodeRegex.Match(value ?? "");
        return match.Success ? match.Value.ToUpperInvariant() : "";
    }

    private sealed record CategoryMapAlias(string ProductKey, string GsCode, string ProductName);
}
