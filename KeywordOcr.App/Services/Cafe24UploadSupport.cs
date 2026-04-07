using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace KeywordOcr.App.Services;

internal static class Cafe24UploadSupport
{
    private static readonly Regex GsFolderRegex = new(@"^GS\d{7}[A-Z0-9]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex GsCodeRegex = new(@"(GS\d{7}[A-Z0-9]*)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly string[] ImageExtensions = { ".jpg", ".jpeg", ".png", ".webp", ".bmp" };

    public static IXLWorksheet? ResolveWorksheet(
        IXLWorkbook workbook,
        IEnumerable<string> preferredSheetNames,
        bool allowDefaultFallback = true)
    {
        var preferredNames = preferredSheetNames
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Select(name => name.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        foreach (var preferredName in preferredNames)
        {
            var matchedWorksheet = workbook.Worksheets
                .FirstOrDefault(worksheet => string.Equals(
                    worksheet.Name.Trim(),
                    preferredName,
                    StringComparison.OrdinalIgnoreCase));
            if (matchedWorksheet is not null)
            {
                return matchedWorksheet;
            }
        }

        if (!allowDefaultFallback)
        {
            return null;
        }

        return workbook.Worksheets
            .FirstOrDefault(worksheet => string.Equals(
                worksheet.Name.Trim(),
                "분리추출후",
                StringComparison.OrdinalIgnoreCase))
            ?? workbook.Worksheets.FirstOrDefault();
    }

    public static string ResolveWorkingDirectory(string sourcePath, string exportRoot, Cafe24UploadOptions options)
    {
        var searchRoots = new List<string>();
        // exportRoot를 우선 탐색 (LLM 결과 파일이 하위 폴더에 있을 수 있으므로)
        if (!string.IsNullOrWhiteSpace(exportRoot))
        {
            searchRoots.Add(exportRoot);
        }
        if (File.Exists(sourcePath))
        {
            var sourceDirectory = Path.GetDirectoryName(sourcePath);
            if (!string.IsNullOrWhiteSpace(sourceDirectory))
            {
                searchRoots.Add(sourceDirectory);
            }
        }
        if (!string.IsNullOrWhiteSpace(options.ExportDir))
        {
            searchRoots.Add(options.ExportDir);
        }

        foreach (var root in searchRoots.Where(Directory.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var direct = FindLatestFileInDirectory(root, "업로드용_*.xlsx");
            if (direct is not null)
            {
                return Path.GetDirectoryName(direct) ?? root;
            }

            var recursive = FindLatestFileBySearch(root, "업로드용_*.xlsx");
            if (recursive is not null)
            {
                return Path.GetDirectoryName(recursive) ?? root;
            }
        }

        throw new FileNotFoundException("업로드용 엑셀 파일을 찾지 못했습니다. 먼저 업로드용 엑셀을 생성해 주세요.");
    }

    public static string ResolveImageRoot(string workingDirectory, Cafe24UploadOptions options, string dateTag)
    {
        var candidates = new List<string?>
        {
            options.ImageRoot,
            Path.Combine(workingDirectory, "listing_images", dateTag),
            Path.Combine(workingDirectory, "listing_images")
        };

        var nestedListing = FindDirectoryByName(workingDirectory, "listing_images");
        if (!string.IsNullOrWhiteSpace(nestedListing))
        {
            candidates.Add(Path.Combine(nestedListing, dateTag));
            candidates.Add(nestedListing);
        }

        foreach (var candidate in candidates)
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                continue;
            }

            var normalized = NormalizeListingPath(candidate);
            if (!Directory.Exists(normalized))
            {
                continue;
            }

            var resolved = DescendIntoSingleDateFolder(normalized);
            if (HasGsFolders(resolved))
            {
                return resolved;
            }
        }

        throw new DirectoryNotFoundException("listing_images 폴더를 찾지 못했습니다.");
    }

    public static List<DirectoryInfo> GetGsFolders(string imageRoot)
    {
        return SafeEnumerateDirectories(imageRoot)
            .Where(directory => GsFolderRegex.IsMatch(directory.Name))
            .Where(directory =>
            {
                // GS코드 A만 처리 (B, C 등은 스킵)
                var name = directory.Name.ToUpperInvariant();
                if (name.Length == 9) return true; // GS1234567 (접미사 없음) → 포함
                if (name.Length > 9)
                {
                    var suffix = name[9];
                    return suffix == 'A'; // A만 포함, B/C/... 제외
                }
                return true;
            })
            .OrderBy(directory => directory.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>GS코드 → (상품명, 검색어설정, 검색키워드) 매핑을 엑셀에서 읽어옵니다.</summary>
    public static Dictionary<string, (string ProductName, string ProductTag, string SearchKeyword)> ReadProductKeywordData(
        string workbookPath,
        string sheetName = "분리추출후",
        bool allowDefaultFallback = true)
    {
        var result = new Dictionary<string, (string, string, string)>(StringComparer.OrdinalIgnoreCase);
        using var workbook = WorkbookFileLoader.OpenReadOnly(workbookPath);
        var worksheet = ResolveWorksheet(workbook, new[] { sheetName }, allowDefaultFallback);
        if (worksheet is null)
        {
            return result;
        }

        var headers = BuildHeaderMap(worksheet);

        if (!headers.TryGetValue("상품명", out var nameCol)) return result;
        headers.TryGetValue("검색어설정", out var tagCol);
        headers.TryGetValue("검색키워드", out var kwCol);
        headers.TryGetValue("자체 상품코드", out var codeCol);

        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
        for (var row = 2; row <= lastRow; row++)
        {
            var name = worksheet.Cell(row, nameCol).GetFormattedString().Trim();
            if (string.IsNullOrWhiteSpace(name)) continue;

            // GS코드 추출: 상품명 → 자체 상품코드 순으로 시도
            var gsMatch = System.Text.RegularExpressions.Regex.Match(name, @"GS\d{7,9}[A-Z]?");
            var gs9 = gsMatch.Success ? gsMatch.Value : "";
            if (string.IsNullOrWhiteSpace(gs9) && codeCol > 0)
            {
                var code = worksheet.Cell(row, codeCol).GetFormattedString().Trim();
                var codeMatch = System.Text.RegularExpressions.Regex.Match(code, @"GS\d{7,9}[A-Z]?");
                gs9 = codeMatch.Success ? codeMatch.Value : "";
            }
            if (gs9.Length > 9) gs9 = gs9[..9];
            if (string.IsNullOrWhiteSpace(gs9)) continue;

            var tag = tagCol > 0 ? worksheet.Cell(row, tagCol).GetFormattedString().Trim() : "";
            var kw = kwCol > 0 ? worksheet.Cell(row, kwCol).GetFormattedString().Trim() : "";

            if (!result.ContainsKey(gs9))
                result[gs9] = (name, tag, kw);
        }

        return result;
    }

    public static List<string> ReadUploadProductNames(
        string workbookPath,
        string sheetName = "분리추출후",
        bool allowDefaultFallback = true)
    {
        using var workbook = WorkbookFileLoader.OpenReadOnly(workbookPath);
        var worksheet = ResolveWorksheet(workbook, new[] { sheetName }, allowDefaultFallback);
        if (worksheet is null)
        {
            return new List<string>();
        }

        var headers = BuildHeaderMap(worksheet);
        if (!headers.TryGetValue("상품명", out var productNameColumn))
        {
            return new List<string>();
        }

        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
        var productNames = new List<string>();
        for (var row = 2; row <= lastRow; row++)
        {
            var productName = worksheet.Cell(row, productNameColumn).GetFormattedString().Trim();
            if (!string.IsNullOrWhiteSpace(productName))
            {
                productNames.Add(productName);
            }
        }

        return productNames;
    }

    public static Dictionary<string, List<OptionSupplyItem>> LoadOptionSupplyPrices(string workbookPath)
    {
        var result = new Dictionary<string, List<OptionSupplyItem>>(StringComparer.OrdinalIgnoreCase);
        using var workbook = WorkbookFileLoader.OpenReadOnly(workbookPath);

        IXLWorksheet? worksheet = null;
        foreach (var name in new[] { "분리추출전", "분리추출후" })
        {
            if (workbook.Worksheets.Contains(name))
            {
                worksheet = workbook.Worksheet(name);
                break;
            }
        }
        if (worksheet is null)
        {
            return result;
        }

        var headers = BuildHeaderMap(worksheet);
        var supplyColumn = headers.FirstOrDefault(pair => pair.Key.Contains("공급가", StringComparison.OrdinalIgnoreCase));
        if (string.IsNullOrWhiteSpace(supplyColumn.Key) || !headers.TryGetValue("상품명", out var productNameColumn))
        {
            return result;
        }

        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
        for (var row = 2; row <= lastRow; row++)
        {
            var productName = worksheet.Cell(row, productNameColumn).GetFormattedString().Trim();
            if (string.IsNullOrWhiteSpace(productName))
            {
                continue;
            }

            var match = GsCodeRegex.Match(productName);
            if (!match.Success)
            {
                continue;
            }

            var gs = match.Groups[1].Value.ToUpperInvariant();
            var gs9 = gs.Length >= 9 ? gs[..9] : gs;
            var suffix = productName[(match.Index + match.Length)..].Trim();
            var supplyPrice = ParseDecimal(worksheet.Cell(row, supplyColumn.Value).GetFormattedString(), 0m);

            if (!result.TryGetValue(gs9, out var items))
            {
                items = new List<OptionSupplyItem>();
                result[gs9] = items;
            }
            items.Add(new OptionSupplyItem(suffix, supplyPrice));
        }

        return result;
    }

    public static PriceReviewData LoadPriceReview(string? path)
    {
        var result = new PriceReviewData();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return result;
        }

        using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
        var root = document.RootElement;

        if (root.TryGetProperty("checked_gs", out var checkedGs) && checkedGs.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in checkedGs.EnumerateArray())
            {
                var value = item.GetString()?.Trim().ToUpperInvariant();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    result.CheckedGs.Add(value.Length >= 9 ? value[..9] : value);
                }
            }
        }

        if (root.TryGetProperty("image_selections", out var imageSelections) && imageSelections.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in imageSelections.EnumerateObject())
            {
                var key = property.Name.Trim().ToUpperInvariant();
                if (property.Value.ValueKind != JsonValueKind.Object)
                {
                    continue;
                }

                int? mainIndex = null;
                var additional = new List<int>();
                if (property.Value.TryGetProperty("main", out var mainElement) && mainElement.ValueKind == JsonValueKind.Number && mainElement.TryGetInt32(out var parsedMain))
                {
                    mainIndex = parsedMain;
                }
                if (property.Value.TryGetProperty("additional", out var addElement) && addElement.ValueKind == JsonValueKind.Array)
                {
                    foreach (var value in addElement.EnumerateArray())
                    {
                        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var parsedIndex))
                        {
                            additional.Add(parsedIndex);
                        }
                    }
                }

                result.ImageSelections[key.Length >= 9 ? key[..9] : key] = new ImageSelection(mainIndex, additional);
            }
        }

        if (root.TryGetProperty("edited_amounts", out var editedAmounts) && editedAmounts.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in editedAmounts.EnumerateObject())
            {
                var key = property.Name.Trim().ToUpperInvariant();
                if (property.Value.ValueKind != JsonValueKind.Array)
                {
                    continue;
                }

                var amounts = new List<decimal>();
                foreach (var item in property.Value.EnumerateArray())
                {
                    if (item.ValueKind == JsonValueKind.Number && item.TryGetDecimal(out var decimalValue))
                    {
                        amounts.Add(decimalValue);
                    }
                    else if (item.ValueKind == JsonValueKind.String && decimal.TryParse(item.GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var stringValue))
                    {
                        amounts.Add(stringValue);
                    }
                }

                result.EditedAmounts[key.Length >= 9 ? key[..9] : key] = amounts;
            }
        }

        return result;
    }

    public static Cafe24Product? FindMatchingProduct(
        string gs,
        IReadOnlyList<string> uploadNames,
        IReadOnlyList<Cafe24Product> products,
        IReadOnlyDictionary<string, Cafe24Product> productsByName,
        string matchMode,
        int matchPrefix)
    {
        foreach (var product in products)
        {
            if (!string.IsNullOrWhiteSpace(product.CustomProductCode) && product.CustomProductCode.Contains(gs, StringComparison.OrdinalIgnoreCase))
            {
                return product;
            }
        }

        foreach (var uploadName in uploadNames.Where(name => name.Contains(gs, StringComparison.OrdinalIgnoreCase)))
        {
            if (productsByName.TryGetValue(uploadName, out var exact))
            {
                return exact;
            }
        }

        var normalizedGs = NormalizeName(gs);
        foreach (var product in products)
        {
            if (string.IsNullOrWhiteSpace(product.ProductName))
            {
                continue;
            }
            if (product.ProductName.Contains(gs, StringComparison.OrdinalIgnoreCase))
            {
                return product;
            }

            var normalizedName = NormalizeName(product.ProductName);
            switch (matchMode)
            {
                case "EXACT":
                    if (string.Equals(normalizedName, normalizedGs, StringComparison.OrdinalIgnoreCase))
                    {
                        return product;
                    }
                    break;
                case "CONTAINS":
                    if (!string.IsNullOrWhiteSpace(normalizedGs) && normalizedName.Contains(normalizedGs, StringComparison.OrdinalIgnoreCase))
                    {
                        return product;
                    }
                    break;
                default:
                    var prefix = normalizedName.Length <= matchPrefix ? normalizedName : normalizedName[..matchPrefix];
                    if (!string.IsNullOrWhiteSpace(normalizedGs) && prefix.Contains(normalizedGs, StringComparison.OrdinalIgnoreCase))
                    {
                        return product;
                    }
                    break;
            }
        }

        return null;
    }

    public static (string? MainImagePath, List<string> AdditionalImagePaths) PickImages(string folderPath, int mainIndex, int addStart, int addMax)
    {
        var files = GetImageFiles(folderPath);
        if (mainIndex <= 0 || files.Count < mainIndex)
        {
            return (null, new List<string>());
        }

        return (files[mainIndex - 1], files.Skip(Math.Max(addStart - 1, 0)).Take(Math.Max(addMax, 0)).ToList());
    }

    public static (string? MainImagePath, List<string> AdditionalImagePaths) PickImagesBySelection(string folderPath, ImageSelection selection)
    {
        var files = GetImageFiles(folderPath);
        if (!selection.MainIndex.HasValue || selection.MainIndex.Value < 0 || selection.MainIndex.Value >= files.Count)
        {
            return (null, new List<string>());
        }

        var additional = selection.AdditionalIndices
            .Where(index => index >= 0 && index < files.Count)
            .Select(index => files[index])
            .ToList();
        return (files[selection.MainIndex.Value], additional);
    }

    public static PriceCalculation CalcOptionPrices(IReadOnlyList<decimal> supplyPrices)
    {
        if (supplyPrices.Count == 0)
        {
            return new PriceCalculation(0, 0, Array.Empty<decimal>());
        }

        List<int> selling;
        var minimum = supplyPrices.Min();
        if (minimum <= 100m)
        {
            var unique = supplyPrices.Distinct().OrderBy(value => value).ToList();
            var basePrice = Ceil10(unique[0] * GetMultiplier(unique[0]));
            var mapped = new Dictionary<decimal, int>();
            for (var index = 0; index < unique.Count; index++)
            {
                mapped[unique[index]] = basePrice + (index * 10);
            }
            selling = supplyPrices.Select(price => mapped[price]).ToList();
        }
        else
        {
            var unique = supplyPrices.Distinct().OrderBy(value => value).ToList();
            var mapped = unique.ToDictionary(price => price, price => Ceil100(price * GetMultiplier(price)));
            var grouped = unique.GroupBy(price => mapped[price]).Where(group => group.Count() > 1);
            foreach (var group in grouped)
            {
                var ordered = group.OrderBy(value => value).ToList();
                for (var index = 1; index < ordered.Count; index++)
                {
                    var difference = (ordered[index] - ordered[0]) * 2m;
                    var adjustment = difference <= 49m ? 50 : Ceil100(difference);
                    mapped[ordered[index]] = group.Key + adjustment;
                }
            }
            selling = supplyPrices.Select(price => mapped[price]).ToList();
        }

        var baseSelling = selling[0];
        var consumer = Ceil100(baseSelling * 1.2m);
        var additional = selling.Select(price => (decimal)(price - baseSelling)).ToList();
        return new PriceCalculation(baseSelling, consumer, additional);
    }

    public static string WriteLogWorkbook(IReadOnlyList<Dictionary<string, string>> rows, string workingDirectory, string? configuredPath)
    {
        var logPath = ResolveLogPath(workingDirectory, configuredPath);
        Directory.CreateDirectory(Path.GetDirectoryName(logPath) ?? workingDirectory);

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("업로드로그");
        var headers = new[] { "GS", "PRODUCT_NO", "STATUS", "MAIN", "ADD_COUNT", "ADD_FILES", "SELECT_MAIN_IDX", "SELECT_ADD_IDX", "PRICE", "ERROR" };
        for (var column = 0; column < headers.Length; column++)
        {
            worksheet.Cell(1, column + 1).Value = headers[column];
            worksheet.Cell(1, column + 1).Style.Font.Bold = true;
        }
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];
            for (var column = 0; column < headers.Length; column++)
            {
                row.TryGetValue(headers[column], out var value);
                worksheet.Cell(rowIndex + 2, column + 1).Value = value ?? string.Empty;
            }
        }
        worksheet.SheetView.FreezeRows(1);
        worksheet.ColumnsUsed().AdjustToContents();
        workbook.SaveAs(logPath);
        return logPath;
    }

    public static Dictionary<string, string> CreateLogRow(
        string gs,
        string productNo = "",
        string status = "",
        string? mainImagePath = null,
        IReadOnlyList<string>? additionalImagePaths = null,
        ImageSelection? selection = null,
        string priceStatus = "",
        string error = "")
    {
        return new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["GS"] = gs,
            ["PRODUCT_NO"] = productNo,
            ["STATUS"] = status,
            ["MAIN"] = string.IsNullOrWhiteSpace(mainImagePath) ? string.Empty : Path.GetFileName(mainImagePath),
            ["ADD_COUNT"] = (additionalImagePaths?.Count ?? 0).ToString(CultureInfo.InvariantCulture),
            ["ADD_FILES"] = additionalImagePaths is null ? string.Empty : string.Join(",", additionalImagePaths.Select(Path.GetFileName)),
            ["SELECT_MAIN_IDX"] = selection?.MainIndex?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            ["SELECT_ADD_IDX"] = selection is null ? string.Empty : string.Join(",", selection.AdditionalIndices),
            ["PRICE"] = priceStatus,
            ["ERROR"] = error
        };
    }

    public static string? FindLatestFileInDirectory(string directory, string pattern)
    {
        if (!Directory.Exists(directory))
        {
            return null;
        }

        return SafeEnumerateFiles(directory, pattern)
            .Select(path => new FileInfo(path))
            .OrderByDescending(file => file.LastWriteTimeUtc)
            .Select(file => file.FullName)
            .FirstOrDefault();
    }

    public static string? ExtractDateTag(string path)
    {
        var match = Regex.Match(Path.GetFileNameWithoutExtension(path), @"(20\d{6})");
        return match.Success ? match.Groups[1].Value : null;
    }

    public static string UnwrapMessage(Exception exception)
    {
        return exception.InnerException is null ? exception.Message : UnwrapMessage(exception.InnerException);
    }

    private static string? FindLatestFileBySearch(string root, string pattern)
    {
        if (!Directory.Exists(root))
        {
            return null;
        }
        try
        {
            return Directory.EnumerateFiles(root, pattern, SearchOption.AllDirectories)
                .Select(path => new FileInfo(path))
                .OrderByDescending(file => file.LastWriteTimeUtc)
                .Select(file => file.FullName)
                .FirstOrDefault();
        }
        catch
        {
            return null;
        }
    }

    private static string? FindDirectoryByName(string root, string name)
    {
        if (!Directory.Exists(root))
        {
            return null;
        }
        try
        {
            return Directory.EnumerateDirectories(root, name, SearchOption.AllDirectories).FirstOrDefault();
        }
        catch
        {
            return null;
        }
    }

    private static bool HasGsFolders(string path)
    {
        return SafeEnumerateDirectories(path).Any(directory => GsFolderRegex.IsMatch(directory.Name));
    }

    private static string DescendIntoSingleDateFolder(string path)
    {
        var current = path;
        while (true)
        {
            if (HasGsFolders(current))
            {
                return current;
            }
            var subdirs = SafeEnumerateDirectories(current).ToList();
            if (subdirs.Count != 1)
            {
                return current;
            }
            current = subdirs[0].FullName;
        }
    }

    private static IEnumerable<DirectoryInfo> SafeEnumerateDirectories(string path)
    {
        try
        {
            return new DirectoryInfo(path).EnumerateDirectories();
        }
        catch
        {
            return Enumerable.Empty<DirectoryInfo>();
        }
    }

    private static IEnumerable<string> SafeEnumerateFiles(string path, string pattern)
    {
        try
        {
            return Directory.EnumerateFiles(path, pattern, SearchOption.TopDirectoryOnly);
        }
        catch
        {
            return Enumerable.Empty<string>();
        }
    }

    private static List<string> GetImageFiles(string folderPath)
    {
        return SafeEnumerateFiles(folderPath, "*.*")
            .Where(path => ImageExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static Dictionary<string, int> BuildHeaderMap(IXLWorksheet worksheet)
    {
        var result = new Dictionary<string, int>(StringComparer.Ordinal);
        var headerRow = worksheet.FirstRowUsed();
        if (headerRow is null)
        {
            return result;
        }
        var lastCell = headerRow.LastCellUsed();
        if (lastCell is null)
        {
            return result;
        }

        var lastColumn = lastCell.Address.ColumnNumber;
        for (var column = 1; column <= lastColumn; column++)
        {
            var header = worksheet.Cell(1, column).GetFormattedString().Trim();
            if (!string.IsNullOrWhiteSpace(header) && !result.ContainsKey(header))
            {
                result[header] = column;
            }
        }
        return result;
    }

    private static string ResolveLogPath(string workingDirectory, string? configuredPath)
    {
        if (string.IsNullOrWhiteSpace(configuredPath))
        {
            return Path.Combine(workingDirectory, $"cafe24_upload_log_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }
        if (Directory.Exists(configuredPath))
        {
            return Path.Combine(configuredPath, $"cafe24_upload_log_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }
        if (Path.GetExtension(configuredPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return configuredPath;
        }

        var directory = Path.GetDirectoryName(configuredPath);
        var stem = Path.GetFileNameWithoutExtension(configuredPath);
        directory = string.IsNullOrWhiteSpace(directory) ? workingDirectory : directory;
        stem = string.IsNullOrWhiteSpace(stem) ? "cafe24_upload_log" : stem;
        return Path.Combine(directory, $"{stem}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
    }

    private static string NormalizeListingPath(string path)
    {
        var normalized = Path.GetFullPath(path);
        var duplicate = $"listing_images{Path.DirectorySeparatorChar}listing_images";
        while (normalized.Contains(duplicate, StringComparison.OrdinalIgnoreCase))
        {
            normalized = normalized.Replace(duplicate, "listing_images", StringComparison.OrdinalIgnoreCase);
        }
        return normalized;
    }

    private static string NormalizeName(string value)
    {
        return Regex.Replace(value ?? string.Empty, @"[^0-9가-힣A-Za-z]", string.Empty);
    }

    private static int Ceil10(decimal value) => (int)(Math.Ceiling(value / 10m) * 10m);

    private static int Ceil100(decimal value) => (int)(Math.Ceiling(value / 100m) * 100m);

    private static decimal GetMultiplier(decimal supplyPrice)
    {
        if (supplyPrice >= 20000m)
        {
            return 1.6m;
        }
        if (supplyPrice >= 10000m)
        {
            return 1.8m;
        }
        return 2.0m;
    }

    private static decimal ParseDecimal(string? value, decimal fallback)
    {
        return decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var invariant)
            ? invariant
            : decimal.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out var current)
                ? current
                : fallback;
    }
}
