using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace KeywordOcr.App.Services;

public sealed class Cafe24CreateProductService
{
    private static readonly Regex GsCodeRegex = new(@"(GS\d{7}[A-Z0-9]*)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly Cafe24ConfigStore _configStore;
    private readonly Cafe24ApiClient _apiClient = new();

    public Cafe24CreateProductService(string v2Root, string legacyRoot)
    {
        _configStore = new Cafe24ConfigStore(v2Root, legacyRoot);
    }

    public async Task<Cafe24CreateProductsResult> CreateAsync(
        string sourcePath,
        string exportRoot,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var tokenState = _configStore.LoadTokenState();
        ValidateTokenConfig(tokenState.Config);

        var options = _configStore.LoadUploadOptions(exportRoot);
        var workingDirectory = Cafe24UploadSupport.ResolveWorkingDirectory(sourcePath, exportRoot, options);

        // sourcePath가 직접 xlsx 파일이면 그걸 우선 사용 (LLM 결과 파일)
        var uploadWorkbookPath = File.Exists(sourcePath) && sourcePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
            ? sourcePath
            : Cafe24UploadSupport.FindLatestFileInDirectory(workingDirectory, "업로드용_*.xlsx");
        if (uploadWorkbookPath is null)
        {
            throw new FileNotFoundException("업로드용 엑셀을 찾지 못했습니다. 먼저 업로드용 엑셀을 생성해 주세요.", workingDirectory);
        }

        var rows = ReadRows(uploadWorkbookPath);
        if (rows.Count == 0)
        {
            throw new InvalidDataException("업로드용 엑셀에 등록할 행이 없습니다.");
        }

        progress?.Report($"신규등록 기준 파일: {Path.GetFileName(uploadWorkbookPath)}");
        progress?.Report($"작업 폴더: {workingDirectory}");
        progress?.Report($"대상 행 수: {rows.Count}개");

        var priceReview = Cafe24UploadSupport.LoadPriceReview(options.PriceDataPath);
        var dateTag = Cafe24UploadSupport.ExtractDateTag(uploadWorkbookPath) ?? options.DateTag ?? DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
        var imageRoot = TryResolveImageRoot(workingDirectory, options, dateTag, progress);
        var existingProducts = await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.GetProductsAsync(cfg, false, cancellationToken), cancellationToken);
        var existingByName = existingProducts
            .Where(product => !string.IsNullOrWhiteSpace(product.ProductName))
            .GroupBy(product => product.ProductName, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

        var createdCount = 0;
        var skippedCount = 0;
        var errorCount = 0;
        var logRows = new List<Dictionary<string, string>>();

        for (var index = 0; index < rows.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var row = rows[index];
            var productName = GetValue(row, "상품명");
            var customProductCode = GetValue(row, "자체 상품코드");
            var gsCode = ExtractGsCode(row);
            progress?.Report($"[{index + 1}/{rows.Count}] {productName} 신규등록 준비");

            if (string.IsNullOrWhiteSpace(productName))
            {
                skippedCount += 1;
                logRows.Add(Cafe24UploadSupport.CreateLogRow(gsCode, status: "SKIP_NO_NAME", error: "상품명이 비어 있습니다."));
                continue;
            }

            var isDuplicate = MatchesExistingProduct(existingProducts, existingByName, productName, customProductCode, gsCode);

            var request = BuildCreateRequest(row);
            if (!request.TryGetValue("product_name", out var requestProductName) || string.IsNullOrWhiteSpace(requestProductName?.ToString()))
            {
                skippedCount += 1;
                logRows.Add(Cafe24UploadSupport.CreateLogRow(gsCode, status: "SKIP_INVALID", error: "API 요청에 필요한 상품명이 없습니다."));
                continue;
            }

            try
            {
                var productNo = await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.CreateProductAsync(cfg, request, cancellationToken), cancellationToken);
                if (productNo <= 0)
                {
                    throw new InvalidDataException("신규 등록 응답에서 product_no를 찾지 못했습니다.");
                }

                var priceStatus = "CREATE_ONLY";
                // ── 옵션 추가금액 설정 ──
                var optionAdditionals = GetValue(row, "옵션추가금");
                if (!string.IsNullOrWhiteSpace(optionAdditionals))
                {
                    priceStatus = await UpdateVariantPricesAsync(tokenState, productNo, optionAdditionals, progress, index, rows.Count, productName, cancellationToken);
                }

                var imageStatus = "NO_IMAGE";
                if (!string.IsNullOrWhiteSpace(imageRoot) && !string.IsNullOrWhiteSpace(gsCode))
                {
                    imageStatus = await UploadImagesAsync(tokenState, imageRoot, gsCode, productNo, options, priceReview, cancellationToken);
                }

                createdCount += 1;
                var statusLabel = isDuplicate ? "CREATED_DUP" : "CREATED";
                var dupNote = isDuplicate ? " (중복상품)" : "";
                logRows.Add(Cafe24UploadSupport.CreateLogRow(
                    gsCode,
                    productNo: productNo.ToString(CultureInfo.InvariantCulture),
                    status: statusLabel,
                    priceStatus: $"{priceStatus}|{imageStatus}",
                    error: isDuplicate ? "중복상품입니다." : ""));
                progress?.Report($"[{index + 1}/{rows.Count}] {productName} 신규등록 완료{dupNote} ({imageStatus})");

                existingProducts.Add(new Cafe24Product(productNo, productName, customProductCode));
                if (!existingByName.ContainsKey(productName))
                {
                    existingByName[productName] = new Cafe24Product(productNo, productName, customProductCode);
                }
            }
            catch (Cafe24ReauthenticationRequiredException)
            {
                throw;
            }
            catch (Exception ex)
            {
                errorCount += 1;
                logRows.Add(Cafe24UploadSupport.CreateLogRow(gsCode, status: "ERROR", error: Cafe24UploadSupport.UnwrapMessage(ex)));
                progress?.Report($"[{index + 1}/{rows.Count}] {productName} 신규등록 실패: {Cafe24UploadSupport.UnwrapMessage(ex)}");
            }
        }

        var logPath = Cafe24UploadSupport.WriteLogWorkbook(logRows, workingDirectory, null);
        progress?.Report($"신규등록 로그 저장: {logPath}");

        return new Cafe24CreateProductsResult(workingDirectory, uploadWorkbookPath, logPath, rows.Count, createdCount, skippedCount, errorCount);
    }

    private async Task<string> UpdateVariantPricesAsync(
        Cafe24TokenState tokenState,
        int productNo,
        string optionAdditionals,
        IProgress<string>? progress,
        int index,
        int totalCount,
        string productName,
        CancellationToken cancellationToken)
    {
        try
        {
            var amounts = optionAdditionals.Split('|')
                .Select(s => decimal.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : 0m)
                .ToList();

            // 모든 추가금이 0이면 스킵
            if (amounts.All(a => a == 0m))
                return "PRICE_ALL_ZERO";

            var variants = await ExecuteWithRefreshAsync(tokenState,
                cfg => _apiClient.GetVariantsAsync(cfg, productNo, cancellationToken), cancellationToken);

            if (variants.Count == 0)
                return "NO_VARIANTS";

            var updated = 0;
            for (var i = 0; i < Math.Min(variants.Count, amounts.Count); i++)
            {
                if (amounts[i] == 0m)
                    continue;

                await ExecuteWithRefreshAsync(tokenState,
                    cfg => _apiClient.UpdateVariantAsync(cfg, productNo, variants[i].VariantCode, amounts[i], cancellationToken),
                    cancellationToken);
                updated++;
            }

            progress?.Report($"[{index + 1}/{totalCount}] {productName} 옵션가격 {updated}건 설정");
            return $"PRICE_OK:{updated}";
        }
        catch (Exception ex)
        {
            progress?.Report($"[{index + 1}/{totalCount}] {productName} 옵션가격 실패: {Cafe24UploadSupport.UnwrapMessage(ex)}");
            return $"PRICE_ERR:{Cafe24UploadSupport.UnwrapMessage(ex)}";
        }
    }

    private async Task<string> UploadImagesAsync(
        Cafe24TokenState tokenState,
        string imageRoot,
        string gsCode,
        int productNo,
        Cafe24UploadOptions options,
        PriceReviewData priceReview,
        CancellationToken cancellationToken)
    {
        var folder = FindImageFolder(imageRoot, gsCode);
        if (folder is null)
        {
            return "NO_LOCAL_IMAGE";
        }

        var gs9 = gsCode.Length >= 9 ? gsCode[..9] : gsCode;
        var selection = priceReview.ImageSelections.TryGetValue(gs9, out var imageSelection) ? imageSelection : null;
        var (mainImagePath, additionalImagePaths) = selection is null
            ? Cafe24UploadSupport.PickImages(folder.FullName, options.MainIndex, options.AddStart, options.AddMax)
            : Cafe24UploadSupport.PickImagesBySelection(folder.FullName, selection);

        if (string.IsNullOrWhiteSpace(mainImagePath))
        {
            return "NO_MAIN_IMAGE";
        }

        await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.UploadMainImageAsync(cfg, productNo, mainImagePath, cancellationToken), cancellationToken);
        foreach (var imagePath in additionalImagePaths)
        {
            await ExecuteWithRefreshAsync(tokenState, cfg => _apiClient.UploadAdditionalImageAsync(cfg, productNo, imagePath, cancellationToken), cancellationToken);
        }

        return $"IMAGE_OK:{1 + additionalImagePaths.Count}";
    }

    private static Dictionary<string, object?> BuildCreateRequest(IReadOnlyDictionary<string, string> row)
    {
        var request = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["product_name"] = GetValue(row, "상품명"),
            ["display"] = ToCafe24Flag(GetValue(row, "진열상태"), "T"),
            ["selling"] = ToCafe24Flag(GetValue(row, "판매상태"), "T")
        };

        AddIfNotEmpty(request, "custom_product_code", GetValue(row, "자체 상품코드"));
        AddIfNotEmpty(request, "summary_description", GetValue(row, "상품 요약설명"));
        AddIfNotEmpty(request, "simple_description", GetValue(row, "상품 간략설명"));
        AddIfNotEmpty(request, "description", GetValue(row, "상품 상세설명"));
        var productTag = GetValue(row, "검색어설정");
        if (!string.IsNullOrWhiteSpace(productTag))
        {
            request["product_tag"] = productTag
                .Split(new[] { ',', '/' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim())
                .Where(t => t.Length > 0)
                .ToArray();
        }

        var price = ParseDecimal(GetValue(row, "판매가"), ParseDecimal(GetValue(row, "상품가"), 0m));
        var supplyPrice = ParseDecimal(GetValue(row, "공급가"), 0m);
        if (price > 0m)
        {
            request["price"] = price;
        }
        if (supplyPrice > 0m)
        {
            request["supply_price"] = supplyPrice;
        }

        var categoryNo = ParseInt(GetValue(row, "상품분류 번호"), 0);
        if (categoryNo > 0)
        {
            request["add_category_products"] = new[]
            {
                new Dictionary<string, object?>
                {
                    ["category_no"] = categoryNo,
                    ["recommend"] = ToCafe24Flag(GetValue(row, "상품분류 추천상품영역"), "F"),
                    ["new"] = ToCafe24Flag(GetValue(row, "상품분류 신상품영역"), "F")
                }
            };
        }

        // ── 옵션 설정 ──
        var optionUse = ToCafe24Flag(GetValue(row, "옵션사용"), "F");
        if (optionUse == "T")
        {
            request["has_option"] = "T";
            request["option_type"] = "T"; // 조합형

            var optionInput = GetValue(row, "옵션입력");
            var optionValues = ParseOptionInput(optionInput);
            if (optionValues.Count > 0)
            {
                request["options"] = new[]
                {
                    new Dictionary<string, object?>
                    {
                        ["name"] = "옵션",
                        ["value"] = optionValues.ToArray()
                    }
                };
            }
        }

        return request;
    }

    private static List<Dictionary<string, string>> ReadRows(string workbookPath)
    {
        using var workbook = WorkbookFileLoader.OpenReadOnly(workbookPath);
        var worksheet = workbook.Worksheets.Contains("분리추출후") ? workbook.Worksheet("분리추출후") : workbook.Worksheet(1);
        var headerRow = worksheet.FirstRowUsed();
        if (headerRow is null)
        {
            return new List<Dictionary<string, string>>();
        }

        var lastCell = headerRow.LastCellUsed();
        if (lastCell is null)
        {
            return new List<Dictionary<string, string>>();
        }

        var headers = Enumerable.Range(1, lastCell.Address.ColumnNumber)
            .Select(index => worksheet.Cell(1, index).GetFormattedString().Trim())
            .ToList();

        var rows = new List<Dictionary<string, string>>();
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
        for (var rowIndex = 2; rowIndex <= lastRow; rowIndex++)
        {
            var row = new Dictionary<string, string>(StringComparer.Ordinal);
            for (var column = 0; column < headers.Count; column++)
            {
                var header = headers[column];
                if (string.IsNullOrWhiteSpace(header))
                {
                    continue;
                }

                row[header] = worksheet.Cell(rowIndex, column + 1).GetFormattedString();
            }
            rows.Add(row);
        }

        return rows;
    }

    private static bool IsOptionProduct(IReadOnlyDictionary<string, string> row)
    {
        var optionUse = GetValue(row, "옵션사용");
        if (ToCafe24Flag(optionUse, "F") == "T")
        {
            return true;
        }

        return !string.IsNullOrWhiteSpace(GetValue(row, "옵션입력"));
    }

    private static bool MatchesExistingProduct(
        IReadOnlyList<Cafe24Product> existingProducts,
        IReadOnlyDictionary<string, Cafe24Product> existingByName,
        string productName,
        string customProductCode,
        string gsCode)
    {
        if (!string.IsNullOrWhiteSpace(productName) && existingByName.ContainsKey(productName))
        {
            return true;
        }

        return existingProducts.Any(product =>
            (!string.IsNullOrWhiteSpace(customProductCode) && string.Equals(product.CustomProductCode, customProductCode, StringComparison.OrdinalIgnoreCase))
            || (!string.IsNullOrWhiteSpace(gsCode) && product.CustomProductCode.Contains(gsCode, StringComparison.OrdinalIgnoreCase)));
    }

    private string? TryResolveImageRoot(string workingDirectory, Cafe24UploadOptions options, string dateTag, IProgress<string>? progress)
    {
        try
        {
            return Cafe24UploadSupport.ResolveImageRoot(workingDirectory, options, dateTag);
        }
        catch
        {
            progress?.Report("listing_images 폴더를 찾지 못해 이미지 업로드는 건너뜁니다.");
            return null;
        }
    }

    private static DirectoryInfo? FindImageFolder(string imageRoot, string gsCode)
    {
        var gs9 = gsCode.Length >= 9 ? gsCode[..9] : gsCode;
        var folders = Cafe24UploadSupport.GetGsFolders(imageRoot);
        return folders.FirstOrDefault(folder => string.Equals(folder.Name, gsCode, StringComparison.OrdinalIgnoreCase))
            ?? folders.FirstOrDefault(folder => string.Equals(folder.Name, gs9, StringComparison.OrdinalIgnoreCase))
            ?? folders.FirstOrDefault(folder => folder.Name.StartsWith(gs9, StringComparison.OrdinalIgnoreCase));
    }

    private async Task<T> ExecuteWithRefreshAsync<T>(Cafe24TokenState tokenState, Func<Cafe24TokenConfig, Task<T>> action, CancellationToken cancellationToken)
    {
        try
        {
            return await action(tokenState.Config);
        }
        catch (Cafe24TokenExpiredException)
        {
            await _apiClient.RefreshAccessTokenAsync(tokenState.Config, cancellationToken);
            _configStore.SaveTokenConfig(tokenState.ConfigPath, tokenState.Config);
            return await action(tokenState.Config);
        }
    }

    private async Task ExecuteWithRefreshAsync(Cafe24TokenState tokenState, Func<Cafe24TokenConfig, Task> action, CancellationToken cancellationToken)
    {
        await ExecuteWithRefreshAsync(tokenState, async config =>
        {
            await action(config);
            return true;
        }, cancellationToken);
    }

    private static string ExtractGsCode(IReadOnlyDictionary<string, string> row)
    {
        var values = new[] { GetValue(row, "자체 상품코드"), GetValue(row, "상품명") };
        foreach (var value in values)
        {
            var match = GsCodeRegex.Match(value);
            if (match.Success)
            {
                return match.Groups[1].Value.ToUpperInvariant();
            }
        }
        return string.Empty;
    }

    /// <summary>옵션입력 "옵션{A 설명|B 설명|C 설명}" → ["A 설명", "B 설명", "C 설명"]</summary>
    private static List<string> ParseOptionInput(string optionInput)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(optionInput))
            return result;

        var match = Regex.Match(optionInput, @"\{(.+)\}");
        if (!match.Success)
            return result;

        var inner = match.Groups[1].Value;
        foreach (var part in inner.Split('|'))
        {
            var trimmed = part.Trim();
            if (!string.IsNullOrWhiteSpace(trimmed))
                result.Add(trimmed);
        }

        return result;
    }

    private static string GetValue(IReadOnlyDictionary<string, string> row, string key)
    {
        return row.TryGetValue(key, out var value) ? value?.Trim() ?? string.Empty : string.Empty;
    }

    private static void AddIfNotEmpty(IDictionary<string, object?> request, string key, string value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            request[key] = value;
        }
    }

    private static string ToCafe24Flag(string value, string fallback)
    {
        return value.Trim().ToUpperInvariant() switch
        {
            "Y" or "T" or "TRUE" or "1" => "T",
            "N" or "F" or "FALSE" or "0" => "F",
            _ => fallback
        };
    }

    private static decimal ParseDecimal(string value, decimal fallback)
    {
        if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var invariant))
        {
            return invariant;
        }
        if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out var current))
        {
            return current;
        }
        return fallback;
    }

    private static int ParseInt(string value, int fallback)
    {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : fallback;
    }

    private static void ValidateTokenConfig(Cafe24TokenConfig config)
    {
        if (string.IsNullOrWhiteSpace(config.MallId) || string.IsNullOrWhiteSpace(config.AccessToken))
        {
            throw new InvalidDataException("Cafe24 설정을 찾지 못했습니다. cafe24_token.txt에 MALL_ID와 ACCESS_TOKEN이 필요합니다.");
        }
    }
}
