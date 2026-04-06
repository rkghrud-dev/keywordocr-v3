using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace KeywordOcr.App.Services;

internal sealed class Cafe24MarketDataService
{
    private readonly Cafe24ConfigStore _configStore;
    private readonly Cafe24ApiClient _apiClient = new();
    private readonly Cafe24TokenState _tokenState;
    private readonly Dictionary<string, Cafe24Product> _productsByCode;
    private readonly Dictionary<string, Cafe24Product> _productsByName;
    private readonly Dictionary<int, Cafe24ProductSnapshot?> _snapshotCache = new();
    private readonly Action<string> _log;

    private Cafe24MarketDataService(
        string appRoot,
        Cafe24TokenState tokenState,
        IReadOnlyList<Cafe24Product> products,
        Action<string> log)
    {
        _configStore = new Cafe24ConfigStore(appRoot, appRoot);
        _tokenState = tokenState;
        _log = log;
        _productsByCode = products
            .Where(p => !string.IsNullOrWhiteSpace(p.CustomProductCode))
            .GroupBy(p => p.CustomProductCode, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        _productsByName = products
            .Where(p => !string.IsNullOrWhiteSpace(p.ProductName))
            .GroupBy(p => p.ProductName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
    }

    public static async Task<Cafe24MarketDataService?> TryCreateAsync(
        string sourcePath,
        Action<string> log,
        CancellationToken ct)
    {
        var appRoot = FindAppRoot(sourcePath);

        try
        {
            var configStore = new Cafe24ConfigStore(appRoot, appRoot);
            var tokenState = configStore.LoadTokenState();
            var apiClient = new Cafe24ApiClient();
            var products = await ExecuteWithRefreshAsync(configStore, tokenState, apiClient,
                cfg => apiClient.GetProductsAsync(cfg, false, ct), ct);

            log($"Cafe24 기준상품 인덱스 로드: {products.Count}개");
            return new Cafe24MarketDataService(appRoot, tokenState, products, log);
        }
        catch (Exception ex)
        {
            log($"Cafe24 기준상품 연동 생략: {ShortError(ex.Message)}");
            return null;
        }
    }

    public async Task<bool> TryApplyAsync(Dictionary<string, object?> row, CancellationToken ct)
    {
        var product = MatchProduct(row);
        if (product is null)
            return false;

        var snapshot = await GetSnapshotAsync(product.ProductNo, ct);
        if (snapshot is null)
            return false;

        var applied = false;

        if (snapshot.ListingImageUrls.Count > 0)
        {
            row["_cafe24_image_urls"] = snapshot.ListingImageUrls.ToList();
            applied = true;
        }

        if (!string.IsNullOrWhiteSpace(snapshot.DescriptionHtml))
        {
            row["상품 상세설명"] = snapshot.DescriptionHtml;
            row["상세설명"] = snapshot.DescriptionHtml;
            applied = true;
        }

        var optionInput = BuildOptionInput(snapshot.Variants);
        if (!string.IsNullOrWhiteSpace(optionInput))
        {
            row["옵션사용"] = "Y";
            row["품목구성방식"] = "T";
            row["품목 구성방식"] = "T";
            row["옵션구성방식"] = "T";
            row["옵션 구성방식"] = "T";
            row["옵션표시방식"] = "S";
            row["옵션 표시방식"] = "S";
            row["옵션입력"] = optionInput;
            row["옵션추가금"] = BuildOptionAdditionals(snapshot.Variants);
            applied = true;
        }

        if (applied)
        {
            row["_cafe24_product_no"] = snapshot.ProductNo;
            _log($"  Cafe24 적용: {snapshot.ProductNo} | 이미지 {snapshot.ListingImageUrls.Count}장 | 옵션 {CountUsableVariants(snapshot.Variants)}개");
        }

        return applied;
    }

    private Cafe24Product? MatchProduct(IReadOnlyDictionary<string, object?> row)
    {
        var customCode = GetRowString(row, "자체 상품코드");
        if (!string.IsNullOrWhiteSpace(customCode) && _productsByCode.TryGetValue(customCode, out var byCode))
            return byCode;

        var productName = GetRowString(row, "상품명");
        if (!string.IsNullOrWhiteSpace(productName) && _productsByName.TryGetValue(productName, out var byName))
            return byName;

        return null;
    }

    private async Task<Cafe24ProductSnapshot?> GetSnapshotAsync(int productNo, CancellationToken ct)
    {
        if (_snapshotCache.TryGetValue(productNo, out var cached))
            return cached;

        try
        {
            var snapshot = await ExecuteWithRefreshAsync(_configStore, _tokenState, _apiClient,
                cfg => _apiClient.GetProductSnapshotAsync(cfg, productNo, ct), ct);
            _snapshotCache[productNo] = snapshot;
            return snapshot;
        }
        catch (Exception ex)
        {
            _log($"  Cafe24 상품조회 실패: {productNo} | {ShortError(ex.Message)}");
            _snapshotCache[productNo] = null;
            return null;
        }
    }

    private static async Task<T> ExecuteWithRefreshAsync<T>(
        Cafe24ConfigStore configStore,
        Cafe24TokenState tokenState,
        Cafe24ApiClient apiClient,
        Func<Cafe24TokenConfig, Task<T>> action,
        CancellationToken ct)
    {
        try
        {
            return await action(tokenState.Config);
        }
        catch (Cafe24TokenExpiredException)
        {
            await apiClient.RefreshAccessTokenAsync(tokenState.Config, ct);
            configStore.SaveTokenConfig(tokenState.ConfigPath, tokenState.Config);
            return await action(tokenState.Config);
        }
    }

    private static string FindAppRoot(string sourcePath)
    {
        var current = File.Exists(sourcePath)
            ? Path.GetDirectoryName(sourcePath)
            : sourcePath;

        while (!string.IsNullOrWhiteSpace(current))
        {
            if (Directory.Exists(Path.Combine(current, "KeywordOcr.App"))
                && Directory.Exists(Path.Combine(current, "backend")))
            {
                return current;
            }

            current = Directory.GetParent(current)?.FullName;
        }

        return AppContext.BaseDirectory;
    }

    private static string GetRowString(IReadOnlyDictionary<string, object?> row, string key)
    {
        if (!row.TryGetValue(key, out var value) || value is null)
            return string.Empty;

        return value.ToString()?.Trim() ?? string.Empty;
    }

    private static int CountUsableVariants(IReadOnlyList<Cafe24Variant> variants)
    {
        return variants.Count(v => !string.IsNullOrWhiteSpace(v.OptionSummary));
    }

    private static string BuildOptionInput(IReadOnlyList<Cafe24Variant> variants)
    {
        var values = variants
            .Where(v => !string.IsNullOrWhiteSpace(v.OptionSummary))
            .Select((variant, index) => NormalizeOptionLabel(variant.OptionSummary, index))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToList();

        return values.Count == 0 ? string.Empty : $"옵션{{{string.Join("|", values)}}}";
    }

    private static string BuildOptionAdditionals(IReadOnlyList<Cafe24Variant> variants)
    {
        var usable = variants
            .Where(v => !string.IsNullOrWhiteSpace(v.OptionSummary))
            .Select(v => Convert.ToInt32(Math.Round(v.AdditionalAmount, MidpointRounding.AwayFromZero)))
            .Select(v => v.ToString(CultureInfo.InvariantCulture))
            .ToList();

        return usable.Count == 0 ? string.Empty : string.Join("|", usable);
    }

    private static string NormalizeOptionLabel(string value, int index)
    {
        var trimmed = value.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
            return string.Empty;

        if (Regex.IsMatch(trimmed, @"^[A-Z]\s+"))
            return trimmed;

        var label = index < 26 ? ((char)('A' + index)).ToString() : $"OPT{index + 1}";
        return $"{label} {trimmed}";
    }

    private static string ShortError(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
            return "알 수 없는 오류";

        return message.Length > 140 ? message[..140] + "..." : message;
    }
}