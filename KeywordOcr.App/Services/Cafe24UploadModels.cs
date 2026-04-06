using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace KeywordOcr.App.Services;

public sealed record Cafe24UploadResult(
    string WorkingDirectory,
    string UploadWorkbookPath,
    string LogPath,
    int TotalCount,
    int SuccessCount,
    int ErrorCount,
    int SkippedCount);
public sealed record Cafe24CreateProductsResult(
    string WorkingDirectory,
    string UploadWorkbookPath,
    string LogPath,
    int TotalCount,
    int CreatedCount,
    int SkippedCount,
    int ErrorCount);

internal sealed class Cafe24TokenConfig
{
    public string MallId { get; set; } = string.Empty;

    public string AccessToken { get; set; } = string.Empty;

    public string RefreshToken { get; set; } = string.Empty;

    public string ClientId { get; set; } = string.Empty;

    public string ClientSecret { get; set; } = string.Empty;

    public string RedirectUri { get; set; } = string.Empty;

    public string Scope { get; set; } = string.Empty;

    public string ShopNo { get; set; } = "1";

    public string ApiVersion { get; set; } = "2025-12-01";
}

internal sealed class Cafe24TokenState
{
    public Cafe24TokenState(string configPath, Cafe24TokenConfig config)
    {
        ConfigPath = configPath;
        Config = config;
    }

    public string ConfigPath { get; }

    public Cafe24TokenConfig Config { get; }
}

public sealed class Cafe24UploadOptions
{
    public string? TokenFilePath { get; set; }

    public string DateTag { get; set; } = string.Empty;

    public int MainIndex { get; set; } = 2;

    public int AddStart { get; set; } = 3;

    public int AddMax { get; set; } = 10;

    public string ExportDir { get; set; } = @"C:\code\exports";

    public string? ImageRoot { get; set; }

    public int RetryCount { get; set; } = 1;

    public double RetryDelaySeconds { get; set; } = 1.0;

    public string? LogPath { get; set; }

    public string MatchMode { get; set; } = "PREFIX";

    public int MatchPrefix { get; set; } = 20;

    public string? GsListPath { get; set; }

    public string? PriceDataPath { get; set; }
}

internal sealed record Cafe24Product(int ProductNo, string ProductName, string CustomProductCode);

internal sealed record Cafe24Variant(string VariantCode, IReadOnlyList<string> OptionValues, decimal AdditionalAmount = 0m)
{
    public string OptionSummary => string.Join(" / ", OptionValues);
}

internal sealed record Cafe24ProductSnapshot(
    int ProductNo,
    string ProductName,
    string CustomProductCode,
    string DescriptionHtml,
    string? RepresentativeImageUrl,
    IReadOnlyList<string> AdditionalImageUrls,
    IReadOnlyList<Cafe24Variant> Variants)
{
    public IReadOnlyList<string> ListingImageUrls => BuildListingImageUrls();

    private IReadOnlyList<string> BuildListingImageUrls()
    {
        var urls = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (!string.IsNullOrWhiteSpace(RepresentativeImageUrl) && seen.Add(RepresentativeImageUrl))
            urls.Add(RepresentativeImageUrl);

        foreach (var url in AdditionalImageUrls)
        {
            if (!string.IsNullOrWhiteSpace(url) && seen.Add(url))
                urls.Add(url);
        }

        if (urls.Count <= 1 && !string.IsNullOrWhiteSpace(DescriptionHtml))
        {
            foreach (Match match in Regex.Matches(DescriptionHtml, @"<img[^>]+src=[""']([^""']+)", RegexOptions.IgnoreCase))
            {
                var imgUrl = match.Groups[1].Value.Trim();
                if (!string.IsNullOrWhiteSpace(imgUrl) && seen.Add(imgUrl))
                    urls.Add(imgUrl);
                if (urls.Count >= 9)
                    break;
            }
        }

        return urls;
    }
}

internal sealed record OptionSupplyItem(string Suffix, decimal SupplyPrice);

internal sealed record ImageSelection(int? MainIndex, List<int> AdditionalIndices, int? MainIndexB = null);

internal sealed class PriceReviewData
{
    public HashSet<string> CheckedGs { get; } = new(StringComparer.OrdinalIgnoreCase);

    public Dictionary<string, ImageSelection> ImageSelections { get; } = new(StringComparer.OrdinalIgnoreCase);

    public Dictionary<string, List<decimal>> EditedAmounts { get; } = new(StringComparer.OrdinalIgnoreCase);
}

internal sealed record PriceCalculation(int BaseSellingPrice, int ConsumerPrice, IReadOnlyList<decimal> AdditionalAmounts);

internal sealed class Cafe24TokenExpiredException : Exception
{
}


internal sealed class Cafe24ReauthenticationRequiredException : Exception
{
    public Cafe24ReauthenticationRequiredException(string message, Exception? innerException = null)
        : base(message, innerException)
    {
    }
}
