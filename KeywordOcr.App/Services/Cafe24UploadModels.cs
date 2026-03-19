using System;
using System.Collections.Generic;

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

internal sealed record Cafe24Variant(string VariantCode, IReadOnlyList<string> OptionValues)
{
    public string OptionSummary => string.Join(" / ", OptionValues);
}

internal sealed record OptionSupplyItem(string Suffix, decimal SupplyPrice);

internal sealed record ImageSelection(int? MainIndex, List<int> AdditionalIndices);

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