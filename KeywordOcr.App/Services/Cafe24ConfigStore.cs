using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace KeywordOcr.App.Services;

internal sealed class Cafe24ConfigStore
{
    private readonly string _v2Root;
    private readonly string _legacyRoot;

    public Cafe24ConfigStore(string v2Root, string legacyRoot)
    {
        _v2Root = v2Root;
        _legacyRoot = legacyRoot;
    }

    public string GetUploadConfigPath()
    {
        return Path.Combine(_v2Root, "cafe24_upload_config.txt");
    }

    public Cafe24TokenState LoadTokenState(string? preferredTokenFilePath = null)
    {
        var configuredTokenPath = ResolveConfiguredTokenPath();
        var sharedTokenJsonPath = Cafe24SharedTokenStore.GetDefaultPath();
        var desktopKeywordocr = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "keywordocr");
        var tokenFilePath = FindFirstExisting(
            sharedTokenJsonPath,
            ResolveOptionalPath(preferredTokenFilePath) ?? string.Empty,
            configuredTokenPath,
            Path.Combine(_v2Root, "cafe24_token.txt"),
            Path.Combine(_legacyRoot, "cafe24_token.txt"),
            Path.Combine(desktopKeywordocr, "cafe24_token.txt"));
        var envPath = FindFirstExisting(
            Path.Combine(_v2Root, ".env"),
            Path.Combine(_legacyRoot, ".env"));

        var fileValues = string.IsNullOrWhiteSpace(tokenFilePath)
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : LoadTokenFile(tokenFilePath);
        var envValues = string.IsNullOrWhiteSpace(envPath)
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : LoadKeyValueFile(envPath);

        var config = new Cafe24TokenConfig
        {
            MallId = PickValue(fileValues, envValues, "MALL_ID", "CAFE24_MALL_ID"),
            AccessToken = PickValue(fileValues, envValues, "ACCESS_TOKEN", "CAFE24_ACCESS_TOKEN"),
            RefreshToken = PickValue(fileValues, envValues, "REFRESH_TOKEN", "CAFE24_REFRESH_TOKEN"),
            ClientId = PickValue(fileValues, envValues, "CLIENT_ID", "CAFE24_CLIENT_ID"),
            ClientSecret = PickValue(fileValues, envValues, "CLIENT_SECRET", "CAFE24_CLIENT_SECRET"),
            RedirectUri = PickValue(fileValues, envValues, "REDIRECT_URI", "CAFE24_REDIRECT_URI"),
            Scope = PickValue(fileValues, envValues, "SCOPE", "CAFE24_SCOPE"),
            ShopNo = PickValue(fileValues, envValues, "SHOP_NO", "CAFE24_SHOP_NO", "1"),
            ApiVersion = PickValue(fileValues, envValues, "API_VERSION", "CAFE24_API_VERSION", "2025-12-01")
        };

        var configPath = string.IsNullOrWhiteSpace(tokenFilePath)
            ? sharedTokenJsonPath
            : tokenFilePath;
        return new Cafe24TokenState(configPath, config);
    }

    public Cafe24TokenState LoadTokenStateB()
    {
        var sharedTokenJsonPath = Cafe24SharedTokenStore.GetDefaultPathB();
        if (!File.Exists(sharedTokenJsonPath))
            throw new FileNotFoundException("B마켓 토큰 파일을 찾지 못했습니다.", sharedTokenJsonPath);

        var fileValues = Cafe24SharedTokenStore.LoadAsKeyValues(sharedTokenJsonPath);
        var envValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var config = new Cafe24TokenConfig
        {
            MallId = PickValue(fileValues, envValues, "MALL_ID", "CAFE24_MALL_ID"),
            AccessToken = PickValue(fileValues, envValues, "ACCESS_TOKEN", "CAFE24_ACCESS_TOKEN"),
            RefreshToken = PickValue(fileValues, envValues, "REFRESH_TOKEN", "CAFE24_REFRESH_TOKEN"),
            ClientId = PickValue(fileValues, envValues, "CLIENT_ID", "CAFE24_CLIENT_ID"),
            ClientSecret = PickValue(fileValues, envValues, "CLIENT_SECRET", "CAFE24_CLIENT_SECRET"),
            RedirectUri = PickValue(fileValues, envValues, "REDIRECT_URI", "CAFE24_REDIRECT_URI"),
            Scope = PickValue(fileValues, envValues, "SCOPE", "CAFE24_SCOPE"),
            ShopNo = PickValue(fileValues, envValues, "SHOP_NO", "CAFE24_SHOP_NO", "1"),
            ApiVersion = PickValue(fileValues, envValues, "API_VERSION", "CAFE24_API_VERSION", "2025-12-01")
        };

        return new Cafe24TokenState(sharedTokenJsonPath, config);
    }

    public Cafe24UploadOptions LoadUploadOptions(string exportRoot)
    {
        var configPath = FindFirstExisting(
            Path.Combine(_v2Root, "cafe24_upload_config.txt"),
            Path.Combine(_legacyRoot, "cafe24_upload_config.txt"));
        var values = string.IsNullOrWhiteSpace(configPath)
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : LoadKeyValueFile(configPath);

        return new Cafe24UploadOptions
        {
            TokenFilePath = ResolveOptionalPath(GetConfigValue(values, "TOKEN_FILE")),
            DateTag = GetConfigValue(values, "DATE_TAG"),
            MainIndex = ParseInt(GetConfigValue(values, "MAIN_INDEX"), 2),
            AddStart = ParseInt(GetConfigValue(values, "ADD_START"), 3),
            AddMax = ParseInt(GetConfigValue(values, "ADD_MAX"), 10),
            ExportDir = !string.IsNullOrWhiteSpace(exportRoot)
                ? exportRoot
                : ResolveOptionalPath(GetConfigValue(values, "EXPORT_DIR")) ?? @"C:\code\exports",
            ImageRoot = ResolveOptionalPath(GetConfigValue(values, "IMAGE_ROOT")),
            RetryCount = ParseInt(GetConfigValue(values, "RETRY_COUNT"), 1),
            RetryDelaySeconds = ParseDouble(GetConfigValue(values, "RETRY_DELAY"), 1.0),
            LogPath = ResolveOptionalPath(GetConfigValue(values, "LOG_PATH")),
            MatchMode = string.IsNullOrWhiteSpace(GetConfigValue(values, "MATCH_MODE"))
                ? "PREFIX"
                : GetConfigValue(values, "MATCH_MODE").ToUpperInvariant(),
            MatchPrefix = ParseInt(GetConfigValue(values, "MATCH_PREFIX"), 20),
            GsListPath = ResolveOptionalPath(GetConfigValue(values, "GS_LIST")),
            PriceDataPath = ResolveOptionalPath(GetConfigValue(values, "PRICE_DATA"))
        };
    }

    public void SaveUploadOptions(string path, Cafe24UploadOptions options)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var lines = new List<string>
        {
            $"TOKEN_FILE={options.TokenFilePath ?? string.Empty}",
            $"DATE_TAG={options.DateTag}",
            $"MAIN_INDEX={options.MainIndex.ToString(CultureInfo.InvariantCulture)}",
            $"ADD_START={options.AddStart.ToString(CultureInfo.InvariantCulture)}",
            $"ADD_MAX={options.AddMax.ToString(CultureInfo.InvariantCulture)}",
            $"EXPORT_DIR={options.ExportDir}",
            $"IMAGE_ROOT={options.ImageRoot ?? string.Empty}",
            $"RETRY_COUNT={options.RetryCount.ToString(CultureInfo.InvariantCulture)}",
            $"RETRY_DELAY={options.RetryDelaySeconds.ToString(CultureInfo.InvariantCulture)}",
            $"LOG_PATH={options.LogPath ?? string.Empty}",
            $"MATCH_MODE={options.MatchMode}",
            $"MATCH_PREFIX={options.MatchPrefix.ToString(CultureInfo.InvariantCulture)}",
            $"GS_LIST={options.GsListPath ?? string.Empty}",
            $"PRICE_DATA={options.PriceDataPath ?? string.Empty}"
        };

        File.WriteAllLines(path, lines, new UTF8Encoding(false));
    }

    public void SaveTokenConfig(string path, Cafe24TokenConfig config)
    {
        if (string.Equals(Path.GetExtension(path), ".json", StringComparison.OrdinalIgnoreCase))
        {
            Cafe24SharedTokenStore.Save(path, config);
            return;
        }

        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var lines = new[]
        {
            $"MALL_ID={config.MallId}",
            $"ACCESS_TOKEN={config.AccessToken}",
            $"REFRESH_TOKEN={config.RefreshToken}",
            $"CLIENT_ID={config.ClientId}",
            $"CLIENT_SECRET={config.ClientSecret}",
            $"REDIRECT_URI={config.RedirectUri}",
            $"SCOPE={config.Scope}",
            $"SHOP_NO={config.ShopNo}",
            $"API_VERSION={config.ApiVersion}"
        };
        File.WriteAllLines(path, lines, new UTF8Encoding(false));
    }

    private string? ResolveOptionalPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        if (Path.IsPathRooted(path))
        {
            return path;
        }

        var candidates = new[]
        {
            Path.Combine(_v2Root, path),
            Path.Combine(_legacyRoot, path)
        };
        return candidates.FirstOrDefault(File.Exists)
            ?? candidates.FirstOrDefault(Directory.Exists)
            ?? Path.Combine(_v2Root, path);
    }

    private string ResolveConfiguredTokenPath()
    {
        var configPath = FindFirstExisting(
            Path.Combine(_v2Root, "cafe24_upload_config.txt"),
            Path.Combine(_legacyRoot, "cafe24_upload_config.txt"));
        if (string.IsNullOrWhiteSpace(configPath))
        {
            return string.Empty;
        }

        var values = LoadKeyValueFile(configPath);
        return ResolveOptionalPath(GetConfigValue(values, "TOKEN_FILE")) ?? string.Empty;
    }

    private static Dictionary<string, string> LoadTokenFile(string path)
    {
        return string.Equals(Path.GetExtension(path), ".json", StringComparison.OrdinalIgnoreCase)
            ? Cafe24SharedTokenStore.LoadAsKeyValues(path)
            : LoadKeyValueFile(path);
    }

    private static Dictionary<string, string> LoadKeyValueFile(string path)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var rawLine in File.ReadAllLines(path, Encoding.UTF8))
        {
            var line = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#", StringComparison.Ordinal))
            {
                continue;
            }

            var separatorIndex = line.IndexOf('=');
            if (separatorIndex <= 0)
            {
                continue;
            }

            var key = line[..separatorIndex].Trim();
            var value = line[(separatorIndex + 1)..].Trim().Trim('"').Trim('\'');
            if (!string.IsNullOrWhiteSpace(key))
            {
                values[key] = value;
            }
        }
        return values;
    }

    private static string FindFirstExisting(params string[] candidates)
    {
        return candidates.FirstOrDefault(path => File.Exists(path) || Directory.Exists(path)) ?? string.Empty;
    }

    private static string PickValue(
        IReadOnlyDictionary<string, string> primary,
        IReadOnlyDictionary<string, string> secondary,
        string primaryKey,
        string secondaryKey,
        string defaultValue = "")
    {
        if (primary.TryGetValue(primaryKey, out var primaryValue) && !string.IsNullOrWhiteSpace(primaryValue))
        {
            return primaryValue;
        }
        if (primary.TryGetValue(secondaryKey, out var primarySecondaryValue) && !string.IsNullOrWhiteSpace(primarySecondaryValue))
        {
            return primarySecondaryValue;
        }
        if (secondary.TryGetValue(secondaryKey, out var secondaryValue) && !string.IsNullOrWhiteSpace(secondaryValue))
        {
            return secondaryValue;
        }
        if (secondary.TryGetValue(primaryKey, out var secondaryPlainValue) && !string.IsNullOrWhiteSpace(secondaryPlainValue))
        {
            return secondaryPlainValue;
        }
        return defaultValue;
    }

    private static string GetConfigValue(IReadOnlyDictionary<string, string> values, string key)
    {
        return values.TryGetValue(key, out var value) ? value : string.Empty;
    }

    private static int ParseInt(string? value, int fallback)
    {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : fallback;
    }

    private static double ParseDouble(string? value, double fallback)
    {
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) ? parsed : fallback;
    }
}


