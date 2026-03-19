using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;

namespace KeywordOcr.App.Services;

internal static class Cafe24SharedTokenStore
{
    public static string GetDefaultPath()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "key",
            "cafe24_token.json");
    }

    public static Dictionary<string, string> LoadAsKeyValues(string path)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(path))
        {
            return values;
        }

        using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
        var root = document.RootElement;
        Add(root, values, "MallId", "MALL_ID");
        Add(root, values, "AccessToken", "ACCESS_TOKEN");
        Add(root, values, "RefreshToken", "REFRESH_TOKEN");
        Add(root, values, "ClientId", "CLIENT_ID");
        Add(root, values, "ClientSecret", "CLIENT_SECRET");
        Add(root, values, "RedirectUri", "REDIRECT_URI");
        Add(root, values, "ApiVersion", "API_VERSION");
        Add(root, values, "ShopNo", "SHOP_NO");
        Add(root, values, "Scope", "SCOPE");
        return values;
    }

    public static void Save(string path, Cafe24TokenConfig config)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var payload = new Dictionary<string, string?>
        {
            ["MallId"] = config.MallId,
            ["ClientId"] = config.ClientId,
            ["ClientSecret"] = config.ClientSecret,
            ["AccessToken"] = config.AccessToken,
            ["RefreshToken"] = config.RefreshToken,
            ["RedirectUri"] = config.RedirectUri,
            ["ApiVersion"] = config.ApiVersion,
            ["ShopNo"] = config.ShopNo,
            ["Scope"] = config.Scope,
            ["UpdatedAt"] = DateTime.Now.ToString("o")
        };

        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(path, json, new UTF8Encoding(false));
    }

    private static void Add(JsonElement root, IDictionary<string, string> values, string jsonKey, string targetKey)
    {
        if (!root.TryGetProperty(jsonKey, out var element))
        {
            return;
        }

        var value = element.ValueKind switch
        {
            JsonValueKind.String => element.GetString(),
            JsonValueKind.Number => element.GetRawText(),
            JsonValueKind.True => bool.TrueString,
            JsonValueKind.False => bool.FalseString,
            _ => null
        };

        if (!string.IsNullOrWhiteSpace(value))
        {
            values[targetKey] = value;
        }
    }
}