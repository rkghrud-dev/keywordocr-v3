using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace KeywordOcr.App.Services;

internal sealed class Cafe24ApiClient
{
    private const string DefaultOAuthScope = "mall.read_order,mall.write_order,mall.read_shipping,mall.write_shipping,mall.read_product,mall.write_product";

    private readonly HttpClient _httpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(60)
    };

    public string BuildAuthorizeUrl(Cafe24TokenConfig config, string state = "keywordocr")
    {
        if (string.IsNullOrWhiteSpace(config.MallId)
            || string.IsNullOrWhiteSpace(config.ClientId)
            || string.IsNullOrWhiteSpace(config.RedirectUri))
        {
            throw new InvalidDataException("Cafe24 다시 인증에 필요한 MALL_ID / CLIENT_ID / REDIRECT_URI 설정이 없습니다.");
        }

        var query = new Dictionary<string, string>
        {
            ["response_type"] = "code",
            ["client_id"] = config.ClientId,
            ["redirect_uri"] = config.RedirectUri,
            ["scope"] = string.IsNullOrWhiteSpace(config.Scope) ? DefaultOAuthScope : config.Scope,
            ["state"] = string.IsNullOrWhiteSpace(state) ? "keywordocr" : state
        };

        return $"https://{config.MallId}.cafe24api.com/api/v2/oauth/authorize?{BuildQueryString(query)}";
    }

    public async Task ExchangeAuthorizationCodeAsync(Cafe24TokenConfig config, string authorizationCode, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(authorizationCode))
        {
            throw new InvalidDataException("Cafe24 authorization code가 비어 있습니다.");
        }

        if (string.IsNullOrWhiteSpace(config.MallId)
            || string.IsNullOrWhiteSpace(config.ClientId)
            || string.IsNullOrWhiteSpace(config.ClientSecret)
            || string.IsNullOrWhiteSpace(config.RedirectUri))
        {
            throw new InvalidDataException("Cafe24 다시 인증에 필요한 MALL_ID / CLIENT_ID / CLIENT_SECRET / REDIRECT_URI 설정이 없습니다.");
        }

        var url = $"https://{config.MallId}.cafe24api.com/api/v2/oauth/token";
        using var request = new HttpRequestMessage(HttpMethod.Post, url);
        var basicToken = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{config.ClientId}:{config.ClientSecret}"));
        request.Headers.Authorization = new AuthenticationHeaderValue("Basic", basicToken);
        request.Content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["grant_type"] = "authorization_code",
            ["code"] = authorizationCode,
            ["redirect_uri"] = config.RedirectUri
        });

        try
        {
            using var document = await SendJsonAsync(request, cancellationToken, treatUnauthorizedAsTokenExpired: false);
            config.AccessToken = GetString(document.RootElement, "access_token");
            var refreshToken = GetString(document.RootElement, "refresh_token");
            if (!string.IsNullOrWhiteSpace(refreshToken))
            {
                config.RefreshToken = refreshToken;
            }

            if (string.IsNullOrWhiteSpace(config.AccessToken))
            {
                throw new InvalidDataException("Cafe24 다시 인증 응답에서 ACCESS_TOKEN을 받지 못했습니다.");
            }
        }
        catch (HttpRequestException ex) when (IsAuthorizationCodeError(ex.Message))
        {
            throw new InvalidDataException("Cafe24 다시 인증 실패: callback URL 또는 code 값을 새로 받아 다시 입력해 주세요.", ex);
        }
        catch (HttpRequestException ex) when (IsCredentialError(ex.Message))
        {
            throw new InvalidDataException("Cafe24 다시 인증 실패: CLIENT_ID / CLIENT_SECRET 값을 확인해 주세요.", ex);
        }
    }

    public static string ExtractAuthorizationCode(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return string.Empty;
        }

        var code = ExtractQueryValue(input, "code");
        return string.IsNullOrWhiteSpace(code)
            ? input.Trim()
            : code.Trim();
    }

    public async Task<List<Cafe24Product>> GetProductsAsync(Cafe24TokenConfig config, bool onlySelling, CancellationToken cancellationToken)
    {
        var products = new List<Cafe24Product>();
        var offset = 0;
        const int limit = 100;

        while (true)
        {
            var url = $"https://{config.MallId}.cafe24api.com/api/v2/admin/products?shop_no={Uri.EscapeDataString(config.ShopNo)}&limit={limit}&offset={offset}";
            using var request = CreateRequest(HttpMethod.Get, url, config);
            using var document = await SendJsonAsync(request, cancellationToken);

            if (!document.RootElement.TryGetProperty("products", out var productsElement) || productsElement.ValueKind != JsonValueKind.Array)
            {
                break;
            }

            var pageCount = 0;
            foreach (var item in productsElement.EnumerateArray())
            {
                pageCount += 1;
                var selling = GetString(item, "selling");
                if (onlySelling && !string.Equals(selling, "T", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                products.Add(new Cafe24Product(
                    GetInt(item, "product_no"),
                    GetString(item, "product_name"),
                    GetString(item, "custom_product_code")));
            }

            if (pageCount < limit)
            {
                break;
            }

            offset += limit;
        }

        return products;
    }

    public Task<List<Cafe24Product>> GetSellingProductsAsync(Cafe24TokenConfig config, CancellationToken cancellationToken)
    {
        return GetProductsAsync(config, true, cancellationToken);
    }

    public async Task<int> CreateProductAsync(Cafe24TokenConfig config, IReadOnlyDictionary<string, object?> requestPayload, CancellationToken cancellationToken)
    {
        var url = $"https://{config.MallId}.cafe24api.com/api/v2/admin/products";
        var payload = new Dictionary<string, object?>
        {
            ["request"] = requestPayload
        };

        using var request = CreateJsonRequest(HttpMethod.Post, url, config, payload);
        using var document = await SendJsonAsync(request, cancellationToken);
        return ExtractProductNo(document.RootElement);
    }

    public async Task UploadMainImageAsync(Cafe24TokenConfig config, int productNo, string imagePath, CancellationToken cancellationToken)
    {
        var url = $"https://{config.MallId}.cafe24api.com/api/v2/admin/products/{productNo}/images";
        var payload = new
        {
            request = new
            {
                detail_image = Convert.ToBase64String(File.ReadAllBytes(imagePath)),
                image_upload_type = "A"
            }
        };

        using var request = CreateJsonRequest(HttpMethod.Post, url, config, payload);
        using var _ = await SendJsonAsync(request, cancellationToken);
    }

    public async Task UploadAdditionalImageAsync(Cafe24TokenConfig config, int productNo, string imagePath, CancellationToken cancellationToken)
    {
        var url = $"https://{config.MallId}.cafe24api.com/api/v2/admin/products/{productNo}/additionalimages";
        var payload = new
        {
            request = new
            {
                additional_image = new[] { Convert.ToBase64String(File.ReadAllBytes(imagePath)) }
            }
        };

        using var request = CreateJsonRequest(HttpMethod.Post, url, config, payload);
        using var _ = await SendJsonAsync(request, cancellationToken);
    }

    public async Task DeleteAdditionalImagesAsync(Cafe24TokenConfig config, int productNo, CancellationToken cancellationToken)
    {
        var url = $"https://{config.MallId}.cafe24api.com/api/v2/admin/products/{productNo}/additionalimages?shop_no={Uri.EscapeDataString(config.ShopNo)}";
        using var request = CreateRequest(HttpMethod.Delete, url, config);
        using var _ = await SendJsonAsync(request, cancellationToken);
    }

    public async Task<List<Cafe24Variant>> GetVariantsAsync(Cafe24TokenConfig config, int productNo, CancellationToken cancellationToken)
    {
        var url = $"https://{config.MallId}.cafe24api.com/api/v2/admin/products/{productNo}/variants?shop_no={Uri.EscapeDataString(config.ShopNo)}";
        using var request = CreateRequest(HttpMethod.Get, url, config);
        using var document = await SendJsonAsync(request, cancellationToken);

        var variants = new List<Cafe24Variant>();
        if (!document.RootElement.TryGetProperty("variants", out var variantsElement) || variantsElement.ValueKind != JsonValueKind.Array)
        {
            return variants;
        }

        foreach (var item in variantsElement.EnumerateArray())
        {
            var variantCode = GetString(item, "variant_code");
            if (string.IsNullOrWhiteSpace(variantCode))
            {
                continue;
            }

            var optionValues = new List<string>();
            if (item.TryGetProperty("options", out var optionsElement) && optionsElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var option in optionsElement.EnumerateArray())
                {
                    var value = GetString(option, "value");
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        optionValues.Add(value);
                    }
                }
            }

            variants.Add(new Cafe24Variant(variantCode, optionValues));
        }

        return variants;
    }

    public async Task UpdateVariantAsync(Cafe24TokenConfig config, int productNo, string variantCode, decimal additionalAmount, CancellationToken cancellationToken)
    {
        var url = $"https://{config.MallId}.cafe24api.com/api/v2/admin/products/{productNo}/variants/{Uri.EscapeDataString(variantCode)}";
        var payload = new
        {
            shop_no = int.TryParse(config.ShopNo, NumberStyles.Integer, CultureInfo.InvariantCulture, out var shopNo) ? shopNo : 1,
            request = new
            {
                additional_amount = additionalAmount.ToString("0.00", CultureInfo.InvariantCulture)
            }
        };

        using var request = CreateJsonRequest(HttpMethod.Put, url, config, payload);
        using var _ = await SendJsonAsync(request, cancellationToken);
    }

    public async Task RefreshAccessTokenAsync(Cafe24TokenConfig config, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(config.RefreshToken) || string.IsNullOrWhiteSpace(config.ClientId) || string.IsNullOrWhiteSpace(config.ClientSecret))
        {
            throw new InvalidDataException("Cafe24 토큰 자동 갱신에 필요한 설정이 없습니다.");
        }

        var url = $"https://{config.MallId}.cafe24api.com/api/v2/oauth/token";
        using var request = new HttpRequestMessage(HttpMethod.Post, url);
        var basicToken = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{config.ClientId}:{config.ClientSecret}"));
        request.Headers.Authorization = new AuthenticationHeaderValue("Basic", basicToken);

        var formData = new Dictionary<string, string>
        {
            ["grant_type"] = "refresh_token",
            ["refresh_token"] = config.RefreshToken
        };
        if (!string.IsNullOrWhiteSpace(config.RedirectUri))
        {
            formData["redirect_uri"] = config.RedirectUri;
        }
        request.Content = new FormUrlEncodedContent(formData);

        try
        {
            using var document = await SendJsonAsync(request, cancellationToken, treatUnauthorizedAsTokenExpired: false);
            config.AccessToken = GetString(document.RootElement, "access_token");
            var refreshToken = GetString(document.RootElement, "refresh_token");
            if (!string.IsNullOrWhiteSpace(refreshToken))
            {
                config.RefreshToken = refreshToken;
            }

            if (string.IsNullOrWhiteSpace(config.AccessToken))
            {
                throw new InvalidDataException("Cafe24 토큰 갱신 응답에서 ACCESS_TOKEN을 받지 못했습니다.");
            }
        }
        catch (HttpRequestException ex) when (RequiresReauthentication(ex.Message))
        {
            throw new Cafe24ReauthenticationRequiredException("Cafe24 토큰을 자동 갱신할 수 없습니다. 다시 인증해 주세요.", ex);
        }
    }

    private HttpRequestMessage CreateRequest(HttpMethod method, string url, Cafe24TokenConfig config)
    {
        var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", config.AccessToken);
        request.Headers.Add("X-Cafe24-Api-Version", config.ApiVersion);
        return request;
    }

    private HttpRequestMessage CreateJsonRequest(HttpMethod method, string url, Cafe24TokenConfig config, object payload)
    {
        var request = CreateRequest(method, url, config);
        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions { DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull });
        request.Content = new StringContent(json, Encoding.UTF8, "application/json");
        return request;
    }

    private async Task<JsonDocument> SendJsonAsync(HttpRequestMessage request, CancellationToken cancellationToken, bool treatUnauthorizedAsTokenExpired = true)
    {
        using var response = await _httpClient.SendAsync(request, cancellationToken);
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        if (treatUnauthorizedAsTokenExpired && response.StatusCode == HttpStatusCode.Unauthorized)
        {
            throw new Cafe24TokenExpiredException();
        }

        if (!response.IsSuccessStatusCode)
        {
            throw new HttpRequestException($"Cafe24 API 오류 ({(int)response.StatusCode} {response.ReasonPhrase}): {body}");
        }

        return string.IsNullOrWhiteSpace(body) ? JsonDocument.Parse("{}") : JsonDocument.Parse(body);
    }

    private static string BuildQueryString(IReadOnlyDictionary<string, string> values)
    {
        var builder = new StringBuilder();
        foreach (var pair in values)
        {
            if (builder.Length > 0)
            {
                builder.Append('&');
            }

            builder
                .Append(Uri.EscapeDataString(pair.Key))
                .Append('=')
                .Append(Uri.EscapeDataString(pair.Value ?? string.Empty));
        }

        return builder.ToString();
    }

    private static string ExtractQueryValue(string input, string key)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return string.Empty;
        }

        var source = input.Trim();
        if (Uri.TryCreate(source, UriKind.Absolute, out var absoluteUri))
        {
            source = absoluteUri.Query;
        }
        else
        {
            var questionIndex = source.IndexOf('?');
            if (questionIndex >= 0 && questionIndex + 1 < source.Length)
            {
                source = source[(questionIndex + 1)..];
            }
        }

        source = source.Trim().TrimStart('?', '#');
        if (string.IsNullOrWhiteSpace(source))
        {
            return string.Empty;
        }

        var pairs = source.Split('&', StringSplitOptions.RemoveEmptyEntries);
        foreach (var pair in pairs)
        {
            var separatorIndex = pair.IndexOf('=');
            var rawKey = separatorIndex >= 0 ? pair[..separatorIndex] : pair;
            var rawValue = separatorIndex >= 0 ? pair[(separatorIndex + 1)..] : string.Empty;
            var decodedKey = Uri.UnescapeDataString(rawKey.Replace("+", "%20"));
            if (!string.Equals(decodedKey, key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            return Uri.UnescapeDataString(rawValue.Replace("+", "%20"));
        }

        return string.Empty;
    }

    private static bool RequiresReauthentication(string message)
    {
        return IsCredentialError(message)
            || message.Contains("invalid_grant", StringComparison.OrdinalIgnoreCase)
            || message.Contains("invalid_client", StringComparison.OrdinalIgnoreCase)
            || message.Contains("invalid_request", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsCredentialError(string message)
    {
        return message.Contains("client_secret", StringComparison.OrdinalIgnoreCase)
            || message.Contains("client_id", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAuthorizationCodeError(string message)
    {
        return message.Contains("invalid_grant", StringComparison.OrdinalIgnoreCase)
            || message.Contains("authorization_code", StringComparison.OrdinalIgnoreCase)
            || message.Contains("expired", StringComparison.OrdinalIgnoreCase)
            || message.Contains("redirect_uri", StringComparison.OrdinalIgnoreCase)
            || message.Contains("invalid_request", StringComparison.OrdinalIgnoreCase);
    }

    private static int ExtractProductNo(JsonElement root)
    {
        if (root.TryGetProperty("product", out var productElement))
        {
            var productNo = GetInt(productElement, "product_no");
            if (productNo > 0)
            {
                return productNo;
            }
        }

        var rootProductNo = GetInt(root, "product_no");
        if (rootProductNo > 0)
        {
            return rootProductNo;
        }

        if (root.TryGetProperty("products", out var productsElement) && productsElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in productsElement.EnumerateArray())
            {
                var productNo = GetInt(item, "product_no");
                if (productNo > 0)
                {
                    return productNo;
                }
            }
        }

        return 0;
    }

    private static string GetString(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value))
        {
            return string.Empty;
        }

        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString() ?? string.Empty,
            JsonValueKind.Number => value.GetRawText(),
            JsonValueKind.True => bool.TrueString,
            JsonValueKind.False => bool.FalseString,
            _ => string.Empty
        };
    }

    private static int GetInt(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value))
        {
            return 0;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number))
        {
            return number;
        }

        if (value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
        {
            return parsed;
        }

        return 0;
    }
}
