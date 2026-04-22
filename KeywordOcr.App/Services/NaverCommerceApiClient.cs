using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace KeywordOcr.App.Services;

/// <summary>
/// 네이버 커머스 API 클라이언트 (순수 C#)
/// - bcrypt + base64 OAuth2 인증
/// - 카테고리 검색, 이미지 업로드, 상품 등록
/// </summary>
public sealed class NaverCommerceApiClient : IDisposable
{
    private const string ApiBase = "https://api.commerce.naver.com/external";

    private readonly string _clientId;
    private readonly string _clientSecret;
    private readonly HttpClient _http;
    private readonly long? _referenceOriginProductNo;
    private readonly long? _referenceChannelProductNo;

    private string? _cachedToken;
    private DateTime _tokenExpiry = DateTime.MinValue;

    public long? ReferenceOriginProductNo => _referenceOriginProductNo;
    public long? ReferenceChannelProductNo => _referenceChannelProductNo;

    public NaverCommerceApiClient(
        string clientId,
        string clientSecret,
        long? referenceOriginProductNo = null,
        long? referenceChannelProductNo = null)
    {
        _clientId = clientId;
        _clientSecret = clientSecret;
        _referenceOriginProductNo = referenceOriginProductNo;
        _referenceChannelProductNo = referenceChannelProductNo;
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
    }

    public static NaverCommerceApiClient FromKeyFile()
    {
        var keyFile = DesktopKeyStore.GetPath("naver_client_key.txt");

        var kv = new Dictionary<string, string>();
        if (File.Exists(keyFile))
        {
            foreach (var line in File.ReadAllLines(keyFile, Encoding.UTF8))
            {
                var trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith('#')) continue;
                var eqIdx = trimmed.IndexOf('=');
                if (eqIdx <= 0) continue;
                kv[trimmed[..eqIdx].Trim()] = trimmed[(eqIdx + 1)..].Trim();
            }
        }

        var clientId = kv.GetValueOrDefault("NAVER_COMMERCE_CLIENT_ID", "");
        var clientSecret = kv.GetValueOrDefault("NAVER_COMMERCE_CLIENT_SECRET", "");
        var referenceOriginProductNo = ParseOptionalLong(kv.GetValueOrDefault("NAVER_REFERENCE_ORIGIN_PRODUCT_NO", ""));
        var referenceChannelProductNo = ParseOptionalLong(kv.GetValueOrDefault("NAVER_REFERENCE_CHANNEL_PRODUCT_NO", ""));

        if (string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
            throw new InvalidOperationException($"네이버 커머스 API 키를 찾을 수 없습니다: {keyFile}");

        return new NaverCommerceApiClient(clientId, clientSecret, referenceOriginProductNo, referenceChannelProductNo);
    }

    private static long? ParseOptionalLong(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return null;

        return long.TryParse(raw.Trim(), out var value) ? value : null;
    }

    // ── 인증 ───────────────────────────────────────

    private async Task<string> GetAccessTokenAsync(CancellationToken ct)
    {
        if (_cachedToken is not null && DateTime.UtcNow < _tokenExpiry)
            return _cachedToken;

        var timestamp = (long)((DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()) - 3000);
        var password = $"{_clientId}_{timestamp}";
        var hashed = BCrypt.Net.BCrypt.HashPassword(password, _clientSecret);
        var sign = Convert.ToBase64String(Encoding.UTF8.GetBytes(hashed));

        var form = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["client_id"] = _clientId,
            ["timestamp"] = timestamp.ToString(),
            ["client_secret_sign"] = sign,
            ["grant_type"] = "client_credentials",
            ["type"] = "SELF",
        });

        using var resp = await _http.PostAsync($"{ApiBase}/v1/oauth2/token", form, ct);
        var json = await resp.Content.ReadAsStringAsync(ct);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        _cachedToken = root.GetProperty("access_token").GetString()!;
        var expiresIn = root.TryGetProperty("expires_in", out var exp) ? exp.GetInt32() : 600;
        _tokenExpiry = DateTime.UtcNow.AddSeconds(expiresIn - 60);

        return _cachedToken;
    }

    // ── API 호출 ───────────────────────────────────

    private async Task<JsonDocument> CallAsync(
        string method, string path,
        HttpContent? content = null,
        string? query = null,
        CancellationToken ct = default)
    {
        var token = await GetAccessTokenAsync(ct);
        var url = $"{ApiBase}{path}";
        if (!string.IsNullOrEmpty(query)) url += "?" + query;

        using var req = new HttpRequestMessage(new HttpMethod(method), url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (content is not null) req.Content = content;

        using var resp = await _http.SendAsync(req, ct);
        var raw = await resp.Content.ReadAsStringAsync(ct);

        if (string.IsNullOrEmpty(raw))
            return JsonDocument.Parse($$$"""{"_error":{{{(int)resp.StatusCode}}},"_msg":"빈 응답"}""");

        try
        {
            return JsonDocument.Parse(raw);
        }
        catch
        {
            if (raw.Length > 500) raw = raw[..500];
            raw = raw.Replace("\"", "'");
            return JsonDocument.Parse($$$"""{"_error":{{{(int)resp.StatusCode}}},"_msg":"{{{raw}}}"}""");
        }
    }

    // ── 공개 API ───────────────────────────────────

    /// <summary>상품모델 검색으로 카테고리 추천</summary>
    public async Task<JsonDocument> PredictCategoryAsync(string productName, CancellationToken ct = default)
    {
        var clean = System.Text.RegularExpressions.Regex.Replace(productName, @"[A-Z]{1,2}\d{5,}[A-Z]?", "").Trim();
        clean = System.Text.RegularExpressions.Regex.Replace(clean, @"\d+(\.\d+)?\s*(cm|mm|m|g|kg|ml|L|개|매|장|ea)", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();
        clean = System.Text.RegularExpressions.Regex.Replace(clean, @"\s+", " ").Trim();
        if (string.IsNullOrEmpty(clean)) clean = productName;

        var query = $"name={Uri.EscapeDataString(clean)}";
        return await CallAsync("GET", "/v1/product-models", query: query, ct: ct);
    }

    /// <summary>기준 원상품 조회</summary>
    public async Task<JsonDocument> GetOriginProductAsync(long originProductNo, CancellationToken ct = default)
    {
        return await CallAsync("GET", $"/v2/products/origin-products/{originProductNo}", ct: ct);
    }

    /// <summary>기준 채널상품 조회</summary>
    public async Task<JsonDocument> GetChannelProductAsync(long channelProductNo, CancellationToken ct = default)
    {
        return await CallAsync("GET", $"/v2/products/channel-products/{channelProductNo}", ct: ct);
    }

    /// <summary>이미지 업로드 (로컬 파일 또는 URL)</summary>
    public async Task<string> UploadImageAsync(string imageSource, CancellationToken ct = default)
    {
        byte[] imageData;
        string fileName;

        if (File.Exists(imageSource))
        {
            imageData = await File.ReadAllBytesAsync(imageSource, ct);
            fileName = Path.GetFileName(imageSource);
        }
        else
        {
            using var imgResp = await _http.GetAsync(imageSource, ct);
            imageData = await imgResp.Content.ReadAsByteArrayAsync(ct);
            var uri = new Uri(imageSource);
            fileName = Path.GetFileName(uri.LocalPath);
            if (string.IsNullOrEmpty(fileName)) fileName = "image.jpg";
        }

        var mime = "image/jpeg";
        if (imageData.Length >= 8)
        {
            if (imageData[0] == 0x89 && imageData[1] == 0x50) mime = "image/png";
            else if (imageData[0] == 0x47 && imageData[1] == 0x49) mime = "image/gif";
            else if (imageData[0] == 0x42 && imageData[1] == 0x4D) mime = "image/bmp";
        }

        var ext = mime switch
        {
            "image/png" => ".png",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            _ => ".jpg",
        };
        var baseName = Path.GetFileNameWithoutExtension(fileName);
        fileName = baseName + ext;

        var token = await GetAccessTokenAsync(ct);

        using var formContent = new MultipartFormDataContent();
        var imageContent = new ByteArrayContent(imageData);
        imageContent.Headers.ContentType = new MediaTypeHeaderValue(mime);
        formContent.Add(imageContent, "imageFiles", fileName);

        using var req = new HttpRequestMessage(HttpMethod.Post, $"{ApiBase}/v1/product-images/upload");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Content = formContent;

        using var resp = await _http.SendAsync(req, ct);
        var json = await resp.Content.ReadAsStringAsync(ct);
        using var doc = JsonDocument.Parse(json);

        if (doc.RootElement.TryGetProperty("images", out var images) && images.GetArrayLength() > 0)
        {
            var url = images[0].GetProperty("url").GetString();
            if (!string.IsNullOrEmpty(url)) return url;
        }

        throw new InvalidOperationException($"네이버 이미지 업로드 실패: {json[..Math.Min(json.Length, 200)]}");
    }

    /// <summary>판매자 관리 코드(sellerManagementCode)가 GS로 시작하는 채널 상품 목록 조회</summary>
    public async Task<List<(string GsCode, string ProductName)>> GetExistingGsCodesAsync(CancellationToken ct = default)
    {
        var result = new List<(string, string)>();
        var page = 0;
        const int pageSize = 100;

        while (true)
        {
            var query = $"size={pageSize}&page={page}";
            using var doc = await CallAsync("GET", "/v2/products/channel-products", query: query, ct: ct);
            var root = doc.RootElement;

            JsonElement contents;
            if (!root.TryGetProperty("contents", out contents) || contents.ValueKind != JsonValueKind.Array)
                break;

            var count = 0;
            foreach (var item in contents.EnumerateArray())
            {
                count++;
                var code = item.TryGetProperty("sellerManagementCode", out var cProp) ? cProp.GetString() ?? "" : "";
                if (code.StartsWith("GS", StringComparison.OrdinalIgnoreCase))
                {
                    var name = item.TryGetProperty("name", out var nProp) ? nProp.GetString() ?? "" : "";
                    result.Add((code.ToUpperInvariant(), name));
                }
            }

            if (count < pageSize)
                break;
            page++;
        }

        return result;
    }

    /// <summary>상품 등록</summary>
    public async Task<JsonDocument> CreateProductAsync(JsonElement productJson, CancellationToken ct = default)
    {
        var body = new StringContent(
            JsonSerializer.Serialize(productJson),
            Encoding.UTF8, "application/json");
        return await CallAsync("POST", "/v2/products", content: body, ct: ct);
    }

    public void Dispose() => _http.Dispose();
}
