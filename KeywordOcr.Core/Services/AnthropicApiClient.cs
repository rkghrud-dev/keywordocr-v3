using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace KeywordOcr.Core.Services;

/// <summary>
/// Anthropic Claude API 직접 HTTP 클라이언트 (Python anthropic_wrapper.py 포팅)
/// </summary>
public class AnthropicApiClient : IDisposable
{
    private readonly HttpClient _http;
    private readonly string _apiKey;

    private const string BaseUrl = "https://api.anthropic.com/v1";
    private const string AnthropicVersion = "2023-06-01";
    public const string DefaultModel = "claude-haiku-4-5-20251001";

    private static readonly JsonSerializerOptions _jsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        WriteIndented = false,
    };

    public AnthropicApiClient(string apiKey)
    {
        _apiKey = apiKey;
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(120) };
        _http.DefaultRequestHeaders.Add("x-api-key", apiKey);
        _http.DefaultRequestHeaders.Add("anthropic-version", AnthropicVersion);
        _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
    }

    // ── 텍스트 완성 ────────────────────────────────────────────────────────────

    public async Task<string?> CompleteAsync(
        string systemPrompt,
        string userPrompt,
        string? model = null,
        int maxTokens = 1024,
        double temperature = 0.1,
        CancellationToken ct = default)
    {
        var body = BuildTextBody(systemPrompt, userPrompt, model ?? DefaultModel, maxTokens, temperature);
        return await PostMessagesAsync(body, ct);
    }

    // ── Vision 분석 ────────────────────────────────────────────────────────────

    public async Task<string?> VisionAnalyzeAsync(
        string imagePath,
        string systemPrompt,
        string userPrompt,
        string? model = null,
        int maxTokens = 1500,
        CancellationToken ct = default)
    {
        var imageBytes = await File.ReadAllBytesAsync(imagePath, ct);
        var base64 = Convert.ToBase64String(imageBytes);
        var mediaType = GetMediaType(imagePath);

        var contentArr = new JsonArray
        {
            new JsonObject
            {
                ["type"] = "image",
                ["source"] = new JsonObject
                {
                    ["type"] = "base64",
                    ["media_type"] = mediaType,
                    ["data"] = base64,
                }
            },
            new JsonObject
            {
                ["type"] = "text",
                ["text"] = userPrompt,
            }
        };

        var body = new JsonObject
        {
            ["model"] = model ?? DefaultModel,
            ["max_tokens"] = maxTokens,
            ["temperature"] = 0.1,
            ["system"] = systemPrompt,
            ["messages"] = new JsonArray
            {
                new JsonObject { ["role"] = "user", ["content"] = contentArr }
            }
        };

        return await PostMessagesAsync(body.ToJsonString(), ct);
    }

    // ── Vision 분석 (다중 이미지) ──────────────────────────────────────────────

    public async Task<string?> VisionAnalyzeMultipleAsync(
        IReadOnlyList<string> imagePaths,
        string systemPrompt,
        string userPrompt,
        string? model = null,
        int maxTokens = 1800,
        CancellationToken ct = default)
    {
        var contentArr = new JsonArray();

        foreach (var path in imagePaths.Take(5))   // API 이미지 제한 고려
        {
            if (!File.Exists(path)) continue;
            var bytes = await File.ReadAllBytesAsync(path, ct);
            var b64 = Convert.ToBase64String(bytes);
            var mt = GetMediaType(path);
            contentArr.Add(new JsonObject
            {
                ["type"] = "image",
                ["source"] = new JsonObject
                {
                    ["type"] = "base64",
                    ["media_type"] = mt,
                    ["data"] = b64,
                }
            });
        }

        contentArr.Add(new JsonObject { ["type"] = "text", ["text"] = userPrompt });

        var body = new JsonObject
        {
            ["model"] = model ?? DefaultModel,
            ["max_tokens"] = maxTokens,
            ["temperature"] = 0.1,
            ["system"] = systemPrompt,
            ["messages"] = new JsonArray
            {
                new JsonObject { ["role"] = "user", ["content"] = contentArr }
            }
        };

        return await PostMessagesAsync(body.ToJsonString(), ct);
    }

    // ── JSON 완성 (response_format 지원 없으므로 프롬프트로 유도) ──────────────

    public async Task<Dictionary<string, object?>?> CompleteJsonAsync(
        string systemPrompt,
        string userPrompt,
        string? model = null,
        int maxTokens = 900,
        CancellationToken ct = default)
    {
        var raw = await CompleteAsync(systemPrompt, userPrompt, model, maxTokens, ct: ct);
        if (string.IsNullOrWhiteSpace(raw)) return null;

        // JSON 블록 추출 (```json ... ``` 또는 { ... })
        var json = ExtractJsonBlock(raw);
        if (string.IsNullOrWhiteSpace(json)) return null;

        try
        {
            return JsonSerializer.Deserialize<Dictionary<string, object?>>(json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }
        catch
        {
            return null;
        }
    }

    // ── 내부 ───────────────────────────────────────────────────────────────────

    private static string BuildTextBody(
        string system, string user, string model, int maxTokens, double temperature)
    {
        return JsonSerializer.Serialize(new
        {
            model,
            max_tokens = maxTokens,
            temperature,
            system,
            messages = new[] { new { role = "user", content = user } }
        });
    }

    private async Task<string?> PostMessagesAsync(string bodyJson, CancellationToken ct)
    {
        using var req = new HttpRequestMessage(HttpMethod.Post, $"{BaseUrl}/messages");
        req.Content = new StringContent(bodyJson, Encoding.UTF8, "application/json");

        using var resp = await _http.SendAsync(req, ct);
        var respText = await resp.Content.ReadAsStringAsync(ct);

        if (!resp.IsSuccessStatusCode)
            throw new HttpRequestException($"Anthropic API error {(int)resp.StatusCode}: {respText}");

        using var doc = JsonDocument.Parse(respText);
        var root = doc.RootElement;
        if (root.TryGetProperty("content", out var arr) && arr.GetArrayLength() > 0)
            return arr[0].GetProperty("text").GetString();
        return null;
    }

    private static string GetMediaType(string path)
    {
        var ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        return ext switch
        {
            "jpg" or "jpeg" => "image/jpeg",
            "png"           => "image/png",
            "gif"           => "image/gif",
            "webp"          => "image/webp",
            _               => "image/jpeg",
        };
    }

    private static string ExtractJsonBlock(string text)
    {
        // ```json ... ``` 블록
        var m = System.Text.RegularExpressions.Regex.Match(
            text, @"```json\s*([\s\S]*?)```", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (m.Success) return m.Groups[1].Value.Trim();

        // ``` ... ``` 블록
        m = System.Text.RegularExpressions.Regex.Match(text, @"```([\s\S]*?)```");
        if (m.Success) return m.Groups[1].Value.Trim();

        // 첫 번째 { } 블록
        var start = text.IndexOf('{');
        var end = text.LastIndexOf('}');
        if (start >= 0 && end > start) return text[start..(end + 1)];

        return text.Trim();
    }

    // ── API 키 파일 로더 ───────────────────────────────────────────────────────

    public static string? LoadApiKey()
    {
        var path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Desktop",
            "key",
            "anthropic_api_key.txt");

        return File.Exists(path) ? File.ReadAllText(path).Trim() : null;
    }

    public void Dispose() => _http.Dispose();
}
