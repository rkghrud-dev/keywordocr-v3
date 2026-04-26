using System.Text.Json;
using System.Text.RegularExpressions;
using KeywordOcr.Core.Models;

namespace KeywordOcr.Core.Services;

/// <summary>
/// 전체 파이프라인 오케스트레이터 (Python pipeline.py run() 포팅)
/// Python 프로세스 없이 C#만으로 동작.
/// </summary>
public class PipelineService : IDisposable
{
    private readonly PipelineConfig _cfg;
    private readonly AnthropicApiClient? _claude;
    private readonly OcrService? _ocr;
    private readonly MarketKeywordsService _marketKw;
    private readonly ListingImageService? _listing;
    private readonly ExcelIoService _excel;

    // ── Vision 분석 시스템 프롬프트 ──────────────────────────────────────────

    private const string VisionSystemPrompt =
        "당신은 한국 이커머스 상품 이미지 분석 전문가입니다. " +
        "이미지에서 상품의 특성을 JSON 형식으로 정확하게 분석해주세요. " +
        "각 항목은 한국어 명사/명사구로만 답하세요.";

    private const string VisionUserPrompt =
        """
다음 JSON 스키마로만 반환하라. JSON 이외의 텍스트는 절대 포함하지 마라:
{
  "core_identity": {
    "category": string[],
    "product_type_correction": string[],
    "structure": string[],
    "material_visual": string[],
    "color": string[]
  },
  "installation_and_physical": {
    "mount_type": string[],
    "installation_method": string[]
  },
  "usage_context": {
    "usage_location": string[],
    "usage_purpose": string[],
    "target_user": string[],
    "usage_scenario": string[],
    "indoor_outdoor": string[]
  },
  "functional_inference": {
    "primary_function": string[],
    "problem_solving_keyword": string[],
    "convenience_feature": string[]
  },
  "search_boost_elements": {
    "installation_keywords": string[],
    "space_keywords": string[],
    "benefit_keywords": string[],
    "longtail_candidates": string[]
  }
}

규칙:
- 이미지에서 보이는 것만 추출. 추측하지 말 것.
- 각 항목 최대 4개, 2~6자 명사구
- 광고문구/감성어/수식어 금지
""";

    // ── 생성자 ────────────────────────────────────────────────────────────────

    public PipelineService(PipelineConfig cfg)
    {
        _cfg = cfg;

        // API 키 로드
        var apiKey = !string.IsNullOrEmpty(cfg.AnthropicApiKey)
            ? cfg.AnthropicApiKey
            : AnthropicApiClient.LoadApiKey();

        if (!string.IsNullOrEmpty(apiKey))
            _claude = new AnthropicApiClient(apiKey);

        // OCR
        if (cfg.UseLocalOcr)
            _ocr = new OcrService(cfg.TesseractPath, cfg.KoreanOnly, cfg.DropDigits, cfg.Psm, cfg.Oem);

        // 마켓 키워드
        _marketKw = new MarketKeywordsService(_claude);

        // 대표이미지
        if (cfg.MakeListing)
            _listing = new ListingImageService(new ListingImageService.Config(
                Size:         cfg.ListingSize,
                Pad:          cfg.ListingPad,
                MaxImages:    cfg.ListingMax,
                LogoPath:     cfg.LogoPath,
                LogoRatio:    cfg.LogoRatio,
                LogoOpacity:  cfg.LogoOpacity,
                LogoPos:      cfg.LogoPos,
                AutoContrast: cfg.UseAutoContrast,
                Sharpen:      cfg.UseSharpen,
                SmallRotate:  cfg.UseSmallRotate,
                RotateZoom:   cfg.RotateZoom,
                AngleDeg:     cfg.UltraAngleDeg,
                FlipLr:       cfg.DoFlipLr,
                TrimTol:      cfg.TrimTol,
                JpegQMin:     cfg.JpegQMin,
                JpegQMax:     cfg.JpegQMax));

        _excel = new ExcelIoService();
    }

    // ── 진입점 ────────────────────────────────────────────────────────────────

    /// <summary>전체 파이프라인 실행. 결과 Excel 경로를 반환한다.</summary>
    public async Task<string> RunAsync(
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        progress?.Report("파일 로드 중...");
        var products = _excel.LoadInputFile(_cfg.FilePath);
        if (products.Count == 0)
            throw new InvalidOperationException("처리할 상품이 없습니다.");

        progress?.Report($"총 {products.Count}개 상품 처리 시작");

        var exportRoot = BuildExportRoot();
        Directory.CreateDirectory(exportRoot);

        var results = new List<PipelineResult>();

        // 병렬 처리 (ThreadPool 기반)
        var semaphore = new SemaphoreSlim(_cfg.Threads, _cfg.Threads);
        var tasks = products.Select(async (product, idx) =>
        {
            await semaphore.WaitAsync(ct);
            try
            {
                ct.ThrowIfCancellationRequested();
                progress?.Report($"[{idx + 1}/{products.Count}] {product.GsCode} {product.ProductName[..Math.Min(20, product.ProductName.Length)]}...");
                var result = await ProcessProductAsync(product, exportRoot, ct);
                return result;
            }
            finally
            {
                semaphore.Release();
            }
        });

        foreach (var task in await Task.WhenAll(tasks))
            results.Add(task);

        // 결과 Excel 저장
        var outputPath = Path.Combine(exportRoot, BuildOutputFileName());
        _excel.WriteOutputFile(outputPath, results,
            products.FirstOrDefault()?.RawColumns.Keys.ToList());

        progress?.Report($"완료 → {outputPath}");
        return outputPath;
    }

    /// <summary>키워드만 생성 (대표이미지 제외)</summary>
    public async Task<string> RunKeywordsOnlyAsync(
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        var savedPhase = _cfg.Phase;
        _cfg.Phase = "analysis";
        try { return await RunAsync(progress, ct); }
        finally { _cfg.Phase = savedPhase; }
    }

    // ── 상품 단위 처리 ────────────────────────────────────────────────────────

    private async Task<PipelineResult> ProcessProductAsync(
        ProductRow product,
        string exportRoot,
        CancellationToken ct)
    {
        var result = new PipelineResult
        {
            RowIndex    = product.RowIndex,
            GsCode      = product.GsCode,
            ProductName = product.ProductName,
            RawColumns  = product.RawColumns,
        };

        try
        {
            // 1. 이미지 폴더 탐색
            var imgFolder = FindImageFolder(product.GsCode, exportRoot);

            // 2. OCR
            string ocrText = "";
            if (_ocr != null && !string.IsNullOrEmpty(imgFolder))
                ocrText = _ocr.ExtractTextFromFolder(imgFolder, _cfg.MaxImgs, _cfg.MaxDepth, ct);

            if (_cfg.MergeOcrWithName)
                ocrText = string.Join(" ", product.ProductName, ocrText).Trim();

            result.OcrText = ocrText;

            // 3. Vision 분석 (대표 이미지 최대 3장)
            Dictionary<string, object?>? visionAnalysis = null;
            if (_claude != null && !string.IsNullOrEmpty(imgFolder))
            {
                var imgs = CollectTopImages(imgFolder, 3);
                if (imgs.Count > 0)
                    visionAnalysis = await AnalyzeVisionAsync(imgs, product.ProductName, ct);
            }

            // 4. A마켓 키워드
            var pkgA = await _marketKw.GenerateAsync(
                product.ProductName, ocrText, _cfg.ModelKeyword,
                market: "A", ct: ct);

            // keyword_builder로 검색어설정 조립
            var kwStringA = KeywordBuilderService.BuildKeywordString(
                ocrText, visionAnalysis, _cfg.ATagCount,
                product.ProductName, "A");

            result.SearchKeywordsA = !string.IsNullOrEmpty(kwStringA)
                ? kwStringA : pkgA.SearchKeywords;
            result.CoupangTagsA    = pkgA.CoupangTags;
            result.NaverTagsA      = pkgA.NaverTags;
            result.CandidatePool   = pkgA.CandidatePool;

            // 5. B마켓 키워드 (활성화된 경우)
            if (_cfg.EnableBMarket)
            {
                var avoidFromA = pkgA.CoupangTags.Skip(6).ToList(); // A마켓 후반부 태그 회피
                var pkgB = await _marketKw.GenerateAsync(
                    product.ProductName, ocrText, _cfg.ModelKeyword,
                    market: "B", avoidTerms: avoidFromA, ct: ct);

                var kwStringB = KeywordBuilderService.BuildKeywordString(
                    ocrText, visionAnalysis, _cfg.BTagCount,
                    product.ProductName, "B");

                result.SearchKeywordsB = !string.IsNullOrEmpty(kwStringB)
                    ? kwStringB : pkgB.SearchKeywords;
                result.CoupangTagsB    = pkgB.CoupangTags;
                result.NaverTagsB      = pkgB.NaverTags;
            }

            // 6. 대표이미지 생성
            if (_listing != null && _cfg.Phase != "analysis"
                && !string.IsNullOrEmpty(imgFolder))
            {
                var listingOut = Path.Combine(exportRoot, "listing", product.GsCode);
                var listingPaths = _listing.ProcessFolder(imgFolder, listingOut, ct: ct);
                result.ListingImagePaths = listingPaths;
            }
        }
        catch (Exception ex)
        {
            result.Success      = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    // ── Vision 분석 ───────────────────────────────────────────────────────────

    private async Task<Dictionary<string, object?>?> AnalyzeVisionAsync(
        IReadOnlyList<string> imagePaths,
        string productName,
        CancellationToken ct)
    {
        if (_claude == null || imagePaths.Count == 0) return null;
        try
        {
            string? raw;
            if (imagePaths.Count == 1)
                raw = await _claude.VisionAnalyzeAsync(
                    imagePaths[0], VisionSystemPrompt,
                    $"상품명: {productName}\n\n{VisionUserPrompt}",
                    _cfg.ModelKeyword, 1500, ct);
            else
                raw = await _claude.VisionAnalyzeMultipleAsync(
                    imagePaths,
                    VisionSystemPrompt,
                    $"상품명: {productName}\n\n{VisionUserPrompt}",
                    _cfg.ModelKeyword, 1800, ct);

            if (string.IsNullOrWhiteSpace(raw)) return null;

            return JsonSerializer.Deserialize<Dictionary<string, object?>>(
                ExtractJsonBlock(raw),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }
        catch
        {
            return null;
        }
    }

    // ── 이미지 폴더 탐색 ─────────────────────────────────────────────────────

    private string? FindImageFolder(string gsCode, string exportRoot)
    {
        // 1. 설정된 localImgDir 탐색
        if (!string.IsNullOrEmpty(_cfg.LocalImgDir)
            && Directory.Exists(_cfg.LocalImgDir))
        {
            var direct = Path.Combine(_cfg.LocalImgDir, gsCode);
            if (Directory.Exists(direct)) return direct;

            // GS코드 앞 7자리로 폴더 매칭
            var prefix = Regex.Match(gsCode, @"GS\d{7}").Value;
            if (!string.IsNullOrEmpty(prefix) && _cfg.AllowFolderMatch)
            {
                var match = Directory.EnumerateDirectories(_cfg.LocalImgDir)
                    .FirstOrDefault(d => Path.GetFileName(d).StartsWith(prefix,
                        StringComparison.OrdinalIgnoreCase));
                if (match != null) return match;
            }
        }

        // 2. 입력 파일과 같은 폴더
        var inputDir = Path.GetDirectoryName(_cfg.FilePath) ?? "";
        foreach (var candidate in new[] { gsCode, Regex.Match(gsCode, @"GS\d{7}").Value })
        {
            if (string.IsNullOrEmpty(candidate)) continue;
            var dir = Path.Combine(inputDir, candidate);
            if (Directory.Exists(dir)) return dir;
        }

        return null;
    }

    private static List<string> CollectTopImages(string folder, int max)
    {
        return Directory.EnumerateFiles(folder)
            .Where(f =>
            {
                var ext = Path.GetExtension(f).TrimStart('.').ToLowerInvariant();
                return ext is "jpg" or "jpeg" or "png" or "webp";
            })
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .Take(max)
            .ToList();
    }

    // ── 경로 유틸 ─────────────────────────────────────────────────────────────

    private string BuildExportRoot()
    {
        if (!string.IsNullOrEmpty(_cfg.ExportRootOverride))
            return _cfg.ExportRootOverride;

        var inputDir  = Path.GetDirectoryName(_cfg.FilePath) ?? "";
        var dateStamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
        return Path.Combine(inputDir, $"output_{dateStamp}");
    }

    private string BuildOutputFileName()
    {
        var stem = Path.GetFileNameWithoutExtension(_cfg.FilePath);
        return $"{stem}_result_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
    }

    private static string ExtractJsonBlock(string text)
    {
        var m = Regex.Match(text, @"```json\s*([\s\S]*?)```", RegexOptions.IgnoreCase);
        if (m.Success) return m.Groups[1].Value.Trim();
        m = Regex.Match(text, @"```([\s\S]*?)```");
        if (m.Success) return m.Groups[1].Value.Trim();
        int start = text.IndexOf('{'), end = text.LastIndexOf('}');
        if (start >= 0 && end > start) return text[start..(end + 1)];
        return text.Trim();
    }

    public void Dispose()
    {
        _claude?.Dispose();
        _ocr?.Dispose();
    }
}
