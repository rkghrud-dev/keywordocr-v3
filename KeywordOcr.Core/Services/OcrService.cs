using System.Text;
using System.Text.RegularExpressions;
using Tesseract;

namespace KeywordOcr.Core.Services;

/// <summary>
/// Tesseract OCR 래퍼 (Python pytesseract 포팅)
/// tessdata 폴더에 kor.traineddata, eng.traineddata 필요.
/// </summary>
public class OcrService : IDisposable
{
    private TesseractEngine? _engine;
    private readonly string _tessDataPath;
    private readonly string _lang;
    private readonly int _psm;
    private readonly int _oem;
    private readonly bool _koreanOnly;
    private readonly bool _dropDigits;

    private static readonly Regex _hangulRe = new(@"[가-힣]+", RegexOptions.Compiled);
    private static readonly Regex _noiseRe  = new(@"[^0-9A-Za-z가-힣\s\-\+]", RegexOptions.Compiled);
    private static readonly Regex _wsRe     = new(@"\s+", RegexOptions.Compiled);

    public OcrService(
        string tessDataPath = "",
        bool koreanOnly = true,
        bool dropDigits = true,
        int psm = 11,
        int oem = 3)
    {
        _tessDataPath = string.IsNullOrEmpty(tessDataPath)
            ? FindTessData()
            : tessDataPath;
        _lang      = "kor+eng";
        _psm       = psm;
        _oem       = oem;
        _koreanOnly = koreanOnly;
        _dropDigits = dropDigits;
    }

    // ── 단일 이미지 OCR ───────────────────────────────────────────────────────

    public string ExtractText(string imagePath)
    {
        if (!File.Exists(imagePath)) return "";
        try
        {
            EnsureEngine();
            using var pix = Pix.LoadFromFile(imagePath);
            using var page = _engine!.Process(pix,
                (PageSegMode)_psm);
            var raw = page.GetText() ?? "";
            return CleanOcr(raw);
        }
        catch
        {
            return "";
        }
    }

    // ── 폴더 일괄 OCR ─────────────────────────────────────────────────────────

    public string ExtractTextFromFolder(
        string folderPath,
        int maxImgs = 999,
        int maxDepth = -1,
        CancellationToken ct = default)
    {
        if (!Directory.Exists(folderPath)) return "";

        var images = CollectImages(folderPath, maxDepth)
            .Take(maxImgs)
            .ToList();

        if (images.Count == 0) return "";

        var sb = new StringBuilder();
        foreach (var img in images)
        {
            ct.ThrowIfCancellationRequested();
            var text = ExtractText(img);
            if (!string.IsNullOrWhiteSpace(text))
                sb.AppendLine(text);
        }
        return sb.ToString().Trim();
    }

    // ── 텍스트 후처리 ─────────────────────────────────────────────────────────

    private string CleanOcr(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";

        // 한국어 포함 여부 필터
        if (_koreanOnly && !_hangulRe.IsMatch(raw)) return "";

        var txt = _noiseRe.Replace(raw, " ");

        if (_dropDigits)
            txt = Regex.Replace(txt, @"\b\d+\b", " ");

        txt = _wsRe.Replace(txt, " ").Trim();

        // 너무 짧으면 노이즈로 간주
        if (txt.Length < 4) return "";

        return txt;
    }

    // ── 이미지 수집 ───────────────────────────────────────────────────────────

    private static IEnumerable<string> CollectImages(string root, int maxDepth)
    {
        static IEnumerable<string> Walk(string dir, int depth, int maxD)
        {
            var files = Directory.EnumerateFiles(dir)
                .Where(f => IsImage(f))
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase);
            foreach (var f in files) yield return f;

            if (maxD >= 0 && depth >= maxD) yield break;

            foreach (var sub in Directory.EnumerateDirectories(dir).OrderBy(d => d))
                foreach (var f in Walk(sub, depth + 1, maxD))
                    yield return f;
        }

        return Walk(root, 0, maxDepth);
    }

    private static bool IsImage(string path)
    {
        var ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        return ext is "jpg" or "jpeg" or "png" or "bmp" or "tiff" or "tif" or "webp";
    }

    // ── tessdata 자동 탐색 ────────────────────────────────────────────────────

    private static string FindTessData()
    {
        var candidates = new[]
        {
            @"C:\Program Files\Tesseract-OCR\tessdata",
            @"C:\Program Files (x86)\Tesseract-OCR\tessdata",
            Path.Combine(AppContext.BaseDirectory, "tessdata"),
            Path.Combine(AppContext.BaseDirectory, "..", "tessdata"),
            Path.Combine(AppContext.BaseDirectory, "..", "..", "tessdata"),
        };
        foreach (var c in candidates)
            if (Directory.Exists(c)) return c;
        return "tessdata";
    }

    private void EnsureEngine()
    {
        if (_engine != null) return;
        _engine = new TesseractEngine(_tessDataPath, _lang,
            (EngineMode)_oem);
    }

    public void Dispose()
    {
        _engine?.Dispose();
        _engine = null;
    }
}
