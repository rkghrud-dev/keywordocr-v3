using SkiaSharp;

namespace KeywordOcr.Core.Services;

/// <summary>
/// 대표이미지 합성 서비스 (Python process_listing_images_global 포팅)
/// SkiaSharp 기반 — GDI/Pillow 의존 없음.
/// </summary>
public class ListingImageService
{
    public record Config(
        int Size         = 1200,
        int Pad          = 20,
        int MaxImages    = 20,
        string LogoPath  = "",
        int LogoRatio    = 14,     // 이미지 너비 대비 로고 폭 비율 (1/n)
        int LogoOpacity  = 65,     // 0~100
        string LogoPos   = "tr",   // tl/tr/bl/br/center
        bool AutoContrast = true,
        bool Sharpen      = true,
        bool SmallRotate  = true,
        float RotateZoom  = 1.04f,
        float AngleDeg    = 0.35f,
        bool FlipLr       = true,
        int TrimTol       = 8,
        int JpegQMin      = 88,
        int JpegQMax      = 92
    );

    private readonly Config _cfg;
    private SKBitmap? _logoBitmap;

    public ListingImageService(Config? cfg = null)
    {
        _cfg = cfg ?? new Config();
        if (!string.IsNullOrEmpty(_cfg.LogoPath) && File.Exists(_cfg.LogoPath))
            _logoBitmap = LoadBitmap(_cfg.LogoPath);
    }

    // ── 폴더 일괄 처리 ────────────────────────────────────────────────────────

    public List<string> ProcessFolder(
        string inputFolder,
        string outputFolder,
        bool flipVariant = true,
        CancellationToken ct = default)
    {
        if (!Directory.Exists(inputFolder)) return [];
        Directory.CreateDirectory(outputFolder);

        var images = Directory.EnumerateFiles(inputFolder)
            .Where(IsImage)
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .Take(_cfg.MaxImages)
            .ToList();

        var results = new List<string>();
        foreach (var src in images)
        {
            ct.ThrowIfCancellationRequested();

            // 정방향
            var dst = Path.Combine(outputFolder,
                Path.GetFileNameWithoutExtension(src) + ".jpg");
            if (ProcessSingle(src, dst, flipLr: false))
                results.Add(dst);

            // 좌우반전 (FlipLr 옵션)
            if (flipVariant && _cfg.FlipLr)
            {
                var dstFlip = Path.Combine(outputFolder,
                    Path.GetFileNameWithoutExtension(src) + "_flip.jpg");
                if (ProcessSingle(src, dstFlip, flipLr: true))
                    results.Add(dstFlip);
            }
        }
        return results;
    }

    // ── 단일 이미지 처리 ──────────────────────────────────────────────────────

    public bool ProcessSingle(string srcPath, string dstPath, bool flipLr = false)
    {
        try
        {
            using var bmp = LoadBitmap(srcPath);
            if (bmp == null) return false;

            using var processed = ApplyEffects(bmp, flipLr);
            var quality = Random.Shared.Next(_cfg.JpegQMin, _cfg.JpegQMax + 1);
            SaveJpeg(processed, dstPath, quality);
            return true;
        }
        catch
        {
            return false;
        }
    }

    // ── 효과 파이프라인 ───────────────────────────────────────────────────────

    private SKBitmap ApplyEffects(SKBitmap src, bool flipLr)
    {
        var bmp = ResizeAndPad(src, _cfg.Size, _cfg.Pad, _cfg.TrimTol);

        if (_cfg.AutoContrast)
            bmp = ApplyAutoContrast(bmp);

        if (_cfg.Sharpen)
            bmp = ApplySharpen(bmp);

        if (_cfg.SmallRotate)
            bmp = ApplySmallRotate(bmp, _cfg.AngleDeg, _cfg.RotateZoom);

        if (flipLr)
            bmp = ApplyFlipLr(bmp);

        if (_logoBitmap != null)
            bmp = OverlayLogo(bmp, _logoBitmap, _cfg.LogoRatio, _cfg.LogoOpacity, _cfg.LogoPos);

        return bmp;
    }

    // ── 리사이즈 + 패딩 (흰 배경) ───────────────────────────────────────────

    private static SKBitmap ResizeAndPad(SKBitmap src, int size, int pad, int trimTol)
    {
        // 여백 트리밍
        var trimmed = TrimWhitespace(src, trimTol);

        // 비율 유지 리사이즈
        int contentSize = size - pad * 2;
        float scale = Math.Min((float)contentSize / trimmed.Width,
                               (float)contentSize / trimmed.Height);
        int w = (int)(trimmed.Width  * scale);
        int h = (int)(trimmed.Height * scale);

        var result = new SKBitmap(size, size);
        using var canvas = new SKCanvas(result);
        canvas.Clear(SKColors.White);

        int x = (size - w) / 2;
        int y = (size - h) / 2;

        using var scaled = trimmed.Resize(new SKImageInfo(w, h), SKFilterQuality.High);
        canvas.DrawBitmap(scaled, x, y);

        trimmed.Dispose();
        return result;
    }

    private static SKBitmap TrimWhitespace(SKBitmap src, int tol)
    {
        int left = 0, top = 0, right = src.Width - 1, bottom = src.Height - 1;

        bool IsBackground(SKColor c)
            => c.Red > 255 - tol && c.Green > 255 - tol && c.Blue > 255 - tol;

        // top
        for (int row = 0; row < src.Height; row++)
        {
            bool allBg = true;
            for (int col = 0; col < src.Width; col++)
                if (!IsBackground(src.GetPixel(col, row))) { allBg = false; break; }
            if (!allBg) { top = row; break; }
        }
        // bottom
        for (int row = src.Height - 1; row >= top; row--)
        {
            bool allBg = true;
            for (int col = 0; col < src.Width; col++)
                if (!IsBackground(src.GetPixel(col, row))) { allBg = false; break; }
            if (!allBg) { bottom = row; break; }
        }
        // left
        for (int col = 0; col < src.Width; col++)
        {
            bool allBg = true;
            for (int row = top; row <= bottom; row++)
                if (!IsBackground(src.GetPixel(col, row))) { allBg = false; break; }
            if (!allBg) { left = col; break; }
        }
        // right
        for (int col = src.Width - 1; col >= left; col--)
        {
            bool allBg = true;
            for (int row = top; row <= bottom; row++)
                if (!IsBackground(src.GetPixel(col, row))) { allBg = false; break; }
            if (!allBg) { right = col; break; }
        }

        int cw = right - left + 1, ch = bottom - top + 1;
        if (cw <= 0 || ch <= 0) return src.Copy();

        var cropped = new SKBitmap(cw, ch);
        using var canvas = new SKCanvas(cropped);
        canvas.DrawBitmap(src, new SKRect(left, top, right + 1, bottom + 1),
            new SKRect(0, 0, cw, ch));
        return cropped;
    }

    // ── 자동 대비 (히스토그램 스트레치) ─────────────────────────────────────

    private static SKBitmap ApplyAutoContrast(SKBitmap src)
    {
        // 채널별 min/max 수집
        byte minR = 255, maxR = 0, minG = 255, maxG = 0, minB = 255, maxB = 0;
        for (int y = 0; y < src.Height; y++)
            for (int x = 0; x < src.Width; x++)
            {
                var c = src.GetPixel(x, y);
                if (c.Red   < minR) minR = c.Red;   if (c.Red   > maxR) maxR = c.Red;
                if (c.Green < minG) minG = c.Green;  if (c.Green > maxG) maxG = c.Green;
                if (c.Blue  < minB) minB = c.Blue;   if (c.Blue  > maxB) maxB = c.Blue;
            }

        if (maxR == minR) maxR = (byte)(minR + 1);
        if (maxG == minG) maxG = (byte)(minG + 1);
        if (maxB == minB) maxB = (byte)(minB + 1);

        var result = new SKBitmap(src.Width, src.Height);
        for (int y = 0; y < src.Height; y++)
            for (int x = 0; x < src.Width; x++)
            {
                var c = src.GetPixel(x, y);
                byte r = Scale(c.Red,   minR, maxR);
                byte g = Scale(c.Green, minG, maxG);
                byte b = Scale(c.Blue,  minB, maxB);
                result.SetPixel(x, y, new SKColor(r, g, b, c.Alpha));
            }
        return result;

        static byte Scale(byte v, byte min, byte max)
            => (byte)Math.Clamp((v - min) * 255 / (max - min), 0, 255);
    }

    // ── 샤프닝 (언샤프 마스크 근사) ─────────────────────────────────────────

    private static SKBitmap ApplySharpen(SKBitmap src)
    {
        // 3×3 샤프닝 커널
        float[] kernel = [0f, -1f, 0f, -1f, 5f, -1f, 0f, -1f, 0f];
        var result = new SKBitmap(src.Width, src.Height);

        for (int y = 1; y < src.Height - 1; y++)
            for (int x = 1; x < src.Width - 1; x++)
            {
                float r = 0, g = 0, b = 0;
                for (int ky = -1; ky <= 1; ky++)
                    for (int kx = -1; kx <= 1; kx++)
                    {
                        var c = src.GetPixel(x + kx, y + ky);
                        var k = kernel[(ky + 1) * 3 + (kx + 1)];
                        r += c.Red   * k;
                        g += c.Green * k;
                        b += c.Blue  * k;
                    }
                var orig = src.GetPixel(x, y);
                result.SetPixel(x, y, new SKColor(
                    Clamp(r), Clamp(g), Clamp(b), orig.Alpha));
            }

        // 경계는 원본 복사
        for (int y = 0; y < src.Height; y++)
        {
            result.SetPixel(0, y, src.GetPixel(0, y));
            result.SetPixel(src.Width - 1, y, src.GetPixel(src.Width - 1, y));
        }
        for (int x = 0; x < src.Width; x++)
        {
            result.SetPixel(x, 0, src.GetPixel(x, 0));
            result.SetPixel(x, src.Height - 1, src.GetPixel(x, src.Height - 1));
        }
        return result;

        static byte Clamp(float v) => (byte)Math.Clamp((int)v, 0, 255);
    }

    // ── 미세 회전 ─────────────────────────────────────────────────────────────

    private static SKBitmap ApplySmallRotate(SKBitmap src, float angleDeg, float zoom)
    {
        var result = new SKBitmap(src.Width, src.Height);
        using var canvas = new SKCanvas(result);
        canvas.Clear(SKColors.White);

        var cx = src.Width  / 2f;
        var cy = src.Height / 2f;

        var m = SKMatrix.CreateIdentity();
        m = SKMatrix.Concat(m, SKMatrix.CreateTranslation(-cx, -cy));
        m = SKMatrix.Concat(m, SKMatrix.CreateRotationDegrees(angleDeg));
        m = SKMatrix.Concat(m, SKMatrix.CreateScale(zoom, zoom));
        m = SKMatrix.Concat(m, SKMatrix.CreateTranslation(cx, cy));

        canvas.SetMatrix(m);
        canvas.DrawBitmap(src, 0, 0);
        return result;
    }

    // ── 좌우반전 ─────────────────────────────────────────────────────────────

    private static SKBitmap ApplyFlipLr(SKBitmap src)
    {
        var result = new SKBitmap(src.Width, src.Height);
        using var canvas = new SKCanvas(result);
        var m = SKMatrix.CreateScale(-1, 1, src.Width / 2f, 0);
        canvas.SetMatrix(m);
        canvas.DrawBitmap(src, 0, 0);
        return result;
    }

    // ── 로고 오버레이 ─────────────────────────────────────────────────────────

    private static SKBitmap OverlayLogo(
        SKBitmap bg, SKBitmap logo, int ratio, int opacity, string pos)
    {
        int logoW = bg.Width / ratio;
        float scale = (float)logoW / logo.Width;
        int logoH = (int)(logo.Height * scale);

        const int margin = 10;
        var (lx, ly) = pos.ToLowerInvariant() switch
        {
            "tl" => (margin, margin),
            "tr" => (bg.Width - logoW - margin, margin),
            "bl" => (margin, bg.Height - logoH - margin),
            "br" => (bg.Width - logoW - margin, bg.Height - logoH - margin),
            _    => ((bg.Width - logoW) / 2, (bg.Height - logoH) / 2),
        };

        var result = bg.Copy();
        using var canvas = new SKCanvas(result);

        using var paint = new SKPaint
        {
            ColorFilter = SKColorFilter.CreateBlendMode(
                SKColors.White.WithAlpha((byte)(opacity * 255 / 100)),
                SKBlendMode.DstIn)
        };

        using var scaledLogo = logo.Resize(new SKImageInfo(logoW, logoH), SKFilterQuality.High);
        canvas.DrawBitmap(scaledLogo, lx, ly, paint);
        return result;
    }

    // ── 저장 ─────────────────────────────────────────────────────────────────

    private static void SaveJpeg(SKBitmap bmp, string path, int quality)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var image = SKImage.FromBitmap(bmp);
        using var data  = image.Encode(SKEncodedImageFormat.Jpeg, quality);
        using var fs    = File.OpenWrite(path);
        data.SaveTo(fs);
    }

    // ── 유틸 ─────────────────────────────────────────────────────────────────

    private static SKBitmap? LoadBitmap(string path)
    {
        try { return SKBitmap.Decode(path); }
        catch { return null; }
    }

    private static bool IsImage(string path)
    {
        var ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        return ext is "jpg" or "jpeg" or "png" or "bmp" or "tiff" or "tif" or "webp";
    }
}
