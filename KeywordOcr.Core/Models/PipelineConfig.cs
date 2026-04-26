namespace KeywordOcr.Core.Models;

/// <summary>
/// 파이프라인 실행 설정 (Python PipelineConfig 포팅)
/// </summary>
public class PipelineConfig
{
    // ── 입력 파일 ──────────────────────────────────────
    public string FilePath { get; set; } = "";
    public string ImgTag { get; set; } = "";
    public string TesseractPath { get; set; } = @"C:\Program Files\Tesseract-OCR";

    // ── AI 모델 ────────────────────────────────────────
    public string ModelKeyword { get; set; } = "claude-haiku-4-5-20251001";
    public string ModelLongtail { get; set; } = "claude-haiku-4-5-20251001";
    public string KeywordVersion { get; set; } = "2.0";

    // ── 키워드 글자수 ──────────────────────────────────
    public int MaxWords { get; set; } = 24;
    public int MaxLen { get; set; } = 140;
    public int MinLen { get; set; } = 90;
    public int ANameMin { get; set; } = 80;
    public int ANameMax { get; set; } = 100;
    public int BNameMin { get; set; } = 63;
    public int BNameMax { get; set; } = 98;
    public int ATagCount { get; set; } = 20;
    public int BTagCount { get; set; } = 14;

    // ── OCR ───────────────────────────────────────────
    public bool UseLocalOcr { get; set; } = true;
    public bool MergeOcrWithName { get; set; } = true;
    public bool KoreanOnly { get; set; } = true;
    public bool DropDigits { get; set; } = true;
    public int Psm { get; set; } = 11;
    public int Oem { get; set; } = 3;
    public int MaxImgs { get; set; } = 999;
    public int Threads { get; set; } = 6;
    public int MaxDepth { get; set; } = -1;
    public string LocalImgDir { get; set; } = "";
    public bool AllowFolderMatch { get; set; } = true;

    // ── 대표이미지 ─────────────────────────────────────
    public bool MakeListing { get; set; } = true;
    public int ListingSize { get; set; } = 1200;
    public int ListingPad { get; set; } = 20;
    public int ListingMax { get; set; } = 20;
    public string LogoPath { get; set; } = "";
    public string LogoPathB { get; set; } = "";
    public int LogoRatio { get; set; } = 14;
    public int LogoOpacity { get; set; } = 65;
    public string LogoPos { get; set; } = "tr";
    public bool UseAutoContrast { get; set; } = true;
    public bool UseSharpen { get; set; } = true;
    public bool UseSmallRotate { get; set; } = true;
    public float RotateZoom { get; set; } = 1.04f;
    public float UltraAngleDeg { get; set; } = 0.35f;
    public float UltraTranslatePx { get; set; } = 0.6f;
    public float UltraScalePct { get; set; } = 0.25f;
    public int TrimTol { get; set; } = 8;
    public int JpegQMin { get; set; } = 88;
    public int JpegQMax { get; set; } = 92;
    public bool DoFlipLr { get; set; } = true;

    // ── B마켓 ──────────────────────────────────────────
    public bool EnableBMarket { get; set; } = true;
    public string ImgTagB { get; set; } = "";

    // ── 실행 모드 ──────────────────────────────────────
    /// <summary>"full" | "images" | "analysis"</summary>
    public string Phase { get; set; } = "full";
    public string ExportRootOverride { get; set; } = "";
    public int ChunkSize { get; set; } = 10;
    public bool WriteToR { get; set; } = true;
    public bool Debug { get; set; } = true;

    // ── API 키 (파일에서 로드, 없으면 환경변수) ─────────
    public string AnthropicApiKey { get; set; } = "";
    public string OpenAiApiKey { get; set; } = "";
}
