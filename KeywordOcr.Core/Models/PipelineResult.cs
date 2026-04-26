namespace KeywordOcr.Core.Models;

/// <summary>
/// 상품 한 개의 파이프라인 처리 결과
/// </summary>
public class PipelineResult
{
    public int RowIndex { get; set; }
    public string GsCode { get; set; } = "";
    public string ProductName { get; set; } = "";

    // ── OCR ───────────────────────────────────────────
    public string OcrText { get; set; } = "";

    // ── 키워드 ─────────────────────────────────────────
    /// <summary>A마켓 검색어설정 (공백 구분)</summary>
    public string SearchKeywordsA { get; set; } = "";

    /// <summary>B마켓 검색어설정 (공백 구분)</summary>
    public string SearchKeywordsB { get; set; } = "";

    /// <summary>A마켓 쿠팡 태그</summary>
    public List<string> CoupangTagsA { get; set; } = [];

    /// <summary>B마켓 쿠팡 태그</summary>
    public List<string> CoupangTagsB { get; set; } = [];

    /// <summary>A마켓 네이버 태그</summary>
    public List<string> NaverTagsA { get; set; } = [];

    /// <summary>B마켓 네이버 태그</summary>
    public List<string> NaverTagsB { get; set; } = [];

    /// <summary>후보 풀 전체</summary>
    public List<string> CandidatePool { get; set; } = [];

    // ── 대표이미지 ─────────────────────────────────────
    /// <summary>생성된 대표이미지 경로 목록</summary>
    public List<string> ListingImagePaths { get; set; } = [];

    // ── 상태 ───────────────────────────────────────────
    public bool Success { get; set; } = true;
    public string? ErrorMessage { get; set; }

    /// <summary>원본 행 컬럼 (출력 Excel 전달용)</summary>
    public Dictionary<string, object?> RawColumns { get; set; } = [];
}
