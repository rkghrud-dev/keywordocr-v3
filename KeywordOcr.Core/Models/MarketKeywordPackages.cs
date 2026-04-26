namespace KeywordOcr.Core.Models;

/// <summary>
/// 마켓별 키워드 패키지 (Python MarketKeywordPackages 포팅)
/// </summary>
public class MarketKeywordPackages
{
    /// <summary>공백 구분 검색어 문자열 (Cafe24 검색어설정)</summary>
    public string SearchKeywords { get; set; } = "";

    /// <summary>쿠팡 태그 목록 (A마켓 최대 20개, B마켓 최대 14개)</summary>
    public List<string> CoupangTags { get; set; } = [];

    /// <summary>네이버 태그 목록 (A마켓 최대 10개, B마켓 최대 7개)</summary>
    public List<string> NaverTags { get; set; } = [];

    /// <summary>전체 후보 풀 (버킷 평탄화)</summary>
    public List<string> CandidatePool { get; set; } = [];

    /// <summary>LLM 원본 버킷 맵 (디버그용)</summary>
    public Dictionary<string, List<string>> BucketMap { get; set; } = [];
}
